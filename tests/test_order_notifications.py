import hashlib
import hmac
import json
import os
import sqlite3
import time
from datetime import datetime, timezone
from pathlib import Path

import pytest
import requests
from cryptography.fernet import Fernet
from flask import Flask
from flask_login import LoginManager, UserMixin, login_user
from PIL import Image

import fulfillment_worker
import fulfillment_common
import inv_migrations
import order_notification_api
import order_notification_email
import order_notification_service
from order_notification_api import (
    notification_super_admin_required,
    order_notification_bp,
    verify_event_request,
)
from order_notification_provider import (
    ProviderError,
    WeComBotProvider,
    encrypt_managed_webhook,
    resolve_target_webhook,
    validate_wecom_webhook,
)
from order_notification_renderer import render_order_cards
from order_notification_service import (
    NotificationPermanent,
    NotificationRetry,
    authoritative_snapshot,
    cleanup_expired_cards,
    create_job_for_order,
    create_test_send_job,
    enqueue_synced_orders,
    notification_summary,
    process_notification_job,
)


STORE = "https://example.test"


def _conn(path):
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(
        """
        CREATE TABLE settings (key TEXT PRIMARY KEY,value TEXT);
        CREATE TABLE warehouses (id INTEGER PRIMARY KEY,name TEXT,country TEXT);
        CREATE TABLE sites (
            id INTEGER PRIMARY KEY,url TEXT UNIQUE,manager TEXT,country TEXT,
            cod_on_hold_is_shipped INTEGER DEFAULT 0,
            consumer_key TEXT DEFAULT 'ck_test',consumer_secret TEXT DEFAULT 'cs_test'
        );
        CREATE TABLE orders (
            id TEXT PRIMARY KEY,woo_id INTEGER,number TEXT,version TEXT,status TEXT,
            currency TEXT,total TEXT,payment_method TEXT,payment_method_title TEXT,
            set_paid INTEGER,date_created TEXT,date_modified TEXT,updated_at TEXT,
            source TEXT,warehouse_id INTEGER,billing TEXT,shipping TEXT,
            shipping_lines TEXT,line_items TEXT,customer_note TEXT
        );
        CREATE TABLE oms_integration_jobs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,job_type TEXT,aggregate_type TEXT,
            aggregate_id TEXT,idempotency_key TEXT UNIQUE,payload_json TEXT,payload_hash TEXT,
            status TEXT,attempts INTEGER DEFAULT 0,max_attempts INTEGER DEFAULT 10,
            available_at TEXT,locked_at TEXT,locked_by TEXT,lease_expires_at TEXT,
            last_error_code TEXT,last_error TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,completed_at TEXT
        );
        """
    )
    inv_migrations.up_013(conn)
    inv_migrations.up_014(conn)
    inv_migrations.up_015(conn)
    inv_migrations.up_016(conn)
    conn.execute("INSERT INTO warehouses VALUES (1,'波兰主仓','PL')")
    conn.execute(
        "INSERT INTO sites (id,url,manager,country,cod_on_hold_is_shipped) VALUES (1,?,'Michael','PL',1)",
        (STORE,),
    )
    items = [
        {
            "id": 10,
            "sku": "HF-SPACE-1",
            "name": "Hifancy Space 50000 Puffs – Żurawina 冰爽",
            "quantity": 4,
            "meta_data": [{"key": "flavour", "value": "Borówka / 蓝莓"}],
        }
    ]
    conn.execute(
        """INSERT INTO orders
           (id,woo_id,number,version,status,currency,total,payment_method,payment_method_title,
            set_paid,date_created,date_modified,updated_at,source,warehouse_id,billing,shipping,
            shipping_lines,line_items,customer_note)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            "1-1465", 1465, "1465", "11.0.1", "processing", "PLN", "313.42",
            "cod", "Cash on delivery", 0, "2026-08-13T03:44:30",
            "2026-08-13T03:45:00", "2026-08-13T03:45:01", STORE, 1,
            json.dumps({"first_name": "Jan", "last_name": "Kowalski", "phone": "+48123456123", "email": "jan@example.test"}),
            json.dumps({"city": "Warszawa", "postcode": "00-001", "address_2": "WAW01A"}),
            json.dumps([{"method_title": "InPost Paczkomat"}]),
            json.dumps(items, ensure_ascii=False),
            "Proszę ostrożnie · 请轻放",
        ),
    )
    conn.execute("UPDATE settings SET value='1' WHERE key='order_notification_enabled'")
    conn.execute("UPDATE settings SET value='0' WHERE key='order_notification_debounce_seconds'")
    conn.execute(
        """INSERT INTO settings (key,value) VALUES ('order_notification_render_source','system_card')
           ON CONFLICT(key) DO UPDATE SET value=excluded.value"""
    )
    conn.commit()
    return conn


@pytest.fixture
def db(tmp_path):
    conn = _conn(tmp_path / "notifications.db")
    yield conn
    conn.close()


def _target(conn, channel="FAKE", *, environment="test", secret_ref=None, rate=15):
    conn.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,secret_ref,store_id,warehouse_id,environment,rate_limit_per_minute)
           VALUES ('target-1','测试订单群',?,?,?,?,?,?)""",
        (channel, secret_ref, STORE, 1, environment, rate),
    )
    conn.commit()


def _create(conn, event="ORDER_READY", event_id="evt-1"):
    return create_job_for_order(
        conn,
        "1-1465",
        event_id=event_id,
        requested_event=event,
    )


def test_migration_dark_launch_and_history_safe_down(tmp_path):
    conn = _conn(tmp_path / "migration.db")
    columns = {
        row[1] for row in conn.execute("PRAGMA table_info(notification_targets)")
    }
    assert {
        "manager_scope", "manager_names_json", "secret_ciphertext",
        "webhook_fingerprint", "deleted_at", "copy_to_fallback", "country_code",
    }.issubset(columns)
    inv_migrations.down_016(conn)
    assert "country_code" not in {
        row[1] for row in conn.execute("PRAGMA table_info(notification_targets)")
    }
    inv_migrations.down_015(conn)
    assert "copy_to_fallback" not in {
        row[1] for row in conn.execute("PRAGMA table_info(notification_targets)")
    }
    inv_migrations.down_014(conn)
    rolled_back_columns = {
        row[1] for row in conn.execute("PRAGMA table_info(notification_targets)")
    }
    assert "manager_scope" not in rolled_back_columns
    assert "secret_ciphertext" not in rolled_back_columns
    inv_migrations.up_014(conn)
    inv_migrations.up_015(conn)
    inv_migrations.up_016(conn)
    flags = dict(
        conn.execute(
            "SELECT key,value FROM settings WHERE key LIKE 'order_notification_%enabled'"
        ).fetchall()
    )
    assert flags == {
        "order_notification_enabled": "1",  # test fixture explicitly enabled it
        "order_notification_send_enabled": "0",
        "order_notification_test_send_enabled": "0",
    }
    conn.execute(
        "INSERT INTO notification_targets (id,name,channel_type) VALUES ('x','test','FAKE')"
    )
    conn.commit()
    with pytest.raises(RuntimeError, match="拒绝删除"):
        inv_migrations.down_013(conn)
    assert conn.execute(
        "SELECT 1 FROM sqlite_master WHERE name='notification_targets'"
    ).fetchone()
    conn.close()


def test_copy_to_fallback_migration_refuses_enabled_routes(tmp_path):
    conn = _conn(tmp_path / "copy-migration.db")
    conn.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,copy_to_fallback)
           VALUES ('copying','负责人群','FAKE',1)"""
    )
    conn.commit()
    with pytest.raises(RuntimeError, match="拒绝删除字段"):
        inv_migrations.down_015(conn)
    assert "copy_to_fallback" in {
        row[1] for row in conn.execute("PRAGMA table_info(notification_targets)")
    }
    conn.close()


def test_country_route_migration_refuses_used_country(tmp_path):
    conn = _conn(tmp_path / "country-migration.db")
    conn.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,country_code)
           VALUES ('poland','波兰群','FAKE','PL')"""
    )
    conn.commit()
    with pytest.raises(RuntimeError, match="拒绝删除字段"):
        inv_migrations.down_016(conn)
    assert "country_code" in {
        row[1] for row in conn.execute("PRAGMA table_info(notification_targets)")
    }
    conn.close()


def test_manager_group_route_covers_multiple_sites_and_site_route_overrides(db):
    second_store = "https://second-manager-site.test"
    other_store = "https://other-manager-site.test"
    db.execute(
        "INSERT INTO sites (id,url,manager,country) VALUES (2,?,'Michael','CZ')",
        (second_store,),
    )
    db.execute(
        "INSERT INTO sites (id,url,manager,country) VALUES (3,?,'Alice','AU')",
        (other_store,),
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,manager_scope,manager_names_json,environment)
           VALUES ('manager-team','负责人联合群','FAKE','selected','["Alice","Michael"]','test')"""
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,store_id,environment)
           VALUES ('site-override','站点专属群','FAKE',?,'test')""",
        (STORE,),
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,manager_scope,manager_names_json,environment)
           VALUES ('all-managers','全部负责人兜底群','FAKE','all','[]','test')"""
    )
    db.commit()

    first_snapshot = authoritative_snapshot(db, "1-1465")
    target, error = order_notification_service.resolve_target(
        db, first_snapshot, environment="test"
    )
    assert error is None
    assert target["id"] == "site-override"

    second_snapshot = {
        **first_snapshot,
        "store_id": second_store,
        "site_manager": "Michael",
    }
    target, error = order_notification_service.resolve_target(
        db, second_snapshot, environment="test"
    )
    assert error is None
    assert target["id"] == "manager-team"
    assert order_notification_service.target_matches_snapshot(
        target, second_snapshot
    )
    db.execute("UPDATE orders SET source=? WHERE id='1-1465'", (second_store,))
    db.commit()
    queued = create_job_for_order(
        db,
        "1-1465",
        event_id="manager-route-event",
        requested_event="ORDER_READY",
        target_environment="test",
    )
    assert queued["created"] is True
    assert queued["job"]["target_id"] == "manager-team"

    alice_snapshot = {
        **first_snapshot,
        "store_id": other_store,
        "site_manager": "Alice",
    }
    target, error = order_notification_service.resolve_target(
        db, alice_snapshot, environment="test"
    )
    assert error is None
    assert target["id"] == "manager-team"
    manager_target = dict(
        db.execute(
            "SELECT * FROM notification_targets WHERE id='manager-team'"
        ).fetchone()
    )
    assert order_notification_service.target_matches_snapshot(
        manager_target, alice_snapshot
    )
    unmatched_snapshot = {
        **first_snapshot,
        "store_id": "https://unassigned.test",
        "site_manager": "Bob",
    }
    target, error = order_notification_service.resolve_target(
        db, unmatched_snapshot, environment="test"
    )
    assert error is None
    assert target["id"] == "all-managers"


def test_country_route_covers_all_country_sites_between_site_and_manager_priority(db):
    second_poland = "https://second-poland.test"
    australia = "https://australia.test"
    db.execute(
        "INSERT INTO sites (id,url,manager,country) VALUES (2,?,'Michael','PL')",
        (second_poland,),
    )
    db.execute(
        "INSERT INTO sites (id,url,manager,country) VALUES (3,?,'Michael','AU')",
        (australia,),
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,store_id,environment)
           VALUES ('site-route','指定站点群','FAKE',?,'test')""",
        (STORE,),
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,country_code,environment)
           VALUES ('country-pl','波兰订单群','FAKE','PL','test')"""
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,manager_scope,manager_names_json,environment)
           VALUES ('manager-route','Michael 负责人群','FAKE','selected','["Michael"]','test')"""
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,environment)
           VALUES ('fallback-route','全部站点总群','FAKE','test')"""
    )
    db.commit()

    base = authoritative_snapshot(db, "1-1465")
    target, error = order_notification_service.resolve_target(
        db, base, environment="test"
    )
    assert error is None and target["id"] == "site-route"

    poland_snapshot = {
        **base,
        "store_id": second_poland,
        "site_country": "PL",
        "site_manager": "Michael",
    }
    target, error = order_notification_service.resolve_target(
        db, poland_snapshot, environment="test"
    )
    assert error is None and target["id"] == "country-pl"
    assert order_notification_service.target_matches_snapshot(target, poland_snapshot)

    australia_snapshot = {
        **base,
        "store_id": australia,
        "site_country": "AU",
        "site_manager": "Michael",
    }
    target, error = order_notification_service.resolve_target(
        db, australia_snapshot, environment="test"
    )
    assert error is None and target["id"] == "manager-route"

    unmatched = {
        **australia_snapshot,
        "store_id": "https://alice-australia.test",
        "site_manager": "Alice",
    }
    target, error = order_notification_service.resolve_target(
        db, unmatched, environment="test"
    )
    assert error is None and target["id"] == "fallback-route"

    db.execute("UPDATE orders SET source=? WHERE id='1-1465'", (second_poland,))
    db.commit()
    created = create_job_for_order(
        db,
        "1-1465",
        event_id="country-route-event",
        requested_event="ORDER_READY",
        target_environment="test",
    )
    assert created["created"] is True
    assert created["job"]["target_id"] == "country-pl"


def test_specific_route_can_copy_new_order_to_unique_fallback(db, tmp_path, monkeypatch):
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,manager_scope,manager_names_json,
            environment,copy_to_fallback)
           VALUES ('manager-primary','Michael 负责人群','FAKE','selected','["Michael"]',
                   'test',1)"""
    )
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,manager_scope,manager_names_json,environment)
           VALUES ('fallback-group','全部站点总群','FAKE','all','[]','test')"""
    )
    db.commit()

    created = create_job_for_order(
        db,
        "1-1465",
        event_id="copy-event",
        requested_event="ORDER_READY",
        target_environment="test",
    )

    assert created["created"] is True
    assert created["job"]["target_id"] == "manager-primary"
    assert created["fallback_copy_error"] is None
    assert created["fallback_copy_job"]["target_id"] == "fallback-group"
    assert [job["target_id"] for job in created["jobs"]] == [
        "manager-primary",
        "fallback-group",
    ]
    rows = db.execute(
        """SELECT target_id,idempotency_key,status
             FROM order_notification_jobs ORDER BY queue_job_id"""
    ).fetchall()
    assert [row["target_id"] for row in rows] == [
        "manager-primary",
        "fallback-group",
    ]
    assert len({row["idempotency_key"] for row in rows}) == 2
    assert db.execute(
        "SELECT COUNT(*) FROM oms_integration_jobs WHERE job_type='ORDER_NOTIFICATION'"
    ).fetchone()[0] == 2
    roles = {
        json.loads(row[0])["delivery_role"]
        for row in db.execute(
            """SELECT after_summary FROM notification_audit_logs
                 WHERE action='job_created' AND object_type='order_notification_job'"""
        ).fetchall()
    }
    assert roles == {"primary", "fallback_copy"}

    duplicate = create_job_for_order(
        db,
        "1-1465",
        event_id="copy-event",
        requested_event="ORDER_READY",
        target_environment="test",
    )
    assert duplicate["duplicate"] is True
    assert db.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 2

    monkeypatch.setenv("ORDER_NOTIFICATION_IMAGE_DIR", str(tmp_path / "copy-cards"))
    assert fulfillment_worker.run_one(db) is True
    assert fulfillment_worker.run_one(db) is True
    assert dict(
        db.execute(
            "SELECT status,COUNT(*) FROM order_notification_jobs GROUP BY status"
        ).fetchone()
    ) == {"status": "SENT", "COUNT(*)": 2}


def test_missing_fallback_never_blocks_primary_and_is_audited(db):
    db.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,manager_scope,manager_names_json,
            environment,copy_to_fallback)
           VALUES ('manager-primary','Michael 负责人群','FAKE','selected','["Michael"]',
                   'test',1)"""
    )
    db.commit()

    created = create_job_for_order(
        db,
        "1-1465",
        event_id="copy-missing-event",
        requested_event="ORDER_READY",
        target_environment="test",
    )

    assert created["created"] is True
    assert created["job"]["target_id"] == "manager-primary"
    assert created["fallback_copy_job"] is None
    assert created["fallback_copy_error"] == "fallback_route_missing"
    assert db.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 1
    audit = db.execute(
        """SELECT after_summary FROM notification_audit_logs
             WHERE action='fallback_copy_skipped'"""
    ).fetchone()
    assert json.loads(audit[0])["error"] == "fallback_route_missing"


def test_snapshot_is_minimized_and_uses_date_modified_not_wc_version(db):
    snap = authoritative_snapshot(db, "1-1465")
    assert snap["recipient"]["name_masked"] == "J** K****"
    assert snap["recipient"]["phone_masked"].endswith("123")
    assert "address_1" not in snap["recipient"]
    result = _create(db)
    assert result["created"] is True
    assert result["job"]["order_version"].startswith("2026-08-13T03:45:00:")
    assert "11.0.1" not in result["job"]["order_version"]


def test_duplicate_event_is_idempotent_and_worker_marks_sent(db, tmp_path, monkeypatch):
    _target(db)
    first = _create(db)
    duplicate = _create(db)
    assert duplicate["duplicate"] is True
    assert duplicate["job"]["id"] == first["job"]["id"]
    monkeypatch.setenv("ORDER_NOTIFICATION_IMAGE_DIR", str(tmp_path / "cards"))
    assert fulfillment_worker.run_one(db) is True
    job = db.execute(
        "SELECT status,image_sha256,image_bytes FROM order_notification_jobs"
    ).fetchone()
    queue = db.execute("SELECT status,attempts FROM oms_integration_jobs").fetchone()
    assert dict(job)["status"] == "SENT"
    assert len(job["image_sha256"]) == 64 and job["image_bytes"] > 1000
    assert dict(queue) == {"status": "succeeded", "attempts": 1}


def test_notification_jobs_use_durable_exponential_retry_window(db):
    _target(db)
    created = _create(db)
    queue = db.execute(
        "SELECT max_attempts FROM oms_integration_jobs WHERE id=?",
        (created["job"]["queue_job_id"],),
    ).fetchone()
    assert queue["max_attempts"] == 12
    assert order_notification_service._notification_retry_delay({"attempts": 1}) == 30
    assert order_notification_service._notification_retry_delay({"attempts": 2}) == 60
    assert order_notification_service._notification_retry_delay({"attempts": 7}) == 1800
    assert order_notification_service._notification_retry_delay({"attempts": 20}) == 1800


def test_expired_worker_lease_is_recovered_without_new_job(db):
    _target(db)
    created = _create(db)
    queue_id = created["job"]["queue_job_id"]
    first_claim = fulfillment_worker.claim_job(db)
    assert first_claim["id"] == queue_id
    db.execute(
        "UPDATE oms_integration_jobs SET lease_expires_at='2000-01-01T00:00:00+00:00' WHERE id=?",
        (queue_id,),
    )
    db.commit()
    recovered = fulfillment_worker.claim_job(db)
    assert recovered["id"] == queue_id
    assert db.execute("SELECT COUNT(*) FROM oms_integration_jobs").fetchone()[0] == 1


def test_feature_flag_blocks_wecom_without_network(db, tmp_path):
    _target(db, "WECOM_BOT", environment="production", secret_ref="env:PROD_WEBHOOK")
    created = _create(db)
    result = process_notification_job(
        db,
        {"aggregate_id": created["job"]["id"]},
        {},
        output_dir=str(tmp_path / "cards"),
    )
    assert result["blocked"] == "feature_flag_off"
    row = db.execute("SELECT status,last_error_code FROM order_notification_jobs").fetchone()
    assert dict(row) == {"status": "READY_PREVIEW", "last_error_code": "feature_flag_off"}


def test_automatic_sync_is_new_order_only_and_routes_to_production(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "automatic-new-order.db"
    setup = _conn(db_path)
    setup.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,store_id,warehouse_id,environment,enabled)
           VALUES ('test-route','测试群','WECOM_BOT',?,1,'test',1)""",
        (STORE,),
    )
    setup.execute(
        """INSERT INTO notification_targets
           (id,name,channel_type,store_id,warehouse_id,environment,enabled)
           VALUES ('production-route','生产群','FAKE',?,1,'production',1)""",
        (STORE,),
    )
    setup.execute(
        """INSERT INTO settings(key,value) VALUES
           ('order_notification_auto_watermarks_json',?)""",
        (json.dumps({STORE: {"max_woo_id": 1465, "activated_at": "test"}}),),
    )
    setup.commit()
    setup.close()

    def get_test_conn():
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    monkeypatch.setattr(fulfillment_common, "get_conn", get_test_conn)
    enqueue_synced_orders(
        [{"order_id": "1-1465", "status": "processing", "date_modified": "v1"}]
    )
    verify = get_test_conn()
    assert verify.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 0
    verify.execute(
        "UPDATE settings SET value=? WHERE key='order_notification_auto_watermarks_json'",
        (json.dumps({STORE: {"max_woo_id": 1464, "activated_at": "test"}}),),
    )
    verify.commit()
    verify.close()
    enqueue_synced_orders(
        [{"order_id": "1-1465", "status": "processing", "date_modified": "v1"}]
    )
    verify = get_test_conn()
    first = verify.execute(
        "SELECT event_type,target_id,status FROM order_notification_jobs"
    ).fetchone()
    assert dict(first) == {
        "event_type": "ORDER_READY",
        "target_id": "production-route",
        "status": "PENDING",
    }
    verify.execute(
        "UPDATE orders SET total='319.00',date_modified='2026-08-13T04:00:00' WHERE id='1-1465'"
    )
    verify.commit()
    verify.close()

    enqueue_synced_orders(
        [{"order_id": "1-1465", "status": "processing", "date_modified": "v2"}]
    )
    verify = get_test_conn()
    verify.execute(
        "UPDATE orders SET status='cancelled',date_modified='2026-08-13T04:05:00' WHERE id='1-1465'"
    )
    verify.commit()
    verify.close()
    enqueue_synced_orders(
        [{"order_id": "1-1465", "status": "cancelled", "date_modified": "v3"}]
    )

    verify = get_test_conn()
    assert verify.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 1
    assert verify.execute(
        "SELECT COUNT(*) FROM order_notification_jobs WHERE event_type!='ORDER_READY'"
    ).fetchone()[0] == 0
    verify.close()


def test_worker_rejects_non_test_job_routed_to_test_wecom(
    db, tmp_path, monkeypatch
):
    _target(
        db,
        "WECOM_BOT",
        environment="test",
        secret_ref="env:TEST_WECOM_WEBHOOK",
    )
    db.execute(
        "UPDATE settings SET value='1' WHERE key='order_notification_test_send_enabled'"
    )
    db.commit()
    monkeypatch.setenv(
        "TEST_WECOM_WEBHOOK",
        "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=isolated-test-key",
    )
    created = _create(db)
    session = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    with pytest.raises(NotificationPermanent) as blocked:
        process_notification_job(
            db,
            {"aggregate_id": created["job"]["id"]},
            {},
            session=session,
            output_dir=str(tmp_path / "blocked-test-route"),
        )
    assert blocked.value.code == "automatic_target_environment_invalid"
    assert session.calls == []
    assert dict(db.execute(
        "SELECT status,last_error_code FROM order_notification_jobs WHERE id=?",
        (created["job"]["id"],),
    ).fetchone()) == {
        "status": "DEAD_LETTER",
        "last_error_code": "automatic_target_environment_invalid",
    }


def test_manual_wechat_only_generates_downloadable_card(db, tmp_path):
    _target(db, "MANUAL_WECHAT")
    created = _create(db)
    result = process_notification_job(
        db, {"aggregate_id": created["job"]["id"]}, {}, output_dir=str(tmp_path / "cards")
    )
    assert result == {"manual_ready": True, "images": 1}
    assert db.execute("SELECT status FROM order_notification_jobs").fetchone()[0] == "READY_MANUAL"


def test_retention_cleanup_is_path_confined_and_preserves_audit(db, tmp_path):
    _target(db)
    created = _create(db)
    process_notification_job(
        db, {"aggregate_id": created["job"]["id"]}, {}, output_dir=str(tmp_path / "private")
    )
    outside = tmp_path / "must-stay.png"
    Image.new("RGB", (8, 8), "blue").save(outside)
    paths = json.loads(
        db.execute("SELECT image_paths_json FROM order_notification_jobs").fetchone()[0]
    )
    db.execute(
        "UPDATE order_notification_jobs SET image_paths_json=?,updated_at='2000-01-01 00:00:00'",
        (json.dumps(paths + [str(outside)]),),
    )
    db.commit()
    assert cleanup_expired_cards(db, tmp_path / "private") == 1
    assert not Path(paths[0]).exists()
    assert outside.exists()
    remaining = json.loads(
        db.execute("SELECT image_paths_json FROM order_notification_jobs").fetchone()[0]
    )
    assert remaining == [str(outside)]
    assert db.execute(
        "SELECT COUNT(*) FROM notification_audit_logs WHERE action='card_retention_cleanup'"
    ).fetchone()[0] == 1


def test_application_rate_limit_queues_instead_of_dropping(db, tmp_path):
    _target(db, rate=1)
    created = _create(db)
    job_id = created["job"]["id"]
    db.execute(
        """INSERT INTO order_notification_attempts
           (job_id,attempt_no,started_at,finished_at,result)
           VALUES (?,99,datetime('now'),datetime('now'),'SUCCESS')""",
        (job_id,),
    )
    db.commit()
    with pytest.raises(NotificationRetry) as error:
        process_notification_job(
            db, {"aggregate_id": job_id}, {}, output_dir=str(tmp_path / "cards")
        )
    assert error.value.code == "rate_limited"
    assert db.execute("SELECT status FROM order_notification_jobs").fetchone()[0] == "RETRY_WAIT"
    assert db.execute("SELECT COUNT(*) FROM oms_integration_jobs").fetchone()[0] == 1


def test_worker_dead_letter_synchronizes_notification_job_and_preserves_manual_review(db):
    _target(db)
    created = _create(db)
    notification_job_id = created["job"]["id"]
    db.execute(
        "UPDATE oms_integration_jobs SET status='running',attempts=12,max_attempts=12"
    )
    db.execute(
        "UPDATE order_notification_jobs SET status='RETRY_WAIT' WHERE id=?",
        (notification_job_id,),
    )
    db.commit()
    queue_job = dict(db.execute("SELECT * FROM oms_integration_jobs").fetchone())

    fulfillment_worker.dead_job(
        db,
        queue_job,
        RuntimeError("邮件日志仍未出现"),
        code="admin_new_order_email_not_found",
    )

    assert dict(
        db.execute(
            "SELECT status,last_error_code,last_error_summary FROM order_notification_jobs WHERE id=?",
            (notification_job_id,),
        ).fetchone()
    ) == {
        "status": "DEAD_LETTER",
        "last_error_code": "admin_new_order_email_not_found",
        "last_error_summary": "邮件日志仍未出现",
    }
    assert db.execute("SELECT status FROM oms_integration_jobs").fetchone()[0] == "dead_letter"
    audit = json.loads(
        db.execute(
            "SELECT after_summary FROM notification_audit_logs WHERE action='notification_queue_dead_letter'"
        ).fetchone()[0]
    )
    assert audit["domain_status_changed"] is True
    assert audit["attempts"] == 12
    assert audit["max_attempts"] == 12

    db.execute(
        "UPDATE order_notification_jobs SET status='MANUAL_REVIEW' WHERE id=?",
        (notification_job_id,),
    )
    db.execute("UPDATE oms_integration_jobs SET status='running'")
    db.commit()
    fulfillment_worker.dead_job(
        db,
        dict(db.execute("SELECT * FROM oms_integration_jobs").fetchone()),
        RuntimeError("发送结果未知"),
        code="delivery_unknown",
    )
    assert db.execute(
        "SELECT status FROM order_notification_jobs WHERE id=?", (notification_job_id,)
    ).fetchone()[0] == "MANUAL_REVIEW"


class _Response:
    def __init__(self, status, body):
        self.status_code = status
        self._body = body

    def json(self):
        return self._body


class _Session:
    def __init__(self, response):
        self.response = response
        self.calls = []

    def post(self, url, **kwargs):
        self.calls.append((url, kwargs))
        return self.response


class _SequenceSession:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def post(self, url, **kwargs):
        self.calls.append((url, kwargs))
        return self.responses.pop(0)


class _TimeoutSession:
    def post(self, url, **kwargs):
        raise requests.ReadTimeout("timeout")


def test_wecom_provider_payload_and_transient_error(tmp_path, monkeypatch):
    image = tmp_path / "x.png"
    Image.new("RGB", (32, 32), "red").save(image)
    url = "https://" + "qyapi.weixin.qq.com" + "/cgi-bin/webhook/send?" + "key=test-secret-value"
    monkeypatch.setenv("TEST_WECOM_WEBHOOK", url)
    success = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    result = WeComBotProvider(success).send_images(
        [str(image)], {"secret_ref": "env:TEST_WECOM_WEBHOOK"}
    )
    assert result["accepted"] is True
    payload = success.calls[0][1]["json"]
    assert payload["msgtype"] == "image"
    assert payload["image"]["md5"] == hashlib.md5(image.read_bytes()).hexdigest()
    assert "test-secret-value" not in json.dumps(payload)
    text_success = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    text_result = WeComBotProvider(text_success).send_text(
        "订单 #1465 图片推送延迟", {"secret_ref": "env:TEST_WECOM_WEBHOOK"}
    )
    assert text_result == {"accepted": True, "provider": "wecom", "messages": 1}
    assert text_success.calls[0][1]["json"] == {
        "msgtype": "text",
        "text": {"content": "订单 #1465 图片推送延迟"},
    }
    with pytest.raises(ProviderError) as too_large:
        WeComBotProvider(text_success).send_text(
            "测" * 700, {"secret_ref": "env:TEST_WECOM_WEBHOOK"}
        )
    assert too_large.value.code == "text_size_invalid"
    busy = _Session(_Response(429, {}))
    with pytest.raises(ProviderError) as error:
        WeComBotProvider(busy).send_images(
            [str(image)], {"secret_ref": "env:TEST_WECOM_WEBHOOK"}
        )
    assert error.value.retryable is True
    with pytest.raises(ProviderError):
        validate_wecom_webhook("https://example.com/cgi-bin/webhook/send?key=x")
    with pytest.raises(ProviderError):
        validate_wecom_webhook("https://qyapi.weixin.qq.com:444/cgi-bin/webhook/send?key=x")
    with pytest.raises(ProviderError):
        validate_wecom_webhook("https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=x&next=evil")


def test_managed_webhook_is_encrypted_and_resolved_only_server_side(monkeypatch):
    master_key = Fernet.generate_key().decode("ascii")
    webhook = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=managed-secret"
    monkeypatch.setenv("ORDER_NOTIFICATION_WEBHOOK_MASTER_KEY", master_key)

    ciphertext, fingerprint = encrypt_managed_webhook(webhook)

    assert webhook not in ciphertext
    assert "managed-secret" not in ciphertext
    assert len(fingerprint) == 12
    assert resolve_target_webhook(
        {"secret_ciphertext": ciphertext, "secret_ref": None}
    ) == webhook
    success = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    result = WeComBotProvider(success).send_text(
        "加密目标测试",
        {"secret_ciphertext": ciphertext, "secret_ref": None},
    )
    assert result["accepted"] is True
    assert success.calls[0][0] == webhook
    monkeypatch.setenv("ORDER_NOTIFICATION_WEBHOOK_MASTER_KEY", Fernet.generate_key().decode("ascii"))
    with pytest.raises(ProviderError) as error:
        resolve_target_webhook(
            {"secret_ciphertext": ciphertext, "secret_ref": None}
        )
    assert error.value.code == "webhook_decrypt_failed"


def test_retry_queues_one_privacy_minimized_alert_and_sends_it_idempotently(
    db, monkeypatch
):
    _target(
        db,
        "WECOM_BOT",
        environment="production",
        secret_ref="env:PROD_ALERT_WEBHOOK",
    )
    db.execute("UPDATE settings SET value='1' WHERE key='order_notification_send_enabled'")
    created = _create(db)
    notification_job_id = created["job"]["id"]
    db.execute(
        "UPDATE order_notification_jobs SET status='RETRY_WAIT' WHERE id=?",
        (notification_job_id,),
    )
    db.execute(
        "UPDATE oms_integration_jobs SET status='running',attempts=3,max_attempts=12"
    )
    db.commit()
    original_queue = dict(db.execute("SELECT * FROM oms_integration_jobs").fetchone())

    fulfillment_worker.retry_job(
        db,
        original_queue,
        RuntimeError("raw customer detail must not be copied"),
        delay_seconds=120,
        code="admin_new_order_email_not_found",
    )

    queues = [
        dict(row)
        for row in db.execute("SELECT * FROM oms_integration_jobs ORDER BY id").fetchall()
    ]
    assert len(queues) == 2
    alert_queue = queues[1]
    assert alert_queue["job_type"] == "ORDER_NOTIFICATION_ALERT"
    alert_payload = json.loads(alert_queue["payload_json"])
    serialized = json.dumps(alert_payload, ensure_ascii=False)
    assert alert_payload["phase"] == "delayed"
    assert alert_payload["attempts"] == 3
    assert "Jan" not in serialized
    assert "example.test" not in serialized
    assert "123456123" not in serialized
    assert "raw customer detail" not in serialized

    duplicate_attempt = dict(original_queue)
    duplicate_attempt["attempts"] = 4
    fulfillment_worker.retry_job(
        db,
        duplicate_attempt,
        RuntimeError("still waiting"),
        delay_seconds=240,
        code="admin_new_order_email_not_found",
    )
    assert db.execute("SELECT COUNT(*) FROM oms_integration_jobs").fetchone()[0] == 2
    assert db.execute(
        "SELECT COUNT(*) FROM notification_audit_logs WHERE action='notification_failure_alert_queued'"
    ).fetchone()[0] == 1

    monkeypatch.setenv(
        "PROD_ALERT_WEBHOOK",
        "https://" + "qyapi.weixin.qq.com" + "/cgi-bin/webhook/send?" + "key=alert-test-key",
    )
    session = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    result = order_notification_service.process_notification_alert(
        db,
        alert_queue,
        alert_payload,
        session=session,
    )
    assert result["accepted"] is True
    content = session.calls[0][1]["json"]["text"]["content"]
    assert "example.test" in content
    assert "#1465" in content
    assert "仍在自动重试（3/12）" in content
    assert "Jan" not in content
    assert "jan@example.test" not in content
    assert "+48123456123" not in content
    assert "Borówka" not in content

    repeated_session = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    repeated = order_notification_service.process_notification_alert(
        db,
        alert_queue,
        alert_payload,
        session=repeated_session,
    )
    assert repeated == {"noop": "already_sent", "phase": "delayed"}
    assert repeated_session.calls == []


def test_recovered_image_skips_stale_delayed_alert(db, monkeypatch):
    _target(
        db,
        "WECOM_BOT",
        environment="production",
        secret_ref="env:PROD_ALERT_WEBHOOK",
    )
    db.execute("UPDATE settings SET value='1' WHERE key='order_notification_send_enabled'")
    created = _create(db)
    queue = dict(db.execute("SELECT * FROM oms_integration_jobs").fetchone())
    queue["attempts"] = 3
    queued = order_notification_service.enqueue_notification_failure_alert(
        db,
        queue,
        phase="delayed",
        error_code="admin_new_order_email_not_found",
    )
    db.execute(
        "UPDATE order_notification_jobs SET status='SENT' WHERE id=?",
        (created["job"]["id"],),
    )
    db.commit()
    alert_queue = dict(
        db.execute("SELECT * FROM oms_integration_jobs WHERE id=?", (queued["queue_job_id"],)).fetchone()
    )
    session = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    result = order_notification_service.process_notification_alert(
        db,
        alert_queue,
        json.loads(alert_queue["payload_json"]),
        session=session,
    )
    assert result == {"skipped": "image_recovered", "phase": "delayed"}
    assert session.calls == []


def test_dead_notification_queues_one_final_wecom_alert(db):
    _target(
        db,
        "WECOM_BOT",
        environment="production",
        secret_ref="env:PROD_ALERT_WEBHOOK",
    )
    db.execute("UPDATE settings SET value='1' WHERE key='order_notification_send_enabled'")
    created = _create(db)
    notification_job_id = created["job"]["id"]
    db.execute(
        "UPDATE order_notification_jobs SET status='RETRY_WAIT' WHERE id=?",
        (notification_job_id,),
    )
    db.execute(
        "UPDATE oms_integration_jobs SET status='running',attempts=12,max_attempts=12"
    )
    db.commit()
    queue = dict(db.execute("SELECT * FROM oms_integration_jobs").fetchone())

    fulfillment_worker.dead_job(
        db,
        queue,
        RuntimeError("email unavailable"),
        code="admin_new_order_email_not_found",
    )

    alert = db.execute(
        "SELECT * FROM oms_integration_jobs WHERE job_type='ORDER_NOTIFICATION_ALERT'"
    ).fetchone()
    assert alert is not None
    assert json.loads(alert["payload_json"])["phase"] == "final"
    assert db.execute(
        "SELECT COUNT(*) FROM notification_audit_logs WHERE action='notification_failure_alert_queued' AND request_id='final'"
    ).fetchone()[0] == 1


def test_invalid_webhook_dead_letters_and_unknown_result_needs_manual_review(db, tmp_path, monkeypatch):
    _target(db, "WECOM_BOT", environment="production", secret_ref="env:PROD_BAD_WEBHOOK")
    db.execute("UPDATE settings SET value='1' WHERE key='order_notification_send_enabled'")
    db.commit()
    monkeypatch.setenv("PROD_BAD_WEBHOOK", "https://example.com/not-wecom?key=x")
    created = _create(db)
    with pytest.raises(NotificationPermanent) as invalid:
        process_notification_job(
            db, {"aggregate_id": created["job"]["id"]}, {}, output_dir=str(tmp_path / "bad")
        )
    assert invalid.value.code == "webhook_invalid"
    assert db.execute("SELECT status FROM order_notification_jobs").fetchone()[0] == "DEAD_LETTER"

    db.execute("DELETE FROM order_notification_attempts")
    db.execute("DELETE FROM order_notification_jobs")
    db.execute("DELETE FROM oms_integration_jobs")
    db.commit()
    monkeypatch.setenv(
        "PROD_BAD_WEBHOOK",
        "https://" + "qyapi.weixin.qq.com" + "/cgi-bin/webhook/send?" + "key=not-a-real-test-key",
    )
    created = _create(db, event_id="evt-timeout")
    with pytest.raises(NotificationPermanent) as unknown:
        process_notification_job(
            db,
            {"aggregate_id": created["job"]["id"]},
            {},
            session=_TimeoutSession(),
            output_dir=str(tmp_path / "timeout"),
        )
    assert unknown.value.code == "delivery_unknown"
    assert db.execute("SELECT status FROM order_notification_jobs").fetchone()[0] == "MANUAL_REVIEW"


def test_partial_multipage_retry_reuses_bytes_and_does_not_duplicate_page(db, tmp_path, monkeypatch):
    _target(db, "WECOM_BOT", environment="production", secret_ref="env:PROD_WECOM_WEBHOOK")
    db.execute("UPDATE settings SET value='1' WHERE key='order_notification_send_enabled'")
    items = json.loads(db.execute("SELECT line_items FROM orders WHERE id='1-1465'").fetchone()[0])
    db.execute(
        "UPDATE orders SET line_items=? WHERE id='1-1465'",
        (json.dumps(items * 8, ensure_ascii=False),),
    )
    db.commit()
    monkeypatch.setenv(
        "PROD_WECOM_WEBHOOK",
        "https://" + "qyapi.weixin.qq.com" + "/cgi-bin/webhook/send?" + "key=isolated-test-key",
    )
    created = _create(db)
    first_session = _SequenceSession(
        [_Response(200, {"errcode": 0, "errmsg": "ok"}), _Response(500, {})]
    )
    with pytest.raises(NotificationRetry):
        process_notification_job(
            db,
            {"aggregate_id": created["job"]["id"]},
            {},
            session=first_session,
            output_dir=str(tmp_path / "cards"),
        )
    row = db.execute(
        "SELECT image_paths_json,image_sha256,sent_pages_json FROM order_notification_jobs"
    ).fetchone()
    paths = json.loads(row["image_paths_json"])
    before = [Path(path).read_bytes() for path in paths]
    assert json.loads(row["sent_pages_json"]) == [1]

    retry_session = _SequenceSession([_Response(200, {"errcode": 0, "errmsg": "ok"})])
    result = process_notification_job(
        db,
        {"aggregate_id": created["job"]["id"]},
        {},
        session=retry_session,
        output_dir=str(tmp_path / "ignored-on-retry"),
    )
    after = [Path(path).read_bytes() for path in paths]
    assert result["accepted"] is True
    assert len(retry_session.calls) == 1
    assert before == after
    assert db.execute("SELECT status FROM order_notification_jobs").fetchone()[0] == "SENT"


def test_hmac_timestamp_and_replay_window(monkeypatch):
    monkeypatch.setenv("ORDER_NOTIFICATION_EVENT_SECRET_REF", "TEST_EVENT_SECRET")
    monkeypatch.setenv("TEST_EVENT_SECRET", "super-secret")
    raw = b'{"event_id":"evt-1"}'
    now = 1_786_592_700
    signature = "sha256=" + hmac.new(
        b"super-secret", str(now).encode() + b"." + raw, hashlib.sha256
    ).hexdigest()
    assert verify_event_request(raw, str(now), signature, now=now) == (True, "ok")
    assert verify_event_request(raw + b" ", str(now), signature, now=now)[1] == "signature_invalid"
    assert verify_event_request(raw, str(now - 301), signature, now=now)[1] == "timestamp_expired"


def test_signed_event_endpoint_queues_once_and_replays_safely(tmp_path, monkeypatch):
    db_path = tmp_path / "event-api.db"
    setup = _conn(db_path)
    _target(setup, environment="production")
    setup.execute(
        """INSERT INTO settings(key,value) VALUES
           ('order_notification_auto_watermarks_json',?)""",
        (json.dumps({STORE: {"max_woo_id": 1464, "activated_at": "test"}}),),
    )
    setup.commit()
    setup.close()

    def get_test_conn():
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    monkeypatch.setattr(order_notification_api, "get_conn", get_test_conn)
    monkeypatch.setenv("ORDER_NOTIFICATION_EVENT_SECRET_REF", "TEST_EVENT_SECRET")
    monkeypatch.setenv("TEST_EVENT_SECRET", "isolated-event-secret")
    app = Flask(__name__)
    app.register_blueprint(order_notification_bp)
    payload = {
        "event_id": "evt-api-1",
        "event_type": "ORDER_READY",
        "occurred_at": datetime.now(timezone.utc).isoformat(),
        "store_id": STORE,
        "order_id": "1-1465",
        "source": "hongkong-order-system",
    }
    raw = json.dumps(payload, separators=(",", ":")).encode()
    stamp = str(int(time.time()))
    signature = "sha256=" + hmac.new(
        b"isolated-event-secret", stamp.encode() + b"." + raw, hashlib.sha256
    ).hexdigest()
    headers = {
        "Content-Type": "application/json",
        "X-Webhook-Timestamp": stamp,
        "X-Webhook-Signature": signature,
    }
    client = app.test_client()
    first = client.post("/api/v1/order-notification-events", data=raw, headers=headers)
    duplicate = client.post("/api/v1/order-notification-events", data=raw, headers=headers)
    assert first.status_code == 202 and first.get_json()["queued"] is True
    assert duplicate.status_code == 202 and duplicate.get_json()["duplicate"] is True

    disabled_payload = dict(
        payload,
        event_id="evt-api-update",
        event_type="ORDER_UPDATED",
    )
    disabled_raw = json.dumps(disabled_payload, separators=(",", ":")).encode()
    disabled_signature = "sha256=" + hmac.new(
        b"isolated-event-secret",
        stamp.encode() + b"." + disabled_raw,
        hashlib.sha256,
    ).hexdigest()
    disabled = client.post(
        "/api/v1/order-notification-events",
        data=disabled_raw,
        headers={**headers, "X-Webhook-Signature": disabled_signature},
    )
    assert disabled.status_code == 400
    assert disabled.get_json()["error"] == "event_type_not_enabled"
    verify = get_test_conn()
    assert verify.execute("SELECT COUNT(*) FROM order_notification_event_inbox").fetchone()[0] == 1
    assert verify.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 1
    verify.close()


def test_notification_ui_permission_is_builtin_super_admin_only():
    app = Flask(__name__)
    app.secret_key = "test-only"
    manager = LoginManager(app)

    class AuthUser(UserMixin):
        def __init__(self, user_id, username):
            self.id = user_id
            self.username = username

    users = {
        "super": AuthUser("1", "admin"),
        "role-admin": AuthUser("2", "operations-admin"),
    }

    @manager.user_loader
    def load_user(user_id):
        return next((user for user in users.values() if user.id == user_id), None)

    @app.route("/test-login/<name>")
    def test_login(name):
        login_user(users[name])
        return "ok"

    @app.route("/protected-notifications")
    @notification_super_admin_required
    def protected_notifications():
        return "allowed"

    client = app.test_client()
    client.get("/test-login/role-admin")
    denied = client.get("/protected-notifications")
    assert denied.status_code == 403
    assert denied.get_json()["error"] == "订单群通知仅超级管理员可用"
    client.get("/test-login/super")
    assert client.get("/protected-notifications").status_code == 200


def _notification_api_test_client(db_path, monkeypatch):
    def get_test_conn():
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    monkeypatch.setattr(order_notification_api, "get_conn", get_test_conn)
    monkeypatch.setattr(order_notification_api, "_view_sources", lambda: None)
    monkeypatch.setattr(order_notification_api, "_can_edit_order", lambda conn, order_id: True)
    app = Flask(__name__)
    app.secret_key = "notification-console-test-only"
    manager = LoginManager(app)

    class SuperUser(UserMixin):
        id = "1"
        username = "admin"

    user = SuperUser()

    @manager.user_loader
    def load_user(user_id):
        return user if user_id == user.id else None

    @app.route("/test-login")
    def test_login():
        login_user(user)
        return "ok"

    app.register_blueprint(order_notification_bp)
    client = app.test_client()
    client.get("/test-login")
    return client, get_test_conn


def test_super_admin_console_updates_preview_only_configuration_and_locks_sending(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-config.db"
    setup = _conn(db_path)
    setup.close()
    client, get_test_conn = _notification_api_test_client(db_path, monkeypatch)
    monkeypatch.delenv("ORDER_NOTIFICATION_WEBHOOK_MASTER_KEY", raising=False)

    loaded = client.get("/api/order-notifications/config")
    assert loaded.status_code == 200
    assert loaded.get_json()["sending_locked"] is True
    assert loaded.get_json()["managed_webhook_ready"] is False
    assert loaded.get_json()["managers"] == [
        {"name": "Michael", "site_count": 1}
    ]
    assert loaded.get_json()["countries"] == [
        {"code": "PL", "site_count": 1}
    ]
    locked = client.post(
        "/api/order-notifications/config",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"enabled": True, "production_send_enabled": True},
    )
    assert locked.status_code == 403
    assert locked.get_json()["error"] == "sending_locked"

    payload = loaded.get_json()
    saved = client.post(
        "/api/order-notifications/config",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "enabled": False,
            "debounce_seconds": 30,
            "retention_days": 21,
            "policy": payload["policy"],
        },
    )
    assert saved.status_code == 200
    assert saved.get_json()["enabled"] is False
    verify = get_test_conn()
    assert dict(
        verify.execute(
            """SELECT key,value FROM settings WHERE key IN
               ('order_notification_enabled','order_notification_send_enabled',
                'order_notification_test_send_enabled')"""
        ).fetchall()
    ) == {
        "order_notification_enabled": "0",
        "order_notification_send_enabled": "0",
        "order_notification_test_send_enabled": "0",
    }
    assert verify.execute(
        "SELECT COUNT(*) FROM notification_audit_logs WHERE action='configuration_updated'"
    ).fetchone()[0] == 1
    verify.close()


def test_super_admin_dashboard_exposes_retry_health(db, tmp_path, monkeypatch):
    _target(db)
    created = _create(db)
    db.execute(
        "UPDATE order_notification_jobs SET status='RETRY_WAIT' WHERE id=?",
        (created["job"]["id"],),
    )
    db.execute(
        "UPDATE oms_integration_jobs SET status='retry',attempts=3,max_attempts=12"
    )
    db.commit()
    db_path = tmp_path / "dashboard-health.db"
    backup = sqlite3.connect(db_path)
    db.backup(backup)
    backup.close()
    client, _ = _notification_api_test_client(db_path, monkeypatch)
    monkeypatch.setattr(
        order_notification_api,
        "render_template",
        lambda _template, **context: context,
    )

    response = client.get("/order-notifications")

    assert response.status_code == 200
    body = response.get_json()
    assert body["control"]["exception_count"] == 1
    assert body["control"]["retry_count"] == 1
    assert body["control"]["queue_count"] == 1
    assert body["jobs"][0]["queue_status"] == "retry"
    assert body["jobs"][0]["queue_attempts"] == 3
    assert body["jobs"][0]["queue_max_attempts"] == 12


def test_preview_order_search_filters_site_and_accepts_order_number(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-order-search.db"
    setup = _conn(db_path)
    second_store = "https://second-store.test"
    setup.execute(
        "INSERT INTO sites (id,url,manager,country) VALUES (2,?,'Michael','CZ')",
        (second_store,),
    )
    setup.execute(
        """INSERT INTO orders
           (id,woo_id,number,status,date_modified,updated_at,source,billing,shipping,
            shipping_lines,line_items)
           VALUES ('2-9001',9001,'9001','processing','2026-08-14T08:00:00',
                   '2026-08-14T08:00:01',?,'{}','{}','[]','[]')""",
        (second_store,),
    )
    setup.commit()
    setup.close()
    client, _ = _notification_api_test_client(db_path, monkeypatch)

    filtered = client.get(
        "/api/order-notifications/orders",
        query_string={"site": second_store},
    )
    assert filtered.status_code == 200
    assert [item["id"] for item in filtered.get_json()["orders"]] == ["2-9001"]
    assert filtered.get_json()["status_filter"] == "new"

    by_number = client.get(
        "/api/order-notifications/orders",
        query_string={"q": "#9001"},
    )
    assert by_number.status_code == 200
    data = by_number.get_json()
    assert data["query"] == "9001"
    assert data["orders"][0] == {
        "id": "2-9001",
        "number": "9001",
        "status": "processing",
        "payment_method": None,
        "source": second_store,
        "warehouse_id": None,
        "date_modified": "2026-08-14T08:00:00",
    }
    assert "billing" not in data["orders"][0]

    escaped_wildcard = client.get(
        "/api/order-notifications/orders",
        query_string={"q": "%"},
    )
    assert escaped_wildcard.status_code == 200
    assert escaped_wildcard.get_json()["orders"] == []

    too_long = client.get(
        "/api/order-notifications/orders",
        query_string={"q": "1" * 129},
    )
    assert too_long.status_code == 400
    assert too_long.get_json()["error"] == "order_search_query_too_long"

    invalid_site = client.get(
        "/api/order-notifications/orders",
        query_string={"site": "https://missing.test"},
    )
    assert invalid_site.status_code == 400
    assert invalid_site.get_json()["error"] == "order_search_site_invalid"

    invalid_status = client.get(
        "/api/order-notifications/orders",
        query_string={"status": "bad status!"},
    )
    assert invalid_status.status_code == 400
    assert invalid_status.get_json()["error"] == "order_search_status_invalid"


def test_preview_smart_new_filter_uses_payment_aware_pending_and_excludes_hold(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-smart-new-filter.db"
    setup = _conn(db_path)
    au_store = "https://australia-online.test"
    cz_store = "https://czech-cod-disabled.test"
    setup.execute(
        "INSERT INTO sites (id,url,manager,country,cod_on_hold_is_shipped) VALUES (2,?,'A','AU',0)",
        (au_store,),
    )
    setup.execute(
        "INSERT INTO sites (id,url,manager,country,cod_on_hold_is_shipped) VALUES (3,?,'C','CZ',0)",
        (cz_store,),
    )
    rows = [
        ("1-1466", 1466, "1466", "on-hold", "cod", "2026-08-14T09:01:00", STORE),
        ("2-9001", 9001, "9001", "processing", "custom_gateway", "2026-08-14T09:02:00", au_store),
        ("2-9002", 9002, "9002", "on-hold", "bacs", "2026-08-14T09:03:00", au_store),
        ("3-8001", 8001, "8001", "on-hold", "cod", "2026-08-14T09:04:00", cz_store),
        ("1-1467", 1467, "1467", "failed", "cod", "2026-08-14T09:05:00", STORE),
        ("1-1468", 1468, "1468", "pending", "cod", "2026-08-14T09:06:00", STORE),
    ]
    setup.executemany(
        """INSERT INTO orders
           (id,woo_id,number,status,payment_method,date_modified,updated_at,source,
            billing,shipping,shipping_lines,line_items)
           VALUES (?,?,?,?,?,?,?,?,'{}','{}','[]','[]')""",
        [row[:6] + (row[5], row[6]) for row in rows],
    )
    setup.commit()
    setup.close()
    client, _ = _notification_api_test_client(db_path, monkeypatch)

    smart = client.get(
        "/api/order-notifications/orders", query_string={"status": "new"}
    )
    assert smart.status_code == 200
    assert {item["id"] for item in smart.get_json()["orders"]} == {
        "1-1465",
        "1-1468",
        "2-9001",
    }

    au_smart = client.get(
        "/api/order-notifications/orders",
        query_string={"site": au_store, "status": "new"},
    )
    assert [item["id"] for item in au_smart.get_json()["orders"]] == ["2-9001"]

    all_hold = client.get(
        "/api/order-notifications/orders", query_string={"status": "on-hold"}
    )
    assert {item["id"] for item in all_hold.get_json()["orders"]} == {
        "1-1466",
        "2-9002",
        "3-8001",
    }


def test_safe_preview_renders_real_snapshot_without_job_queue_or_send(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-preview.db"
    setup = _conn(db_path)
    setup.execute("UPDATE settings SET value='0' WHERE key='order_notification_enabled'")
    setup.commit()
    setup.close()
    client, get_test_conn = _notification_api_test_client(db_path, monkeypatch)
    preview_root = tmp_path / "preview-private"
    monkeypatch.setenv("ORDER_NOTIFICATION_IMAGE_DIR", str(preview_root))

    response = client.post(
        "/api/order-notifications/preview",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"order_id": "1-1465", "event_type": "ORDER_READY", "source": "system_card"},
    )
    assert response.status_code == 200
    data = response.get_json()
    assert data["queued"] is False and data["sent"] is False
    assert data["order"] == {
        "id": "1-1465",
        "number": "1465",
        "status": "processing",
        "store": "example.test",
    }
    assert data["routing"]["manager_name"] == "Michael"
    assert data["images"] and data["images"][0]["data_url"].startswith(
        "data:image/png;base64,"
    )
    verify = get_test_conn()
    assert verify.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 0
    assert verify.execute(
        "SELECT COUNT(*) FROM oms_integration_jobs WHERE job_type='ORDER_NOTIFICATION'"
    ).fetchone()[0] == 0
    assert verify.execute(
        "SELECT COUNT(*) FROM notification_audit_logs WHERE action='preview_generated'"
    ).fetchone()[0] == 1
    assert dict(
        verify.execute(
            """SELECT key,value FROM settings WHERE key IN
               ('order_notification_send_enabled','order_notification_test_send_enabled')"""
        ).fetchall()
    ) == {
        "order_notification_send_enabled": "0",
        "order_notification_test_send_enabled": "0",
    }
    verify.close()
    assert list(preview_root.iterdir()) == []


def test_test_send_api_requires_live_preview_flag_and_is_idempotent(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-test-send.db"
    setup = _conn(db_path)
    setup.execute("UPDATE settings SET value='0' WHERE key='order_notification_enabled'")
    _target(
        setup,
        "WECOM_BOT",
        environment="test",
        secret_ref="env:TEST_WECOM_WEBHOOK",
    )
    setup.close()
    monkeypatch.setenv(
        "TEST_WECOM_WEBHOOK",
        "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=isolated-test-key",
    )
    client, get_test_conn = _notification_api_test_client(db_path, monkeypatch)
    monkeypatch.setenv("ORDER_NOTIFICATION_IMAGE_DIR", str(tmp_path / "preview"))

    preview = client.post(
        "/api/order-notifications/preview",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"order_id": "1-1465", "event_type": "ORDER_READY", "source": "system_card"},
    )
    assert preview.status_code == 200
    preview_id = preview.get_json()["preview_id"]
    payload = {
        "order_id": "1-1465",
        "target_id": "target-1",
        "preview_id": preview_id,
        "source": "system_card",
        "confirmed": True,
    }

    fabricated = dict(payload, preview_id="preview-" + "f" * 32)
    rejected = client.post(
        "/api/order-notifications/test-send",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json=fabricated,
    )
    assert rejected.status_code == 409
    assert rejected.get_json()["error"] == "preview_id_invalid"

    locked = client.post(
        "/api/order-notifications/test-send",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json=payload,
    )
    assert locked.status_code == 409
    assert locked.get_json()["error"] == "test_send_flag_off"

    conn = get_test_conn()
    conn.execute(
        "UPDATE settings SET value='1' WHERE key='order_notification_test_send_enabled'"
    )
    conn.commit()
    conn.close()
    queued = client.post(
        "/api/order-notifications/test-send",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json=payload,
    )
    assert queued.status_code == 202
    assert queued.get_json()["queued"] is True
    assert queued.get_json()["sent"] is False

    duplicate = client.post(
        "/api/order-notifications/test-send",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json=payload,
    )
    assert duplicate.status_code == 200
    assert duplicate.get_json()["duplicate"] is True
    assert duplicate.get_json()["job_id"] == queued.get_json()["job_id"]

    verify = get_test_conn()
    assert verify.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 1
    assert verify.execute(
        "SELECT COUNT(*) FROM oms_integration_jobs WHERE job_type='ORDER_NOTIFICATION'"
    ).fetchone()[0] == 1
    stored = json.loads(
        verify.execute("SELECT snapshot_json FROM order_notification_jobs").fetchone()[0]
    )
    assert stored["_notification_mode"] == "test_send"
    assert stored["_notification_render_source"] == "system_card"
    assert verify.execute(
        "SELECT COUNT(*) FROM notification_audit_logs WHERE action='test_send_requested'"
    ).fetchone()[0] == 1
    verify.close()


def test_test_send_worker_uses_reviewed_source_and_only_test_target(
    db, tmp_path, monkeypatch
):
    _target(
        db,
        "WECOM_BOT",
        environment="test",
        secret_ref="env:TEST_WECOM_WEBHOOK",
    )
    db.execute("UPDATE settings SET value='0' WHERE key='order_notification_enabled'")
    db.execute(
        "UPDATE settings SET value='1' WHERE key='order_notification_test_send_enabled'"
    )
    db.execute("UPDATE settings SET value='system_card' WHERE key='order_notification_render_source'")
    db.commit()
    monkeypatch.setenv(
        "TEST_WECOM_WEBHOOK",
        "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=isolated-test-key",
    )
    rendered = []

    def fake_render(_conn, order_id, output_dir, job_id):
        rendered.append((order_id, job_id))
        path = Path(output_dir) / f"{job_id}-email.png"
        path.parent.mkdir(parents=True, exist_ok=True)
        Image.new("RGB", (120, 240), "white").save(path, "PNG")
        raw = path.read_bytes()
        return ([{
            "page": 1,
            "path": str(path),
            "sha256": hashlib.sha256(raw).hexdigest(),
            "width": 120,
            "height": 240,
            "bytes": len(raw),
        }], {
            "log_id": 885,
            "html_sha256": "b" * 64,
            "images_inlined": 2,
            "images_removed": 0,
            "template_version": "woo-admin-email-v1",
        })

    monkeypatch.setattr(order_notification_service, "render_logged_admin_email", fake_render)
    created = create_test_send_job(
        db,
        "1-1465",
        target_id="target-1",
        preview_id="preview-" + "a" * 32,
        render_source="email",
        actor={"type": "user", "id": "1"},
    )
    session = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    result = process_notification_job(
        db,
        {"aggregate_id": created["job"]["id"]},
        {},
        session=session,
        output_dir=str(tmp_path / "test-send"),
    )
    assert result == {"accepted": True, "images": 1, "provider": "WECOM_BOT"}
    assert rendered == [("1-1465", created["job"]["id"])]
    assert len(session.calls) == 1
    assert db.execute(
        "SELECT status FROM order_notification_jobs WHERE id=?", (created["job"]["id"],)
    ).fetchone()[0] == "SENT"


def test_test_send_worker_blocks_target_changed_to_production(
    db, tmp_path, monkeypatch
):
    _target(
        db,
        "WECOM_BOT",
        environment="test",
        secret_ref="env:TEST_WECOM_WEBHOOK",
    )
    db.execute(
        "UPDATE settings SET value='1' WHERE key='order_notification_test_send_enabled'"
    )
    db.execute("UPDATE settings SET value='1' WHERE key='order_notification_send_enabled'")
    db.commit()
    monkeypatch.setenv(
        "TEST_WECOM_WEBHOOK",
        "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=isolated-test-key",
    )
    created = create_test_send_job(
        db,
        "1-1465",
        target_id="target-1",
        preview_id="preview-" + "c" * 32,
        render_source="system_card",
        actor={"type": "user", "id": "1"},
    )
    db.execute(
        "UPDATE notification_targets SET environment='production' WHERE id='target-1'"
    )
    db.commit()
    session = _Session(_Response(200, {"errcode": 0, "errmsg": "ok"}))
    with pytest.raises(NotificationPermanent) as blocked:
        process_notification_job(
            db,
            {"aggregate_id": created["job"]["id"]},
            {},
            session=session,
            output_dir=str(tmp_path / "target-changed"),
        )
    assert blocked.value.code == "test_target_changed"
    assert session.calls == []
    row = db.execute(
        "SELECT status,last_error_code FROM order_notification_jobs WHERE id=?",
        (created["job"]["id"],),
    ).fetchone()
    assert dict(row) == {
        "status": "DEAD_LETTER",
        "last_error_code": "test_target_changed",
    }


def test_target_console_hides_secret_value_and_rejects_ambiguous_route(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-target.db"
    setup = _conn(db_path)
    setup.close()
    client, _ = _notification_api_test_client(db_path, monkeypatch)
    monkeypatch.setenv(
        "TEST_WECOM_WEBHOOK",
        "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=never-return-this",
    )
    created = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "隔离测试群",
            "channel_type": "WECOM_BOT",
            "environment": "test",
            "secret_ref": "env:TEST_WECOM_WEBHOOK",
            "store_id": STORE,
            "warehouse_id": 1,
            "enabled": True,
        },
    )
    assert created.status_code == 201
    assert created.get_json()["secret_available"] is True
    serialized = json.dumps(created.get_json())
    assert "never-return-this" not in serialized
    assert "secret_ref" not in created.get_json()

    production = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "同路由生产群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "production",
            "store_id": STORE,
            "warehouse_id": 1,
            "enabled": True,
        },
    )
    assert production.status_code == 201

    duplicate = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "冲突目标",
            "channel_type": "MANUAL_WECHAT",
            "environment": "test",
            "store_id": STORE,
            "warehouse_id": 1,
            "enabled": True,
        },
    )
    assert duplicate.status_code == 409
    assert duplicate.get_json()["error"] == "route_ambiguous"


def test_frontend_managed_webhook_test_message_and_soft_delete(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "managed-webhook-console.db"
    setup = _conn(db_path)
    setup.execute(
        "INSERT INTO sites (id,url,manager,country) VALUES (2,?,'Alice','AU')",
        ("https://alice-managed.test",),
    )
    setup.commit()
    setup.close()
    client, get_test_conn = _notification_api_test_client(db_path, monkeypatch)
    webhook = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=front-end-secret"
    payload = {
        "name": "负责人联合群",
        "channel_type": "WECOM_BOT",
        "environment": "production",
        "webhook_url": webhook,
        "manager_scope": "selected",
        "manager_names": ["Michael", "Alice"],
        "enabled": False,
    }

    missing_key = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json=payload,
    )
    assert missing_key.status_code == 503
    assert missing_key.get_json()["error"] == "webhook_master_key_missing"

    monkeypatch.setenv(
        "ORDER_NOTIFICATION_WEBHOOK_MASTER_KEY",
        Fernet.generate_key().decode("ascii"),
    )
    created = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json=payload,
    )
    assert created.status_code == 201
    target = created.get_json()
    assert target["manager_names"] == ["Alice", "Michael"]
    assert target["secret_source"] == "managed"
    assert target["secret_available"] is True
    assert target["webhook_fingerprint"]
    assert "front-end-secret" not in json.dumps(target)
    assert "secret_ciphertext" not in target
    target_id = target["id"]
    listed = client.get("/api/order-notifications/targets").get_json()
    assert len(listed) == 1
    assert "front-end-secret" not in json.dumps(listed)
    assert "secret_ciphertext" not in listed[0]

    verify = get_test_conn()
    stored = verify.execute(
        """SELECT secret_ref,secret_ciphertext,webhook_fingerprint
             FROM notification_targets WHERE id=?""",
        (target_id,),
    ).fetchone()
    assert stored["secret_ref"] is None
    assert stored["secret_ciphertext"]
    assert "front-end-secret" not in stored["secret_ciphertext"]
    original_ciphertext = stored["secret_ciphertext"]
    verify.close()

    edited_payload = {**payload, "id": target_id, "name": "负责人联合群（已编辑）"}
    edited_payload.pop("webhook_url")
    edited = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json=edited_payload,
    )
    assert edited.status_code == 200
    assert edited.get_json()["name"] == "负责人联合群（已编辑）"
    verify = get_test_conn()
    assert verify.execute(
        "SELECT secret_ciphertext FROM notification_targets WHERE id=?",
        (target_id,),
    ).fetchone()[0] == original_ciphertext
    verify.close()

    captured = []

    class CaptureProvider:
        def send_text(self, content, target_row):
            captured.append((content, target_row["id"]))
            return {"accepted": True, "provider": "wecom", "messages": 1}

    monkeypatch.setattr(
        order_notification_api,
        "provider_for",
        lambda channel: CaptureProvider(),
    )
    sent = client.post(
        f"/api/order-notifications/targets/{target_id}/test",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"confirmed": True},
    )
    assert sent.status_code == 200
    assert sent.get_json()["sent"] is True
    assert len(captured) == 1
    assert "Alice、Michael" in captured[0][0]
    assert "订单 #" not in captured[0][0]
    assert "front-end-secret" not in captured[0][0]
    limited = client.post(
        f"/api/order-notifications/targets/{target_id}/test",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"confirmed": True},
    )
    assert limited.status_code == 429
    assert limited.get_json()["error"] == "test_message_rate_limited"

    deleted = client.delete(
        f"/api/order-notifications/targets/{target_id}",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"confirmed": True},
    )
    assert deleted.status_code == 200
    assert deleted.get_json() == {
        "id": target_id,
        "deleted": True,
        "secret_cleared": True,
    }
    assert client.get("/api/order-notifications/targets").get_json() == []
    verify = get_test_conn()
    removed = verify.execute(
        """SELECT enabled,secret_ref,secret_ciphertext,webhook_fingerprint,deleted_at
             FROM notification_targets WHERE id=?""",
        (target_id,),
    ).fetchone()
    assert removed["enabled"] == 0
    assert removed["secret_ref"] is None
    assert removed["secret_ciphertext"] is None
    assert removed["webhook_fingerprint"] is None
    assert removed["deleted_at"]
    actions = {
        row[0]
        for row in verify.execute(
            """SELECT action FROM notification_audit_logs
                 WHERE object_type='notification_target' AND object_id=?""",
            (target_id,),
        ).fetchall()
    }
    assert {"target_created", "target_test_started", "target_test_sent", "target_deleted"}.issubset(actions)
    verify.close()


def test_target_delete_refuses_to_orphan_active_notification_job(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "target-delete-active.db"
    setup = _conn(db_path)
    _target(setup)
    created = _create(setup)
    assert created["job"]["status"] == "PENDING"
    setup.close()
    client, _ = _notification_api_test_client(db_path, monkeypatch)

    blocked = client.delete(
        "/api/order-notifications/targets/target-1",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"confirmed": True},
    )

    assert blocked.status_code == 409
    assert blocked.get_json()["error"] == "target_has_active_jobs"
    assert blocked.get_json()["active_jobs"] == 1


def test_target_console_saves_manager_group_route_and_validates_assignment(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-manager-target.db"
    setup = _conn(db_path)
    setup.execute(
        "INSERT INTO sites (id,url,manager,country) VALUES (2,?,'Alice','AU')",
        ("https://alice-store.test",),
    )
    setup.commit()
    setup.close()
    client, _ = _notification_api_test_client(db_path, monkeypatch)

    created = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
                "name": "Michael 负责人群",
                "channel_type": "MANUAL_WECHAT",
                "environment": "test",
                "manager_scope": "selected",
                "manager_names": ["Michael"],
                "enabled": True,
            },
        )
    assert created.status_code == 201
    assert created.get_json()["manager_scope"] == "selected"
    assert created.get_json()["manager_names"] == ["Michael"]
    assert created.get_json()["store_id"] is None

    all_managers = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "全部负责人兜底群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "test",
            "manager_scope": "all",
            "manager_names": ["Michael", "Alice"],
            "enabled": True,
        },
    )
    assert all_managers.status_code == 201
    assert all_managers.get_json()["manager_scope"] == "all"
    assert all_managers.get_json()["manager_names"] == []

    alice = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "Alice 负责人群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "test",
            "manager_scope": "selected",
            "manager_names": ["Alice"],
            "enabled": True,
        },
    )
    assert alice.status_code == 201

    duplicate = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
                "name": "Michael 冲突群",
                "channel_type": "MANUAL_WECHAT",
                "environment": "test",
                "manager_scope": "selected",
                "manager_names": ["Michael"],
                "enabled": True,
        },
    )
    assert duplicate.status_code == 409
    assert duplicate.get_json()["error"] == "route_ambiguous"

    mismatch = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "错误归属",
            "channel_type": "MANUAL_WECHAT",
                "environment": "test",
                "store_id": STORE,
                "manager_scope": "selected",
                "manager_names": ["Alice"],
                "enabled": True,
        },
    )
    assert mismatch.status_code == 400
    assert mismatch.get_json()["error"] == "store_manager_mismatch"

    missing = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "不存在负责人",
                "channel_type": "MANUAL_WECHAT",
                "environment": "test",
                "manager_scope": "selected",
                "manager_names": ["Nobody"],
                "enabled": True,
        },
    )
    assert missing.status_code == 400
    assert missing.get_json()["error"] == "manager_invalid"

    country = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "波兰订单群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "test",
            "country_code": "pl",
            "manager_scope": "all",
            "enabled": True,
        },
    )
    assert country.status_code == 201
    assert country.get_json()["country_code"] == "PL"

    duplicate_country = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "重复波兰群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "test",
            "country_code": "PL",
            "manager_scope": "all",
            "enabled": True,
        },
    )
    assert duplicate_country.status_code == 409
    assert duplicate_country.get_json()["error"] == "route_ambiguous"

    invalid_country = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "不存在国家",
            "channel_type": "MANUAL_WECHAT",
            "environment": "test",
            "country_code": "DE",
            "manager_scope": "all",
            "enabled": True,
        },
    )
    assert invalid_country.status_code == 400
    assert invalid_country.get_json()["error"] == "country_invalid"

    country_mismatch = client.post(
        "/api/order-notifications/targets",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={
            "name": "站点国家冲突",
            "channel_type": "MANUAL_WECHAT",
            "environment": "test",
            "store_id": STORE,
            "country_code": "AU",
            "manager_scope": "all",
            "enabled": True,
        },
    )
    assert country_mismatch.status_code == 400
    assert country_mismatch.get_json()["error"] == "store_country_mismatch"


def test_target_console_validates_copy_fallback_and_protects_total_group(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "console-copy-fallback.db"
    setup = _conn(db_path)
    setup.close()
    client, _ = _notification_api_test_client(db_path, monkeypatch)
    headers = {"X-Requested-With": "XMLHttpRequest"}
    manager_payload = {
        "name": "Michael 负责人群",
        "channel_type": "MANUAL_WECHAT",
        "environment": "production",
        "manager_scope": "selected",
        "manager_names": ["Michael"],
        "copy_to_fallback": True,
        "enabled": True,
    }

    missing = client.post(
        "/api/order-notifications/targets", headers=headers, json=manager_payload
    )
    assert missing.status_code == 409
    assert missing.get_json()["error"] == "fallback_route_missing"

    self_copy = client.post(
        "/api/order-notifications/targets",
        headers=headers,
        json={
            "name": "错误总群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "production",
            "manager_scope": "all",
            "copy_to_fallback": True,
            "enabled": True,
        },
    )
    assert self_copy.status_code == 400
    assert self_copy.get_json()["error"] == "fallback_cannot_copy_to_itself"

    fallback = client.post(
        "/api/order-notifications/targets",
        headers=headers,
        json={
            "name": "全部站点总群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "production",
            "manager_scope": "all",
            "copy_to_fallback": False,
            "enabled": True,
        },
    )
    assert fallback.status_code == 201
    fallback_id = fallback.get_json()["id"]
    assert fallback.get_json()["copy_to_fallback"] is False

    manager = client.post(
        "/api/order-notifications/targets", headers=headers, json=manager_payload
    )
    assert manager.status_code == 201
    manager_id = manager.get_json()["id"]
    assert manager.get_json()["copy_to_fallback"] is True

    invalid_type = client.post(
        "/api/order-notifications/targets",
        headers=headers,
        json={**manager_payload, "id": manager_id, "copy_to_fallback": "yes"},
    )
    assert invalid_type.status_code == 400
    assert invalid_type.get_json()["error"] == "copy_to_fallback_invalid"

    blocked_disable = client.post(
        "/api/order-notifications/targets",
        headers=headers,
        json={
            "id": fallback_id,
            "name": "全部站点总群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "production",
            "manager_scope": "all",
            "copy_to_fallback": False,
            "enabled": False,
        },
    )
    assert blocked_disable.status_code == 409
    assert blocked_disable.get_json()["error"] == "fallback_target_in_use"
    assert blocked_disable.get_json()["dependent_targets"] == 1

    blocked_delete = client.delete(
        f"/api/order-notifications/targets/{fallback_id}",
        headers=headers,
        json={"confirmed": True},
    )
    assert blocked_delete.status_code == 409
    assert blocked_delete.get_json()["error"] == "fallback_target_in_use"

    manager_payload.update({"id": manager_id, "copy_to_fallback": False})
    disabled_copy = client.post(
        "/api/order-notifications/targets", headers=headers, json=manager_payload
    )
    assert disabled_copy.status_code == 200
    assert disabled_copy.get_json()["copy_to_fallback"] is False

    fallback_disabled = client.post(
        "/api/order-notifications/targets",
        headers=headers,
        json={
            "id": fallback_id,
            "name": "全部站点总群",
            "channel_type": "MANUAL_WECHAT",
            "environment": "production",
            "manager_scope": "all",
            "copy_to_fallback": False,
            "enabled": False,
        },
    )
    assert fallback_disabled.status_code == 200
    assert fallback_disabled.get_json()["enabled"] == 0


def test_update_and_cancel_cards_are_distinct_and_audited(db, tmp_path):
    _target(db)
    ready = _create(db)
    process_notification_job(db, {"aggregate_id": ready["job"]["id"]}, {}, output_dir=str(tmp_path))

    db.execute(
        "UPDATE orders SET total='319.00',date_modified='2026-08-13T04:00:00' WHERE id='1-1465'"
    )
    db.commit()
    updated = _create(db, "ORDER_UPDATED", "evt-2")
    assert updated["created"] is True
    update_card = process_notification_job(
        db, {"aggregate_id": updated["job"]["id"]}, {}, output_dir=str(tmp_path)
    )
    assert update_card["accepted"] is True

    db.execute(
        "UPDATE orders SET status='cancelled',date_modified='2026-08-13T04:05:00' WHERE id='1-1465'"
    )
    db.commit()
    cancelled = _create(db, "ORDER_CANCELLED", "evt-3")
    assert cancelled["created"] is True
    process_notification_job(
        db, {"aggregate_id": cancelled["job"]["id"]}, {}, output_dir=str(tmp_path)
    )
    rows = db.execute(
        "SELECT event_type,image_sha256,status FROM order_notification_jobs ORDER BY queue_job_id"
    ).fetchall()
    assert [row["event_type"] for row in rows] == ["ORDER_READY", "ORDER_UPDATED", "ORDER_CANCELLED"]
    assert len({row["image_sha256"] for row in rows}) == 3
    assert all(row["status"] == "SENT" for row in rows)
    summary = notification_summary(db, "1-1465")
    assert summary["latest"]["event_type"] == "ORDER_CANCELLED"
    assert all(job["attempts"] for job in summary["jobs"])


def test_debounce_refreshes_update_diff_and_skips_reverted_change(db, tmp_path):
    _target(db)
    ready = _create(db)
    process_notification_job(db, {"aggregate_id": ready["job"]["id"]}, {}, output_dir=str(tmp_path))

    db.execute(
        "UPDATE orders SET total='319.00',date_modified='2026-08-13T04:00:00' WHERE id='1-1465'"
    )
    db.commit()
    updated = _create(db, "ORDER_UPDATED", "evt-update")
    db.execute(
        "UPDATE orders SET total='325.00',date_modified='2026-08-13T04:01:00' WHERE id='1-1465'"
    )
    db.commit()
    process_notification_job(
        db, {"aggregate_id": updated["job"]["id"]}, {}, output_dir=str(tmp_path / "changed")
    )
    changes = json.loads(
        db.execute("SELECT changed_fields_json FROM order_notification_jobs WHERE id=?", (updated["job"]["id"],)).fetchone()[0]
    )
    assert {"field": "total", "before": "313.42", "after": "325.00"} in changes

    db.execute(
        "UPDATE orders SET total='330.00',date_modified='2026-08-13T04:02:00' WHERE id='1-1465'"
    )
    db.commit()
    reverted = _create(db, "ORDER_UPDATED", "evt-reverted")
    db.execute(
        "UPDATE orders SET total='325.00',date_modified='2026-08-13T04:03:00' WHERE id='1-1465'"
    )
    db.commit()
    assert process_notification_job(
        db, {"aggregate_id": reverted["job"]["id"]}, {}, output_dir=str(tmp_path / "reverted")
    ) == {"skipped": "no_material_change"}


def test_manual_resend_supports_cancelled_and_hold_orders(db, tmp_path):
    _target(db)
    ready = _create(db)
    process_notification_job(db, {"aggregate_id": ready["job"]["id"]}, {}, output_dir=str(tmp_path))

    for index, status in enumerate(("cancelled", "on-hold"), start=1):
        db.execute(
            "UPDATE orders SET status=?,date_modified=? WHERE id='1-1465'",
            (status, f"2026-08-13T05:0{index}:00"),
        )
        db.commit()
        resent = create_job_for_order(
            db,
            "1-1465",
            event_id=f"manual-{status}",
            resend_of=ready["job"]["id"],
            actor={"type": "user", "id": "admin"},
        )
        result = process_notification_job(
            db,
            {"aggregate_id": resent["job"]["id"]},
            {},
            output_dir=str(tmp_path / status),
        )
        assert result["accepted"] is True


def test_multilingual_long_order_paginates_without_truncating_quantity(tmp_path):
    snapshot = {
        "order_id": "1-999",
        "number": "999",
        "store_id": STORE,
        "store_label": "example.test",
        "status": "processing",
        "created_at": "2026-08-13 10:00",
        "notification_at": "2026-08-13 10:01",
        "warehouse_name": "波兰主仓",
        "shipping_method": "InPost Paczkomat",
        "payment_method": "COD",
        "currency": "PLN",
        "total": "999.00",
        "recipient": {"name_masked": "J**", "phone_masked": "***123", "city": "Łódź"},
        "items": [
            {
                "sku": f"SKU-{i}",
                "name": "Bardzo długa nazwa produktu Żółć 中文 English " * 4,
                "variation": "Smak / 口味",
                "quantity": i + 1,
            }
            for i in range(15)
        ],
        "customer_note": "Zażółć gęślą jaźń · 中文备注 · keep upright " * 8,
        "internal_order_url": "/orders?order_id=1-999",
    }
    rendered = render_order_cards(snapshot, "ORDER_UPDATED", tmp_path, "visual-job")
    assert len(rendered) == 3
    assert [item["page"] for item in rendered] == [1, 2, 3]
    assert rendered[0]["height"] >= 2100
    for item in rendered:
        assert item["bytes"] < 2 * 1024 * 1024
        with Image.open(item["path"]) as image:
            assert image.width == 1080
            assert image.height >= 1120


def test_source_files_do_not_contain_inline_wecom_secret():
    root = Path(__file__).resolve().parents[1]
    for path in [
        root / "order_notification_provider.py",
        root / "order_notification_service.py",
        root / "order_notification_api.py",
        root / "deploy" / "order-notification.env.example",
    ]:
        text = path.read_text("utf-8")
        assert "qyapi.weixin.qq.com/cgi-bin/webhook/send?key=" not in text


def test_select_admin_new_order_email_uses_template_even_when_recipient_matches_billing():
    logs = [
        {"id": 1, "status": "sent", "subject": "Your order #1465 has been received", "to": "jan@example.test"},
        {"id": 2, "status": "sent", "subject": "[Shop] New order #14650", "to": "ops@example.test"},
        {"id": 3, "status": "failed", "subject": "[Shop] New order #1465", "to": "ops@example.test"},
        {"id": 4, "status": "sent", "subject": "[Shop] New order #1465", "to": "jan@example.test", "sent_at": "2026-08-13T11:00:00"},
        {"id": 5, "status": "sent", "subject": "[Shop] New order #1465", "to": {"ops@example.test": "Ops"}, "sent_at": "2026-08-13T10:00:00"},
    ]
    selected = order_notification_email.select_admin_new_order_log(
        logs, "1465", "jan@example.test"
    )
    assert selected["id"] == 4

    with pytest.raises(order_notification_email.EmailRenderError) as customer_only:
        order_notification_email.select_admin_new_order_log(
            [logs[0]], "1465", "jan@example.test"
        )
    assert customer_only.value.code == "admin_new_order_email_not_found"


def test_fetch_logged_admin_email_uses_get_only_and_keeps_credentials_in_headers(db):
    class Response:
        status_code = 200

        def __init__(self, payload):
            self.payload = payload

        def json(self):
            return self.payload

    class Session:
        def __init__(self):
            self.calls = []

        def get(self, url, **kwargs):
            self.calls.append((url, kwargs))
            if url.endswith("/5"):
                return Response({
                    "success": True,
                    "id": 5,
                    "status": "sent",
                    "subject": "[Shop] New order #1465",
                    "to": "ops@example.test",
                    "body": "<html><body>Order #1465</body></html>",
                    "plugin": "FluentSMTP",
                })
            return Response({
                "success": True,
                "plugin": "FluentSMTP",
                "logs": [{
                    "id": 5,
                    "status": "sent",
                    "subject": "[Shop] New order #1465",
                    "to": "ops@example.test",
                }],
            })

    session = Session()
    email = order_notification_email.fetch_admin_new_order_email(
        db, "1-1465", session=session
    )
    assert email["log_id"] == 5 and email["plugin"] == "FluentSMTP"
    assert len(session.calls) == 2
    assert all(call[0].startswith(STORE + "/wp-json/woo-tracking/v1/orders/1465/email-logs") for call in session.calls)
    assert all("ck_test" not in call[0] and "cs_test" not in call[0] for call in session.calls)
    assert all(call[1]["headers"]["X-Woo-Tracking-Key"] == "ck_test" for call in session.calls)


def test_sanitize_email_html_inlines_only_same_site_images_and_removes_active_content():
    image_buffer = __import__("io").BytesIO()
    Image.new("RGB", (2, 2), "red").save(image_buffer, "PNG")
    image_bytes = image_buffer.getvalue()

    class ImageResponse:
        status_code = 200
        headers = {"Content-Type": "image/png"}

        def iter_content(self, _size):
            yield image_bytes

    class ImageSession:
        def __init__(self):
            self.urls = []

        def get(self, url, **kwargs):
            self.urls.append((url, kwargs))
            return ImageResponse()

    image_session = ImageSession()
    sanitized, meta = order_notification_email.sanitize_email_html(
        """
        <html><head><style>@import 'https://evil.test/a.css'; .x{background:url(https://evil.test/x)}</style></head>
        <body onload="steal()"><script>steal()</script><form action="https://evil.test"><input></form>
        <a href="https://evil.test">text</a>
        <img src="https://example.test/product.png"><img src="https://tracker.test/pixel.gif">
        </body></html>
        """,
        STORE,
        image_session=image_session,
        public_host_validator=lambda host: True,
    )
    assert "<script" not in sanitized and "<form" not in sanitized and "onload" not in sanitized
    assert "https://evil.test" not in sanitized and "https://tracker.test" not in sanitized
    assert "Content-Security-Policy" in sanitized and "data:image/png;base64," in sanitized
    assert [url for url, _ in image_session.urls] == ["https://example.test/product.png"]
    assert image_session.urls[0][1]["allow_redirects"] is False
    assert "X-Woo-Tracking-Key" not in image_session.urls[0][1]["headers"]
    assert meta == {"images_inlined": 1, "images_removed": 1, "inline_image_bytes": len(image_bytes)}


def test_email_preview_is_superadmin_only_zero_send_and_audits_no_body(
    tmp_path, monkeypatch
):
    db_path = tmp_path / "email-preview.db"
    setup = _conn(db_path)
    setup.close()
    client, get_test_conn = _notification_api_test_client(db_path, monkeypatch)
    preview_root = tmp_path / "email-preview-private"
    monkeypatch.setenv("ORDER_NOTIFICATION_IMAGE_DIR", str(preview_root))

    def fake_render(_conn, order_id, output_dir, preview_id):
        path = Path(output_dir) / "email.png"
        Image.new("RGB", (80, 160), "white").save(path, "PNG")
        return ([{"page": 1, "path": str(path)}], {
            "log_id": 885,
            "plugin": "FluentSMTP",
            "source": "FluentSMTP",
            "subject": "[Shop] New order #1465",
            "sent_at": "2026-08-13T10:00:00",
            "images_inlined": 2,
            "images_removed": 1,
            "html_sha256": "a" * 64,
        })

    monkeypatch.setattr(order_notification_api, "render_logged_admin_email", fake_render)
    response = client.post(
        "/api/order-notifications/preview",
        headers={"X-Requested-With": "XMLHttpRequest"},
        json={"order_id": "1-1465", "event_type": "ORDER_READY", "source": "email"},
    )
    assert response.status_code == 200
    data = response.get_json()
    assert data["source"] == "email" and data["email"]["log_id"] == 885
    assert data["queued"] is False and data["sent"] is False
    verify = get_test_conn()
    assert verify.execute("SELECT COUNT(*) FROM order_notification_jobs").fetchone()[0] == 0
    assert verify.execute("SELECT COUNT(*) FROM oms_integration_jobs WHERE job_type='ORDER_NOTIFICATION'").fetchone()[0] == 0
    audit = verify.execute(
        "SELECT after_summary FROM notification_audit_logs WHERE action='preview_generated'"
    ).fetchone()[0]
    assert "New order" not in audit and "<html" not in audit and "jan@example.test" not in audit
    assert json.loads(audit)["email_log_id"] == 885
    verify.close()
    assert list(preview_root.iterdir()) == []


def test_browser_launch_env_replaces_protected_home_with_private_runtime_dir(
    tmp_path, monkeypatch
):
    monkeypatch.setenv("HOME", "/root")
    runtime_home = tmp_path / "chromium-home"

    env = order_notification_email._browser_launch_env(runtime_home)

    assert env["HOME"] == str(runtime_home.resolve())
    assert env["XDG_CONFIG_HOME"] == str((runtime_home / ".config").resolve())
    assert env["XDG_CACHE_HOME"] == str((runtime_home / ".cache").resolve())
    assert (runtime_home / ".config").is_dir()
    assert (runtime_home / ".cache").is_dir()


def test_shadow_worker_uses_logged_email_for_new_order_but_does_not_send(
    db, tmp_path, monkeypatch
):
    _target(db, "WECOM_BOT", environment="production", secret_ref="env:PROD_GROUP")
    db.execute("UPDATE settings SET value='email' WHERE key='order_notification_render_source'")
    db.commit()

    def fake_render(_conn, order_id, output_dir, job_id):
        path = Path(output_dir) / f"{job_id}-email.png"
        Image.new("RGB", (120, 240), "white").save(path, "PNG")
        raw = path.read_bytes()
        return ([{
            "page": 1,
            "path": str(path),
            "sha256": hashlib.sha256(raw).hexdigest(),
            "width": 120,
            "height": 240,
            "bytes": len(raw),
        }], {
            "log_id": 885,
            "html_sha256": "b" * 64,
            "images_inlined": 2,
            "images_removed": 0,
            "template_version": "woo-admin-email-v1",
        })

    monkeypatch.setattr(order_notification_service, "render_logged_admin_email", fake_render)
    created = _create(db, "ORDER_READY", "email-shadow")
    result = process_notification_job(
        db, {"aggregate_id": created["job"]["id"]}, {}, output_dir=str(tmp_path)
    )
    assert result == {"blocked": "feature_flag_off", "images": 1}
    job = db.execute(
        "SELECT status,template_version FROM order_notification_jobs WHERE id=?",
        (created["job"]["id"],),
    ).fetchone()
    assert dict(job) == {"status": "READY_PREVIEW", "template_version": "woo-admin-email-v1"}
    audit = db.execute(
        "SELECT after_summary FROM notification_audit_logs WHERE action='email_source_rendered'"
    ).fetchone()[0]
    assert json.loads(audit)["email_log_id"] == 885
    assert db.execute(
        "SELECT COUNT(*) FROM order_notification_attempts WHERE result='SUCCESS'"
    ).fetchone()[0] == 0
