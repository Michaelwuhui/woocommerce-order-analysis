"""履约后台 Worker 实时通知优先领取、物流查询释放读事务与断连自愈单元测试。"""

from __future__ import annotations

import sqlite3
import pytest

import carrier_tracking
import fulfillment_worker as fw
from fulfillment_service import enqueue_job
import inv_migrations


def _setup_oms_db(db_path: str):
    """初始化包含多仓履约任务表的临时测试数据库。"""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT)")
    conn.execute("CREATE TABLE sites (id INTEGER PRIMARY KEY, url TEXT UNIQUE, country TEXT)")
    conn.execute(
        "CREATE TABLE warehouses (id INTEGER PRIMARY KEY, name TEXT, code TEXT, country TEXT, is_active INTEGER DEFAULT 1)"
    )
    conn.execute("INSERT INTO warehouses (id, name, code, country) VALUES (1, '波兰仓', 'PL', 'PL')")
    from sync_utils import WC_ORDER_FIELDS
    order_cols = ", ".join("id TEXT PRIMARY KEY" if f == "id" else f"{f} TEXT" for f in WC_ORDER_FIELDS)
    conn.execute(f"CREATE TABLE orders ({order_cols}, woo_id INTEGER, updated_at TEXT)")
    for fn in (
        inv_migrations.up_001,
        inv_migrations.up_006,
        inv_migrations.up_007,
        inv_migrations.up_008,
    ):
        fn(conn)
    conn.commit()
    return conn


def test_claim_job_prioritizes_notifications_and_planning_ahead_of_tracking_polls(tmp_path, monkeypatch):
    """验证即使存在更早入队的批量查物流任务，企微新单通知与锁库存任务仍优先插队执行。"""
    monkeypatch.setenv("WOO_DB_BACKEND", "sqlite")
    db_file = str(tmp_path / "prio.db")
    conn = _setup_oms_db(db_file)
    try:
        enqueue_job(
            conn,
            "POLL_SHIPMENT_TRACKING",
            "shipment",
            "ship-1",
            "track:ship-1:early",
            {"shipment_id": "ship-1"},
            available_at="2026-11-17T10:00:00+00:00",
        )
        enqueue_job(
            conn,
            "RECONCILE",
            "system",
            "fulfillment",
            "reconcile:early",
            {},
            available_at="2026-11-17T10:01:00+00:00",
        )
        enqueue_job(
            conn,
            "POLL_WMS_STATUS",
            "fulfillment",
            "ful-1",
            "wms:mid",
            {"fulfillment_id": "ful-1"},
            available_at="2026-11-17T10:02:00+00:00",
        )
        enqueue_job(
            conn,
            "ORDER_NOTIFICATION",
            "order_notification",
            "notif-2715",
            "order-notification:2715",
            {"notification_job_id": "notif-2715"},
            available_at="2026-11-17T10:10:49+00:00",
        )
        enqueue_job(
            conn,
            "PLAN_ORDER",
            "order",
            "39-2687",
            "plan:39-2687",
            {"order_id": "39-2687"},
            available_at="2026-11-17T10:10:04+00:00",
        )
        conn.commit()

        claimed_types = []
        while True:
            job = fw.claim_job(conn)
            if not job:
                break
            claimed_types.append(job["job_type"])
            fw.finish_job(conn, job["id"])

        assert claimed_types == [
            "PLAN_ORDER",
            "ORDER_NOTIFICATION",
            "POLL_WMS_STATUS",
            "POLL_SHIPMENT_TRACKING",
            "RECONCILE",
        ]
    finally:
        conn.close()


def test_handle_poll_tracking_releases_transaction_before_external_http(tmp_path, monkeypatch):
    """验证查物流任务在发起慢速外部网络请求前已提交并退出数据库读事务。"""
    monkeypatch.setenv("WOO_DB_BACKEND", "sqlite")
    db_file = str(tmp_path / "tx.db")
    conn = _setup_oms_db(db_file)
    try:
        conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('track718_api_key', 'k_test')")
        conn.execute("INSERT INTO orders (id, source, status) VALUES ('39-1', 'https://s.pl', 'shipped')")
        conn.execute(
            """
            INSERT INTO oms_fulfillments (id, order_id, warehouse_id, provider, status, idempotency_key)
            VALUES ('f-1', '39-1', 1, 'manual', 'shipped', 'idem-f1')
            """
        )
        conn.execute(
            """
            INSERT INTO oms_shipments (id, fulfillment_id, carrier_slug, tracking_number, status)
            VALUES ('s-1', 'f-1', 'dpd-pl', '12345678901234', 'shipped')
            """
        )
        conn.commit()

        in_tx_during_http = []

        def fake_track718_detail(number, key, **kwargs):
            in_tx_during_http.append(bool(conn.in_transaction))
            return {
                "ok": True,
                "result": 30,
                "outcome": "in_transit",
                "events": [{"time": "2026-11-17 12:00:00", "status": "In transit"}],
            }

        monkeypatch.setattr(carrier_tracking, "track718_detail", fake_track718_detail)
        out = fw.handle_poll_tracking(
            conn,
            {"id": 1, "aggregate_id": "s-1"},
            {"shipment_id": "s-1"},
        )
        assert out["third_party"] is True
        assert in_tx_during_http == [False]
        assert conn.in_transaction is False
    finally:
        conn.close()


def test_run_one_recovers_via_fallback_connection_when_primary_connection_lost(tmp_path, monkeypatch):
    """验证当任务执行期原数据库连接断开时，自动使用新连接持久化 retry 状态而不崩溃。"""
    monkeypatch.setenv("WOO_DB_BACKEND", "sqlite")
    db_file = str(tmp_path / "fallback.db")
    conn = _setup_oms_db(db_file)
    enqueue_job(
        conn,
        "POLL_SHIPMENT_TRACKING",
        "shipment",
        "s-99",
        "track:s-99",
        {"shipment_id": "s-99"},
        available_at="2026-11-17T10:00:00+00:00",
    )
    conn.commit()

    def fresh_conn():
        c = sqlite3.connect(db_file)
        c.row_factory = sqlite3.Row
        return c

    def boom_dispatch(c, job):
        c.close()
        raise RuntimeError("terminating connection due to idle-in-transaction timeout")

    monkeypatch.setattr(fw, "get_conn", fresh_conn)
    monkeypatch.setattr(fw, "dispatch", boom_dispatch)

    worked = fw.run_one(conn)
    assert worked is True

    verify = fresh_conn()
    try:
        row = verify.execute("SELECT status, last_error_code FROM oms_integration_jobs WHERE aggregate_id='s-99'").fetchone()
        assert row["status"] == "retry"
        assert row["last_error_code"] == "db_connection_lost"
    finally:
        verify.close()
