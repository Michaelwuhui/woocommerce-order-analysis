import concurrent.futures
from datetime import datetime, timezone
import json
import uuid

import pytest
import requests

import db_backend as db
import inv_mapping_service
import sync_service
import sync_tasks
from external_operations import begin_operation, transition_operation
from oid_utils import make_oid
from sync_utils import upsert_orders_in_transaction


pytestmark = pytest.mark.skipif(
    not db.is_postgres_backend(),
    reason="requires the isolated PostgreSQL migration test database",
)


def _connection():
    return sync_service.get_connection()


def _cleanup_test_rows():
    connection = _connection()
    try:
        connection.execute("DELETE FROM sync_task_outbox")
        connection.execute("DELETE FROM sync_runs")
        connection.execute(
            """
            DELETE FROM external_operations
            WHERE created_by LIKE 'pytest:%'
               OR operation_type='pytest_operation'
               OR order_id IN (
                    SELECT id FROM orders WHERE order_key LIKE 'pytest-sync-%'
               )
            """
        )
        test_orders = connection.execute(
            "SELECT id FROM orders WHERE order_key LIKE 'pytest-sync-%'"
        ).fetchall()
        for row in test_orders:
            order_id = row["id"]
            connection.execute("DELETE FROM order_notes WHERE order_id=?", (order_id,))
            connection.execute("DELETE FROM shipping_logs WHERE order_id=?", (order_id,))
            connection.execute("DELETE FROM orders WHERE id=?", (order_id,))
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _snapshot_legacy_state():
    connection = _connection()
    try:
        sites = [
            dict(row)
            for row in connection.execute(
                """
                SELECT id,last_sync,api_status,last_api_error
                FROM sites ORDER BY id
                """
            ).fetchall()
        ]
        sequences = {}
        for name in ("order_notes_id_seq", "shipping_logs_id_seq"):
            row = connection.execute(
                f"SELECT last_value,is_called FROM {name}"
            ).fetchone()
            sequences[name] = (int(row["last_value"]), bool(row["is_called"]))
        return {"sites": sites, "sequences": sequences}
    finally:
        connection.close()


def _restore_legacy_state(snapshot):
    connection = _connection()
    try:
        for site in snapshot["sites"]:
            connection.execute(
                """
                UPDATE sites
                SET last_sync=?,api_status=?,last_api_error=?
                WHERE id=?
                """,
                (
                    site["last_sync"],
                    site["api_status"],
                    site["last_api_error"],
                    site["id"],
                ),
            )
        for name, (last_value, is_called) in snapshot["sequences"].items():
            connection.execute(
                f"SELECT setval('{name}', ?, ?)",
                (last_value, is_called),
            )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


@pytest.fixture(autouse=True)
def clean_pipeline_state():
    _cleanup_test_rows()
    snapshot = _snapshot_legacy_state()
    try:
        yield
    finally:
        _cleanup_test_rows()
        _restore_legacy_state(snapshot)


def _site_rows(limit=4):
    connection = _connection()
    try:
        return connection.execute(
            "SELECT id,url FROM sites ORDER BY id LIMIT ?", (limit,)
        ).fetchall()
    finally:
        connection.close()


def test_quick_auto_and_deep_concurrent_submissions_share_one_run():
    site_id = int(_site_rows(1)[0]["id"])
    modes = ["quick", "auto", "deep", "quick", "auto", "deep"]

    def submit(mode):
        return sync_service.start_sync(
            mode=mode,
            created_by="pytest:mutual-exclusion",
            site_ids=[site_id],
            publish=False,
        )

    with concurrent.futures.ThreadPoolExecutor(max_workers=len(modes)) as pool:
        results = list(pool.map(submit, modes))

    run_ids = {status["run_id"] for status, _created in results}
    assert len(run_ids) == 1
    assert sum(1 for _status, created in results if created) == 1

    connection = _connection()
    try:
        active = connection.execute(
            """
            SELECT count(*) FROM sync_runs
            WHERE status IN ('queued','running','recovering','cancelling')
            """
        ).fetchone()[0]
        assert active == 1
    finally:
        connection.close()


def test_three_distinct_sites_can_lock_while_same_site_is_serial():
    sites = _site_rows(4)
    assert len(sites) >= 4
    holders = [_connection() for _ in range(5)]
    locked = []
    try:
        for position in range(3):
            site_id = int(sites[position]["id"])
            assert sync_tasks._site_lock(holders[position], site_id) is True
            locked.append((holders[position], site_id))
        assert sync_tasks._site_lock(holders[3], int(sites[0]["id"])) is False
        # The fourth distinct lock proves the DB lock is per-site; the worker
        # service's tested concurrency=3 is what caps a batch at three sites.
        assert sync_tasks._site_lock(holders[4], int(sites[3]["id"])) is True
        locked.append((holders[4], int(sites[3]["id"])))
    finally:
        for connection, site_id in locked:
            sync_tasks._site_unlock(connection, site_id)
        for connection in holders:
            connection.close()


def test_lost_queued_delivery_is_republished_from_postgres(monkeypatch):
    site_id = int(_site_rows(1)[0]["id"])
    status, created = sync_service.start_sync(
        mode="quick",
        created_by="pytest:queued-recovery",
        site_ids=[site_id],
        publish=False,
    )
    assert created is True

    connection = _connection()
    try:
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='queued',updated_at=CURRENT_TIMESTAMP - interval '10 minutes'
            WHERE run_id=? AND site_id=? AND page=1
            """,
            (status["run_id"], site_id),
        )
        connection.execute(
            """
            UPDATE sync_task_outbox SET status='published'
            WHERE dedupe_key=?
            """,
            (f"fetch:{status['run_id']}:{site_id}:1",),
        )
        connection.commit()
    finally:
        connection.close()
    monkeypatch.setattr(sync_service, "publish_pending_outbox", lambda limit=100: 0)
    recovered = sync_service.recover_stale_work()
    assert recovered["dispatches"] == 1

    connection = _connection()
    try:
        row = connection.execute(
            """
            SELECT d.status,o.status AS outbox_status
            FROM sync_page_dispatches d
            JOIN sync_task_outbox o
              ON o.dedupe_key=('fetch:' || d.run_id::text || ':' || d.site_id || ':' || d.page)
            WHERE d.run_id=? AND d.site_id=? AND d.page=1
            """,
            (status["run_id"], site_id),
        ).fetchone()
        assert row["status"] == "retry"
        assert row["outbox_status"] == "pending"
    finally:
        connection.close()


def test_empty_outbox_publish_does_not_exhaust_small_connection_pool():
    for _ in range(6):
        assert sync_service.publish_pending_outbox(limit=1) == 0
    connection = _connection()
    try:
        assert connection.execute("SELECT 1").fetchone()[0] == 1
    finally:
        connection.close()


def test_parameterized_sql_preserves_literal_percent_formats_and_like_patterns():
    connection = _connection()
    try:
        row = connection.execute(
            "SELECT strftime('%Y-%m', ?), ? LIKE 'pytest:%'",
            ("2026-09-02T00:00:00", "pytest:value"),
        ).fetchone()
        assert row[0] == "2026-09"
        assert row[1] is True
    finally:
        connection.close()


def test_inventory_warehouse_rows_grouping_is_postgres_safe():
    connection = _connection()
    try:
        rows = inv_mapping_service.warehouse_rows(connection)
        assert isinstance(rows, list)
        if rows:
            assert {'id', 'inventory_authority', 'sku_count'} <= set(rows[0])
    finally:
        connection.close()


def test_cancellation_stops_queued_pages_without_partial_receipts():
    site_ids = [int(row["id"]) for row in _site_rows(3)]
    status, created = sync_service.start_sync(
        mode="deep",
        created_by="pytest:cancel",
        site_ids=site_ids,
        publish=False,
    )
    assert created is True
    cancelled = sync_service.cancel_sync(status["run_id"], requested_by="pytest")
    assert cancelled["status"] == "cancelled"
    assert cancelled["cancellation_requested"] is True

    connection = _connection()
    try:
        assert connection.execute(
            "SELECT count(*) FROM sync_page_receipts WHERE run_id=?",
            (status["run_id"],),
        ).fetchone()[0] == 0
        assert connection.execute(
            """
            SELECT count(*) FROM sync_page_dispatches
            WHERE run_id=? AND status<>'cancelled'
            """,
            (status["run_id"],),
        ).fetchone()[0] == 0
    finally:
        connection.close()


def test_stale_heartbeat_reports_interrupted_and_recovering_state():
    site_id = int(_site_rows(1)[0]["id"])
    status, created = sync_service.start_sync(
        mode="quick",
        created_by="pytest:stale-heartbeat",
        site_ids=[site_id],
        publish=False,
    )
    assert created is True

    connection = _connection()
    try:
        connection.execute(
            """
            UPDATE sync_runs
            SET status='running',
                heartbeat_at=CURRENT_TIMESTAMP - (? * interval '1 second')
            WHERE run_id=?
            """,
            (sync_service.HEARTBEAT_STALE_SECONDS + 5, status["run_id"]),
        )
        connection.commit()
    finally:
        connection.close()

    stale = sync_service.get_run_status(status["run_id"])
    assert stale["status"] == "running"
    assert stale["interruption_state"] == "recovering"
    assert stale["stale_seconds"] >= sync_service.HEARTBEAT_STALE_SECONDS
    assert stale["message"] == "任务已中断/正在恢复"


def test_beat_due_checks_honor_database_settings_without_duplicate_daily_run():
    keys = (
        "autosync_enabled",
        "autosync_interval",
        "deep_sync_enabled",
        "deep_sync_hour",
        "deep_sync_minute",
    )
    connection = _connection()
    originals = {}
    try:
        rows = connection.execute(
            "SELECT key,value FROM settings WHERE key IN (?,?,?,?,?)", keys
        ).fetchall()
        originals = {str(row["key"]): str(row["value"] or "") for row in rows}
        for key, value in (
            ("autosync_enabled", "true"),
            ("autosync_interval", "300"),
            ("deep_sync_enabled", "true"),
            ("deep_sync_hour", "0"),
            ("deep_sync_minute", "0"),
        ):
            connection.execute(
                "INSERT INTO settings(key,value) VALUES (?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value),
            )
        connection.commit()
    finally:
        connection.close()

    try:
        assert sync_tasks._auto_due()[0] is True
        connection = _connection()
        try:
            connection.execute(
                """
                INSERT INTO sync_runs(run_id,mode,status,created_by)
                VALUES (?,'auto','success','celery-beat:auto')
                """,
                (str(uuid.uuid4()),),
            )
            connection.commit()
        finally:
            connection.close()
        assert sync_tasks._auto_due()[0] is False

        assert sync_tasks._deep_due()[0] is True
        connection = _connection()
        try:
            connection.execute(
                """
                INSERT INTO sync_runs(run_id,mode,status,created_by)
                VALUES (?,'deep','success','celery-beat:deep')
                """,
                (str(uuid.uuid4()),),
            )
            connection.commit()
        finally:
            connection.close()
        assert sync_tasks._deep_due()[0] is False
    finally:
        connection = _connection()
        try:
            for key in keys:
                if key in originals:
                    connection.execute(
                        "UPDATE settings SET value=? WHERE key=?",
                        (originals[key], key),
                    )
                else:
                    connection.execute("DELETE FROM settings WHERE key=?", (key,))
            connection.commit()
        finally:
            connection.close()


def _synthetic_order(site_url, woo_id):
    stamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    return {
        "id": woo_id,
        "number": str(woo_id),
        "order_key": f"pytest-sync-{woo_id}",
        "status": "processing",
        "currency": "EUR",
        "date_created": stamp,
        "date_modified": stamp,
        "discount_total": "0.00",
        "shipping_total": "1.25",
        "total": "11.25",
        "total_tax": "0.00",
        "prices_include_tax": False,
        "billing": {},
        "shipping": {},
        "meta_data": [],
        "line_items": [
            {"id": 1, "sku": "PYTEST-SKU", "name": "Synthetic item", "quantity": 1}
        ],
        "tax_lines": [],
        "shipping_lines": [],
        "fee_lines": [],
        "coupon_lines": [],
        "refunds": [],
        "source": site_url,
    }


def test_writer_receipt_makes_redelivery_exactly_once():
    site = _site_rows(1)[0]
    site_id = int(site["id"])
    woo_id = 991230001
    order_id = make_oid(site_id, woo_id)
    status, _created = sync_service.start_sync(
        mode="quick",
        created_by="pytest:writer-idempotency",
        site_ids=[site_id],
        publish=False,
    )
    orders = [_synthetic_order(site["url"], woo_id)]
    notes = [{
        "id": 991230002,
        "_local_order_id": order_id,
        "note": "synthetic note",
        "date_created": "2026-09-02T00:00:00",
        "customer_note": False,
        "author": "pytest",
        "added_by_user": False,
    }]
    content_hash = sync_tasks._canonical_hash(orders, notes)
    payload = {
        "run_id": status["run_id"],
        "site_id": site_id,
        "page": 1,
        "orders": orders,
        "notes": notes,
        "content_hash": content_hash,
        "fetched_count": 1,
        "total_pages": 1,
        "is_last_page": True,
    }
    connection = _connection()
    try:
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='fetched',content_hash=?,fetched_count=1
            WHERE run_id=? AND site_id=? AND page=1
            """,
            (content_hash, status["run_id"], site_id),
        )
        connection.commit()
    finally:
        connection.close()

    first = sync_tasks._write_page_transaction(payload)
    second = sync_tasks._write_page_transaction(payload)
    assert first["written"] == 1
    assert second["duplicate"] is True

    connection = _connection()
    try:
        assert connection.execute(
            "SELECT count(*) FROM orders WHERE id=?", (order_id,)
        ).fetchone()[0] == 1
        assert connection.execute(
            "SELECT count(*) FROM order_notes WHERE order_id=? AND wc_note_id=?",
            (order_id, 991230002),
        ).fetchone()[0] == 1
        assert connection.execute(
            """
            SELECT count(*) FROM sync_page_receipts
            WHERE run_id=? AND site_id=? AND page=1
            """,
            (status["run_id"], site_id),
        ).fetchone()[0] == 1
    finally:
        connection.close()


def test_post_commit_hook_is_durable_idempotent_and_recovers_stale_claim(
    monkeypatch,
):
    site = _site_rows(1)[0]
    site_id = int(site["id"])
    status, _created = sync_service.start_sync(
        mode="quick",
        created_by="pytest:post-commit-recovery",
        site_ids=[site_id],
        publish=False,
    )
    run_id = status["run_id"]
    candidates = [{
        "order_id": "pytest-post-commit-order",
        "status": "processing",
        "date_modified": "2026-09-02T00:00:00+00:00",
        "source": str(site["url"]),
    }]
    connection = _connection()
    try:
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='completed',content_hash='post-commit-test',
                fetched_count=1,updated_at=CURRENT_TIMESTAMP
            WHERE run_id=? AND site_id=? AND page=1
            """,
            (run_id, site_id),
        )
        connection.execute(
            """
            INSERT INTO sync_page_receipts
                (run_id,site_id,page,content_hash,fetched_count,written_count,
                 changed_count,is_last_page,planning_candidates,
                 post_commit_status,post_commit_heartbeat_at)
            VALUES (?,?,1,'post-commit-test',1,1,1,true,?::jsonb,
                    'processing',CURRENT_TIMESTAMP - interval '10 minutes')
            """,
            (run_id, site_id, json.dumps(candidates)),
        )
        sync_service.enqueue_outbox(
            connection,
            dedupe_key=f"postcommit:{run_id}:{site_id}:1",
            queue_name="sync_write",
            task_name="woo_sync.post_commit_page",
            payload={"run_id": run_id, "site_id": site_id, "page": 1},
        )
        connection.execute(
            """
            UPDATE sync_task_outbox SET status='published'
            WHERE dedupe_key=?
            """,
            (f"postcommit:{run_id}:{site_id}:1",),
        )
        connection.commit()
    finally:
        connection.close()

    monkeypatch.setattr(sync_service, "publish_pending_outbox", lambda limit=100: 0)
    recovered = sync_service.recover_stale_work()
    assert recovered["post_commit"] == 1

    calls = []
    monkeypatch.setattr(
        sync_tasks,
        "run_post_commit_sync_actions",
        lambda items, *, strict=False: calls.append((items, strict)),
    )
    payload = {"run_id": run_id, "site_id": site_id, "page": 1}
    assert sync_tasks.post_commit_page.run(payload)["run_id"] == run_id
    assert sync_tasks.post_commit_page.run(payload) == {"duplicate_or_skipped": True}
    assert calls == [(candidates, True)]

    connection = _connection()
    try:
        receipt = connection.execute(
            """
            SELECT post_commit_status,post_commit_attempts,post_commit_finished_at
            FROM sync_page_receipts
            WHERE run_id=? AND site_id=? AND page=1
            """,
            (run_id, site_id),
        ).fetchone()
        assert receipt["post_commit_status"] == "completed"
        assert int(receipt["post_commit_attempts"]) == 1
        assert receipt["post_commit_finished_at"] is not None
        outbox_status = connection.execute(
            "SELECT status FROM sync_task_outbox WHERE dedupe_key=?",
            (f"postcommit:{run_id}:{site_id}:1",),
        ).fetchone()[0]
        assert outbox_status == "pending"
    finally:
        connection.close()


def test_external_operation_state_machine_and_failed_retry_claim():
    site_id = int(_site_rows(1)[0]["id"])
    connection = _connection()
    try:
        payload = {"target_status": "completed", "token": "pytest-state"}
        first = begin_operation(
            connection,
            operation_type="pytest_operation",
            order_id="pytest-operation-order",
            site_id=site_id,
            request_payload=payload,
            created_by="pytest:external-state",
        )
        connection.commit()
        assert first["should_execute"] is True

        duplicate = begin_operation(
            connection,
            operation_type="pytest_operation",
            order_id="pytest-operation-order",
            site_id=site_id,
            request_payload=payload,
            created_by="pytest:external-state",
        )
        connection.commit()
        assert duplicate["should_execute"] is False

        for state in ("external_success", "local_committed", "notified"):
            transition_operation(connection, first["operation_id"], state)
            connection.commit()
        terminal = begin_operation(
            connection,
            operation_type="pytest_operation",
            order_id="pytest-operation-order",
            site_id=site_id,
            request_payload=payload,
            created_by="pytest:external-state",
        )
        connection.commit()
        assert terminal["status"] == "notified"
        assert terminal["should_execute"] is False

        failed_payload = {"target_status": "failed-retry", "token": "pytest-state"}
        failed = begin_operation(
            connection,
            operation_type="pytest_operation",
            order_id="pytest-operation-order",
            site_id=site_id,
            request_payload=failed_payload,
            created_by="pytest:external-state",
        )
        transition_operation(connection, failed["operation_id"], "failed")
        connection.commit()
        resumed = begin_operation(
            connection,
            operation_type="pytest_operation",
            order_id="pytest-operation-order",
            site_id=site_id,
            request_payload=failed_payload,
            created_by="pytest:external-state",
        )
        connection.commit()
        assert resumed["resumed"] is True
        assert resumed["should_execute"] is True
        assert int(resumed["attempts"]) == 2
    finally:
        connection.close()


def test_mock_order_note_response_is_mirrored_with_postgres_booleans(monkeypatch):
    import app as app_module

    connection = _connection()
    try:
        site = connection.execute("SELECT * FROM sites ORDER BY id LIMIT 1").fetchone()
        admin = connection.execute(
            "SELECT id,name,username FROM users WHERE username='admin' LIMIT 1"
        ).fetchone()
        assert site and admin
        woo_id = 991230051
        upsert_orders_in_transaction([_synthetic_order(site['url'], woo_id)], connection)
        connection.commit()
        order_id = make_oid(int(site['id']), woo_id)
    finally:
        connection.close()

    note_id = 991230052
    note_text = 'pytest PostgreSQL boolean note mirror'

    class Response:
        status_code = 201
        text = '{}'

        @staticmethod
        def json():
            return {
                'id': note_id,
                'note': note_text,
                'date_created': '2026-09-03T00:00:00',
                'customer_note': False,
                'author': 'WooCommerce',
                'added_by_user': False,
            }

    calls = []

    def fake_post(url, **kwargs):
        calls.append((url, kwargs['json']))
        return Response()

    monkeypatch.setattr(requests, 'post', fake_post)
    app_module.app.config.update(TESTING=True)
    client = app_module.app.test_client()
    with client.session_transaction() as session:
        session['_user_id'] = str(admin['id'])
        session['_fresh'] = True

    response = client.post(
        f'/api/order/{order_id}/note',
        json={'note': note_text, 'notify_customer': False},
    )

    assert response.status_code == 200
    assert response.get_json()['success'] is True
    assert len(calls) == 1
    connection = _connection()
    try:
        stored = connection.execute(
            """SELECT note,customer_note,author,added_by_user
               FROM order_notes WHERE order_id=? AND wc_note_id=?""",
            (order_id, note_id),
        ).fetchone()
        assert stored['note'] == note_text
        assert stored['customer_note'] is False
        assert stored['author'] == (admin['name'] or admin['username'])
        assert stored['added_by_user'] is True
    finally:
        connection.close()


def test_mock_shipping_and_delivery_confirmation_never_repeat_remote_effects(monkeypatch):
    import app as app_module

    connection = _connection()
    try:
        site = connection.execute("SELECT * FROM sites ORDER BY id LIMIT 1").fetchone()
        carrier = connection.execute(
            "SELECT slug FROM shipping_carriers ORDER BY id LIMIT 1"
        ).fetchone()
        admin = connection.execute(
            "SELECT id FROM users WHERE username='admin' LIMIT 1"
        ).fetchone()
        assert site and carrier and admin
        woo_id = 991230101
        order = _synthetic_order(site["url"], woo_id)
        upsert_orders_in_transaction([order], connection)
        connection.commit()
        order_id = make_oid(int(site["id"]), woo_id)
    finally:
        connection.close()

    calls = {"put": 0, "get": 0, "post": 0}
    tracking = "PYTEST-TRACKING-IDEMPOTENT"

    class Response:
        status_code = 200
        text = "{}"

        def __init__(self, payload):
            self.payload = payload

        def json(self):
            return self.payload

    def fake_put(_url, *, json, **_kwargs):
        calls["put"] += 1
        status = json["status"]
        payload = {"id": woo_id, "status": status, "meta_data": [], "line_items": []}
        if status == "on-hold":
            payload["meta_data"] = [{"key": "_tracking_number", "value": tracking}]
        return Response(payload)

    def fake_get(*_args, **_kwargs):
        calls["get"] += 1
        raise AssertionError("verified success responses must not need GET")

    def fake_post(*_args, **_kwargs):
        calls["post"] += 1
        raise AssertionError("send_email=false must not notify")

    monkeypatch.setattr(requests, "put", fake_put)
    monkeypatch.setattr(requests, "get", fake_get)
    monkeypatch.setattr(requests, "post", fake_post)
    monkeypatch.setattr(app_module, "detect_site_tracking_format", lambda *_args: "unknown")
    monkeypatch.setattr(
        app_module, "_manual_partner_only_warehouse_scope", lambda *_args: False
    )
    app_module.app.config.update(TESTING=True)
    client = app_module.app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = str(admin["id"])
        session["_fresh"] = True

    shipment = {
        "order_id": order_id,
        "tracking_number": tracking,
        "carrier_slug": carrier["slug"],
        "send_email": False,
    }
    first_ship = client.post("/api/shipping/ship", json=shipment)
    second_ship = client.post("/api/shipping/ship", json=shipment)
    assert first_ship.status_code == 200
    assert second_ship.status_code == 200
    assert second_ship.get_json()["idempotent"] is True
    assert calls == {"put": 1, "get": 0, "post": 0}

    first_confirm = client.post(f"/api/order/{order_id}/confirm-delivery")
    second_confirm = client.post(f"/api/order/{order_id}/confirm-delivery")
    assert first_confirm.status_code == 200
    assert second_confirm.status_code == 200
    assert calls == {"put": 2, "get": 0, "post": 0}

    connection = _connection()
    try:
        stored = connection.execute(
            "SELECT status,delivery_confirmed FROM orders WHERE id=?", (order_id,)
        ).fetchone()
        assert stored["status"] == "completed"
        assert bool(stored["delivery_confirmed"]) is True
        assert connection.execute(
            "SELECT count(*) FROM shipping_logs WHERE order_id=? AND tracking_number=?",
            (order_id, tracking),
        ).fetchone()[0] == 1
        assert connection.execute(
            """
            SELECT count(*) FROM external_operations
            WHERE order_id=? AND status IN ('local_committed','notified')
            """,
            (order_id,),
        ).fetchone()[0] == 2
    finally:
        connection.close()


def test_ast_gateway_error_accepts_verified_tracking_after_completed_status_hook(monkeypatch):
    """A completed status hook must not turn a committed parcel into a replay."""
    import app as app_module

    connection = _connection()
    try:
        site = connection.execute("SELECT * FROM sites ORDER BY id LIMIT 1").fetchone()
        carrier = connection.execute(
            "SELECT slug FROM shipping_carriers ORDER BY id LIMIT 1"
        ).fetchone()
        admin = connection.execute(
            "SELECT id FROM users WHERE username='admin' LIMIT 1"
        ).fetchone()
        assert site and carrier and admin
        woo_id = 991230102
        upsert_orders_in_transaction(
            [_synthetic_order(site["url"], woo_id)], connection
        )
        connection.commit()
        order_id = make_oid(int(site["id"]), woo_id)
    finally:
        connection.close()

    calls = {"post": 0, "get": 0, "put": 0}
    tracking = "PYTEST-AST-COMPLETED-HOOK"

    class GatewayResponse:
        status_code = 504
        text = "<html>gateway timeout</html>"

        @staticmethod
        def json():
            raise ValueError("edge returned HTML")

    class VerifiedOrderResponse:
        status_code = 200
        text = "{}"

        @staticmethod
        def json():
            return {
                "id": woo_id,
                "status": "completed",
                "meta_data": [
                    {"key": "_tracking_number", "value": tracking},
                ],
                "line_items": [],
            }

    def fake_post(*_args, **_kwargs):
        calls["post"] += 1
        return GatewayResponse()

    def fake_get(*_args, **_kwargs):
        calls["get"] += 1
        return VerifiedOrderResponse()

    def fake_put(*_args, **_kwargs):
        calls["put"] += 1
        raise AssertionError("AST shipment must not use the generic PUT path")

    monkeypatch.setattr(requests, "post", fake_post)
    monkeypatch.setattr(requests, "get", fake_get)
    monkeypatch.setattr(requests, "put", fake_put)
    monkeypatch.setattr(app_module, "detect_site_tracking_format", lambda *_args: "ast")
    monkeypatch.setattr(
        app_module, "_manual_partner_only_warehouse_scope", lambda *_args: False
    )
    app_module.app.config.update(TESTING=True)
    client = app_module.app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = str(admin["id"])
        session["_fresh"] = True

    response = client.post(
        "/api/shipping/ship",
        json={
            "order_id": order_id,
            "tracking_number": tracking,
            "carrier_slug": carrier["slug"],
            "send_email": False,
        },
    )

    assert response.status_code == 200
    assert response.get_json()["success"] is True
    assert calls == {"post": 1, "get": 1, "put": 0}

    connection = _connection()
    try:
        stored = connection.execute(
            "SELECT status FROM orders WHERE id=?", (order_id,)
        ).fetchone()
        assert stored["status"] == "completed"
        shipping_log = connection.execute(
            """
            SELECT status FROM shipping_logs
            WHERE order_id=? AND tracking_number=?
            """,
            (order_id, tracking),
        ).fetchone()
        assert shipping_log["status"] == "shipped"
        operation = connection.execute(
            """
            SELECT status,external_evidence FROM external_operations
            WHERE order_id=? AND operation_type='ship_order'
            """,
            (order_id,),
        ).fetchone()
        assert operation["status"] == "local_committed"
        evidence = operation["external_evidence"]
        if isinstance(evidence, str):
            evidence = json.loads(evidence)
        assert evidence["confirmation"] == "verified_get"
        assert evidence["remote_status"] == "completed"
    finally:
        connection.close()
