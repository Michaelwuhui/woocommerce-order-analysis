"""End-to-end clean-sync writes against an empty isolated PostgreSQL schema."""

from datetime import datetime, timezone
import os

import pytest

import clean_sync
import db_backend as db
import sync_service
from oid_utils import make_oid
from sync_utils import upsert_orders_in_transaction


pytestmark = pytest.mark.skipif(
    not db.is_postgres_backend(), reason="isolated PostgreSQL required"
)

SITE_ID = 991239001
SITE_URL = "https://clean-sync.example.invalid"
WOO_MISSING = 991239101
WOO_PRESENT = 991239102
WOO_REFERENCED = 991239103


class Response:
    def __init__(self, status, body, *, total=None, pages=None):
        self.status_code = status
        self.body = body
        self.headers = {}
        if total is not None:
            self.headers["X-WP-Total"] = str(total)
        if pages is not None:
            self.headers["X-WP-TotalPages"] = str(pages)

    def json(self):
        return self.body


class FakeAPI:
    calls = []
    fail_scan = False

    def __init__(self, **_kwargs):
        pass

    def get(self, path, params=None):
        self.calls.append((path, params))
        if path == "orders" and params and params.get("_fields") == "id":
            if self.fail_scan:
                return Response(403, {"code": "forbidden"})
            return Response(200, [{"id": WOO_PRESENT}], total=1, pages=1)
        if path in {f"orders/{WOO_MISSING}", f"orders/{WOO_REFERENCED}"}:
            return Response(404, {"code": "woocommerce_rest_shop_order_invalid_id"})
        if path == "orders" and params and params.get("status") == "trash":
            return Response(200, [])
        raise AssertionError((path, params))


def _order(woo_id):
    stamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    return {
        "id": woo_id, "number": str(woo_id),
        "order_key": f"pytest-clean-{woo_id}", "status": "processing",
        "currency": "PLN", "date_created": stamp, "date_modified": stamp,
        "discount_total": "0.00", "shipping_total": "0.00",
        "total": "10.00", "total_tax": "0.00",
        "prices_include_tax": False, "billing": {}, "shipping": {},
        "meta_data": [], "line_items": [], "tax_lines": [],
        "shipping_lines": [], "fee_lines": [], "coupon_lines": [],
        "refunds": [], "source": SITE_URL,
    }


@pytest.fixture()
def isolated(monkeypatch):
    database = os.environ.get("WOO_DB_NAME_OVERRIDE", "")
    assert database.startswith("woo_return_loss_test_"), "Isolated test DB required"
    connection = sync_service.get_connection()
    assert connection.execute("SELECT current_database()").fetchone()[0] == database
    connection.close()
    monkeypatch.setattr(sync_service, "publish_pending_outbox", lambda **_kw: 0)
    monkeypatch.setattr(clean_sync, "API", FakeAPI)
    monkeypatch.setattr(clean_sync.time, "sleep", lambda _seconds: None)
    FakeAPI.calls = []
    FakeAPI.fail_scan = False
    connection = sync_service.get_connection()
    try:
        connection.execute(
            """INSERT INTO sites(id,url,consumer_key,consumer_secret,country)
               VALUES (?,?,?,?,'PL')""",
            (SITE_ID, SITE_URL, "ck_fake", "cs_fake"),
        )
        for woo_id in (WOO_MISSING, WOO_PRESENT, WOO_REFERENCED):
            upsert_orders_in_transaction([_order(woo_id)], connection)
        connection.execute(
            """INSERT INTO order_notes(wc_note_id,order_id,note,date_created)
               VALUES (?,?,?,?)""",
            (991239201, make_oid(SITE_ID, WOO_MISSING), "Synthetic note", "2026-09-24"),
        )
        connection.execute(
            "INSERT INTO order_note_sync_state(order_id,last_synced_at) VALUES (?,?)",
            (make_oid(SITE_ID, WOO_MISSING), "2026-09-24"),
        )
        connection.execute(
            """INSERT INTO shipping_logs(order_id,woo_order_id,source,tracking_number)
               VALUES (?,?,?,?)""",
            (make_oid(SITE_ID, WOO_REFERENCED), WOO_REFERENCED, SITE_URL, "TEST-TRACK"),
        )
        connection.commit()
    finally:
        connection.close()
    yield
    connection = sync_service.get_connection()
    try:
        connection.execute("DELETE FROM sync_task_outbox")
        connection.execute("DELETE FROM sync_runs")
        for table in ("order_notes_archive", "order_notes", "order_note_sync_state"):
            connection.execute(
                f"DELETE FROM {table} WHERE order_id IN (?,?,?)",
                tuple(make_oid(SITE_ID, value) for value in (
                    WOO_MISSING, WOO_PRESENT, WOO_REFERENCED
                )),
            )
        connection.execute(
            "DELETE FROM shipping_logs WHERE source=? AND tracking_number='TEST-TRACK'",
            (SITE_URL,),
        )
        for table in ("orders_archive", "orders"):
            connection.execute(f"DELETE FROM {table} WHERE source=?", (SITE_URL,))
        connection.execute("DELETE FROM sites WHERE id=?", (SITE_ID,))
        connection.commit()
    finally:
        connection.close()


def test_clean_run_archives_only_proven_unreferenced_order_and_is_idempotent(isolated):
    status, created = sync_service.start_sync(
        mode="clean", created_by="pytest:clean", site_ids=[SITE_ID], publish=False
    )
    assert created and status["mode"] == "clean"
    connection = sync_service.get_connection()
    try:
        assert connection.execute(
            "SELECT count(*) FROM sync_page_dispatches WHERE run_id=?",
            (status["run_id"],),
        ).fetchone()[0] == 0
        assert connection.execute(
            "SELECT task_name FROM sync_task_outbox WHERE dedupe_key=?",
            (f"clean:{status['run_id']}:{SITE_ID}",),
        ).fetchone()[0] == "woo_sync.clean_site"
    finally:
        connection.close()
    result = clean_sync.run_clean_site({"run_id": status["run_id"], "site_id": SITE_ID})
    assert result == {"removed": 1, "retained": 1}
    assert clean_sync.run_clean_site(
        {"run_id": status["run_id"], "site_id": SITE_ID}
    ) == {"skipped": True}
    final = sync_service.get_run_status(status["run_id"])
    assert final["status"] == "success"
    assert final["written_orders"] == 1
    connection = sync_service.get_connection()
    try:
        missing = make_oid(SITE_ID, WOO_MISSING)
        present = make_oid(SITE_ID, WOO_PRESENT)
        referenced = make_oid(SITE_ID, WOO_REFERENCED)
        assert connection.execute(
            "SELECT count(*) FROM orders WHERE id=?", (missing,)
        ).fetchone()[0] == 0
        assert connection.execute(
            "SELECT archive_reason FROM orders_archive WHERE id=?", (missing,)
        ).fetchone()[0] == "orphaned_remote_deleted"
        assert connection.execute(
            "SELECT note FROM order_notes_archive WHERE order_id=?", (missing,)
        ).fetchone()[0] == "Synthetic note"
        assert connection.execute(
            "SELECT count(*) FROM order_notes WHERE order_id=?", (missing,)
        ).fetchone()[0] == 0
        assert connection.execute(
            "SELECT count(*) FROM order_note_sync_state WHERE order_id=?", (missing,)
        ).fetchone()[0] == 0
        assert connection.execute(
            "SELECT count(*) FROM orders WHERE id IN (?,?)", (present, referenced)
        ).fetchone()[0] == 2
    finally:
        connection.close()


def test_failed_remote_scan_does_not_archive_or_delete(isolated):
    FakeAPI.fail_scan = True
    status, _ = sync_service.start_sync(
        mode="clean", created_by="pytest:failed-scan",
        site_ids=[SITE_ID], publish=False,
    )
    result = clean_sync.run_clean_site({"run_id": status["run_id"], "site_id": SITE_ID})
    assert "error" in result
    assert sync_service.get_run_status(status["run_id"])["status"] == "error"
    connection = sync_service.get_connection()
    try:
        assert connection.execute(
            "SELECT count(*) FROM orders WHERE source=?", (SITE_URL,)
        ).fetchone()[0] == 3
        assert connection.execute(
            "SELECT count(*) FROM orders_archive WHERE source=?", (SITE_URL,)
        ).fetchone()[0] == 0
        assert connection.execute(
            "SELECT count(*) FROM order_notes_archive WHERE order_id=?",
            (make_oid(SITE_ID, WOO_MISSING),),
        ).fetchone()[0] == 0
        assert connection.execute(
            "SELECT count(*) FROM order_notes WHERE order_id=?",
            (make_oid(SITE_ID, WOO_MISSING),),
        ).fetchone()[0] == 1
    finally:
        connection.close()


def test_stale_clean_site_is_requeued_without_deleting_orders(isolated):
    status, _ = sync_service.start_sync(
        mode="clean", created_by="pytest:recovery",
        site_ids=[SITE_ID], publish=False,
    )
    connection = sync_service.get_connection()
    try:
        connection.execute(
            """UPDATE sync_site_progress SET status='fetching',
               heartbeat_at=CURRENT_TIMESTAMP - interval '5 minutes'
               WHERE run_id=? AND site_id=?""",
            (status["run_id"], SITE_ID),
        )
        connection.execute(
            "UPDATE sync_runs SET status='running' WHERE run_id=?",
            (status["run_id"],),
        )
        connection.execute(
            "UPDATE sync_task_outbox SET status='published' WHERE dedupe_key=?",
            (f"clean:{status['run_id']}:{SITE_ID}",),
        )
        connection.commit()
    finally:
        connection.close()
    result = sync_service.recover_stale_work()
    assert result["clean_sites"] == 1
    connection = sync_service.get_connection()
    try:
        progress = connection.execute(
            "SELECT status,retry_count FROM sync_site_progress WHERE run_id=? AND site_id=?",
            (status["run_id"], SITE_ID),
        ).fetchone()
        assert (progress["status"], progress["retry_count"]) == ("recovering", 1)
        assert connection.execute(
            "SELECT status FROM sync_task_outbox WHERE dedupe_key=?",
            (f"clean:{status['run_id']}:{SITE_ID}",),
        ).fetchone()[0] == "pending"
        assert connection.execute(
            "SELECT count(*) FROM orders WHERE source=?", (SITE_URL,)
        ).fetchone()[0] == 3
    finally:
        connection.close()


def test_archive_and_order_removal_roll_back_together(isolated, monkeypatch):
    original_event = clean_sync._event

    def fail_after_writes(connection, run_id, event_type, message, **kwargs):
        if event_type == "clean_site_completed":
            raise RuntimeError("synthetic commit failure")
        return original_event(connection, run_id, event_type, message, **kwargs)

    monkeypatch.setattr(clean_sync, "_event", fail_after_writes)
    status, _ = sync_service.start_sync(
        mode="clean", created_by="pytest:rollback",
        site_ids=[SITE_ID], publish=False,
    )
    result = clean_sync.run_clean_site({"run_id": status["run_id"], "site_id": SITE_ID})
    assert "error" in result
    connection = sync_service.get_connection()
    try:
        missing = make_oid(SITE_ID, WOO_MISSING)
        assert connection.execute(
            "SELECT COUNT(*) FROM orders WHERE id=?", (missing,)
        ).fetchone()[0] == 1
        assert connection.execute(
            "SELECT COUNT(*) FROM orders_archive WHERE id=?", (missing,)
        ).fetchone()[0] == 0
        assert connection.execute(
            "SELECT COUNT(*) FROM order_notes WHERE order_id=?", (missing,)
        ).fetchone()[0] == 1
        assert connection.execute(
            "SELECT COUNT(*) FROM order_notes_archive WHERE order_id=?", (missing,)
        ).fetchone()[0] == 0
    finally:
        connection.close()
