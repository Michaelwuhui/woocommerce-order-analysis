"""A bad domain must finish its site and release the global sync run."""
import os

import pytest
import requests

import db_backend as db
import sync_service
import sync_tasks
from sync_site_health import load_site_health
from sync_utils import upsert_orders_in_transaction
from oid_utils import make_oid


pytestmark = pytest.mark.skipif(
    not db.is_postgres_backend(), reason="isolated PostgreSQL required"
)
SITE_IDS = (881000001, 881000002)


@pytest.fixture()
def isolated_pipeline(monkeypatch):
    database = os.environ.get("WOO_DB_NAME_OVERRIDE", "")
    assert database.startswith("woo_return_loss_test_"), "Use an empty isolated test database"
    connection = db.connect()
    try:
        assert connection.execute("SELECT current_database()").fetchone()[0] == database
        assert connection.execute("SELECT COUNT(*) FROM sites WHERE url NOT LIKE '%.invalid'").fetchone()[0] == 0
        connection.execute("DELETE FROM sync_task_outbox")
        connection.execute("DELETE FROM sync_runs")
        for site_id in SITE_IDS:
            connection.execute(
                "INSERT INTO sites(id,url,consumer_key,consumer_secret,country) "
                "VALUES (?,?,? ,?,'PL')",
                (site_id, f"https://pytest-{site_id}.invalid", "test-key", "test-secret"),
            )
        connection.commit()
    finally:
        connection.close()
    monkeypatch.setattr(sync_service, "publish_pending_outbox", lambda **kw: 0)
    monkeypatch.setattr(sync_tasks, "publish_pending_outbox", lambda **kw: 0)
    monkeypatch.setenv("WOO_SYNC_POST_COMMIT_ACTIONS_ENABLED", "0")
    monkeypatch.setattr(
        requests.sessions.Session, "request",
        lambda *args, **kwargs: pytest.fail("Real network calls forbidden"),
    )
    yield
    connection = db.connect()
    try:
        connection.execute("DELETE FROM sync_task_outbox")
        connection.execute("DELETE FROM sync_runs")
        for site_id in SITE_IDS:
            connection.execute("DELETE FROM orders WHERE source=?", (f"https://pytest-{site_id}.invalid",))
            connection.execute("DELETE FROM sites WHERE id=?", (site_id,))
        connection.commit()
    finally:
        connection.close()


def test_redirect_loop_finishes_run_without_blocking_other_sites(monkeypatch, isolated_pipeline):
    def redirect_loop(*args, **kwargs):
        raise requests.TooManyRedirects("Exceeded 30 redirects")

    monkeypatch.setattr(requests.sessions.Session, "get", redirect_loop)
    status, created = sync_service.start_sync(
        mode="quick", created_by="pytest:redirect-loop", site_ids=SITE_IDS, publish=False
    )
    assert created is True
    run_id = status["run_id"]
    result = sync_tasks.fetch_page.run({"run_id": run_id, "site_id": SITE_IDS[0], "page": 1})
    assert "redirect loop" in result["error"]
    status = sync_service.get_run_status(run_id)
    sites = {site["site_id"]: site for site in status["sites"]}
    assert sites[SITE_IDS[0]]["status"] == "error"
    assert sites[SITE_IDS[0]]["retry_count"] == 0
    assert sites[SITE_IDS[0]]["failure_kind"] == "redirect"
    assert sites[SITE_IDS[0]]["next_check_at"]
    assert sites[SITE_IDS[1]]["status"] == "queued"
    assert status["completed_sites"] == 1

    class Response:
        status_code = 200
        headers = {"X-WP-TotalPages": "1"}

        def json(self):
            return []

    monkeypatch.setattr(requests.sessions.Session, "get", lambda *a, **k: Response())
    result = sync_tasks.fetch_page.run({"run_id": run_id, "site_id": SITE_IDS[1], "page": 1})
    assert result["orders"] == 0
    connection = db.connect()
    try:
        row = connection.execute(
            "SELECT payload FROM sync_task_outbox WHERE dedupe_key=?",
            (f"write:{run_id}:{SITE_IDS[1]}:1",),
        ).fetchone()
        payload = sync_service._json(row["payload"], {})
    finally:
        connection.close()
    sync_tasks._write_page_transaction(payload)

    status = sync_service.get_run_status(run_id)
    assert status["status"] == "error"
    assert status["outcome"] == "partial"
    assert status["succeeded_sites"] == 1
    assert status["unavailable_sites"] == 1
    assert status["completed_sites"] == 2
    assert next(site for site in status["sites"] if site["site_id"] == SITE_IDS[1])["status"] == "success"
    assert sync_service.active_run() is None
    recovered = sync_service.recover_stale_work()
    assert recovered["dispatches"] == 0
    assert recovered["runs"] == 0
    next_status, created = sync_service.start_sync(
        mode="quick", created_by="pytest:next-sync", site_ids=[SITE_IDS[1]], publish=False
    )
    assert created is True
    assert next_status["run_id"] != run_id


def test_scheduled_sync_defers_outage_without_incrementing_failure_count(isolated_pipeline):
    status, _ = sync_service.start_sync(
        mode="auto", created_by="pytest:first-failure", site_ids=[SITE_IDS[0]], publish=False
    )
    sync_service.mark_site_error(status["run_id"], SITE_IDS[0], 1, "DNS failure", failure_kind="dns")
    status, created = sync_service.start_sync(
        mode="auto", created_by="celery-beat:auto", site_ids=SITE_IDS, publish=False
    )
    assert created
    assert status["completed_sites"] == 1
    connection = db.connect()
    try:
        assert load_site_health(connection, SITE_IDS)[SITE_IDS[0]]["failure_count"] == 1
        assert connection.execute(
            "SELECT count(*) FROM sync_task_outbox WHERE payload->>'run_id'=?",
            (status["run_id"],),
        ).fetchone()[0] == 1
        assert connection.execute(
            "SELECT count(*) FROM sync_page_dispatches WHERE run_id=? AND site_id=?",
            (status["run_id"], SITE_IDS[0]),
        ).fetchone()[0] == 0
    finally:
        connection.close()
    sync_service.cancel_sync(status["run_id"], requested_by="pytest")
    status, created = sync_service.start_sync(
        mode="auto", created_by="celery-beat:auto", site_ids=[SITE_IDS[0]], publish=False
    )
    assert created
    assert status["status"] == "error"
    assert sync_service.active_run() is None
    # A manual quick sync can check a store immediately; no switch is needed.
    status, created = sync_service.start_sync(
        mode="quick", created_by="pytest:manual-check", site_ids=[SITE_IDS[0]], publish=False
    )
    assert created
    assert status["sites"][0]["status"] == "queued"


def test_due_recheck_recovers_same_site_and_catches_up_orders(monkeypatch, isolated_pipeline):
    site_id = SITE_IDS[0]
    source = f"https://pytest-{site_id}.invalid"

    def order(woo_id, stamp):
        return {"id": woo_id, "order_key": f"pytest-outage-{woo_id}", "status": "processing",
                "currency": "PLN", "date_created": stamp, "date_modified": stamp,
                "total": "10.00", "total_tax": "0", "shipping_total": "0", "discount_total": "0",
                "billing": {}, "shipping": {}, "meta_data": [], "line_items": [], "source": source}

    baseline = order(700001, "2026-09-29T10:00:00")
    recovered_order = order(700002, "2026-10-02T10:00:00")
    connection = db.connect()
    try:
        upsert_orders_in_transaction([baseline], connection)
        connection.execute("UPDATE sites SET last_sync='2026-09-29 10:00:00' WHERE id=?", (site_id,))
        connection.commit()
    finally:
        connection.close()
    status, _ = sync_service.start_sync(
        mode="auto", created_by="pytest:outage", site_ids=[site_id], publish=False
    )
    sync_service.mark_site_error(status["run_id"], site_id, 1, "DNS failure", failure_kind="dns")
    connection = db.connect()
    try:
        # Duplicate failure delivery does not extend the hold or count twice.
        first_health = load_site_health(connection, [site_id])[site_id]
    finally:
        connection.close()
    sync_service.mark_site_error(status["run_id"], site_id, 1, "DNS failure", failure_kind="dns")
    connection = db.connect()
    try:
        assert load_site_health(connection, [site_id])[site_id]["next_check_at"] == first_health["next_check_at"]
        assert connection.execute("SELECT last_sync FROM sites WHERE id=?", (site_id,)).fetchone()[0] == "2026-09-29 10:00:00"
        connection.execute("UPDATE sync_site_health SET next_check_at=CURRENT_TIMESTAMP-interval '1 minute' WHERE site_id=?", (site_id,))
        connection.commit()
    finally:
        connection.close()
    monkeypatch.setattr(sync_tasks, "_auto_due", lambda: (False, {"interval": 3600}))
    checked = sync_tasks.schedule_auto.run()
    assert checked["created"] is True
    run_id = checked["run_id"]
    status = sync_service.get_run_status(run_id)
    assert status["created_by"] == "celery-beat:site-recovery"
    assert [site["site_id"] for site in status["sites"]] == [site_id]
    params_seen = []

    class Response:
        status_code = 200
        headers = {"X-WP-TotalPages": "1"}

        def json(self):
            return [recovered_order]

    def restored_get(_session, url, **kwargs):
        if url.endswith("/notes"):
            empty = Response()
            empty.json = lambda: []
            return empty
        params_seen.append(kwargs["params"])
        return Response()

    monkeypatch.setattr(requests.sessions.Session, "get", restored_get)
    sync_tasks.fetch_page.run({"run_id": run_id, "site_id": site_id, "page": 1})
    assert params_seen[0]["modified_after"].startswith("2026-09-29T09:50:00")
    connection = db.connect()
    try:
        raw_payload = connection.execute(
            "SELECT payload FROM sync_task_outbox WHERE dedupe_key=?", (f"write:{run_id}:{site_id}:1",)
        ).fetchone()["payload"]
        payload = sync_service._json(raw_payload, {})
    finally:
        connection.close()
    sync_tasks._write_page_transaction(payload)
    assert sync_service.get_run_status(run_id)["status"] == "success"
    connection = db.connect()
    try:
        assert load_site_health(connection, [site_id]) == {}
        ids = [row[0] for row in connection.execute("SELECT id FROM orders WHERE source=? ORDER BY id", (source,)).fetchall()]
        assert ids == [make_oid(site_id, 700001), make_oid(site_id, 700002)]
    finally:
        connection.close()


def test_disabled_auto_sync_does_not_restart_outage_rechecks(monkeypatch, isolated_pipeline):
    monkeypatch.setattr(sync_tasks, "_auto_due", lambda: (False, {"reason": "disabled"}))
    monkeypatch.setattr(sync_tasks, "due_site_rechecks", lambda *_: pytest.fail("Disabled schedule must not recheck"))
    assert sync_tasks.schedule_auto.run()["created"] is False
