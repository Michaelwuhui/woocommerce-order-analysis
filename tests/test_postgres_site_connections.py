"""Real locks/FKs, restricted to a disposable prefixed PostgreSQL database.

The database must already contain the deployed schema, with no production rows.
All HTTP and broker publication are forbidden or replaced with no-ops.
"""
from concurrent.futures import ThreadPoolExecutor
import json
import os
import queue
import time
import uuid

import pytest
import requests

pytest.importorskip("celery")

import clean_sync
import db_backend as db
from product_clone_jobs import enqueue_clone_job
from site_connection_service import SiteConnectionError, lock_active_sites, remove_connection
import sync_service
import sync_tasks


pytestmark = pytest.mark.skipif(not db.is_postgres_backend(), reason="isolated PostgreSQL required")
SITE_ID = 991247007
OTHER_ID = 991247037
SITE_URL = "https://site-connection-contract.invalid"


@pytest.fixture
def isolated(monkeypatch):
    name = os.environ.get("WOO_DB_NAME_OVERRIDE", "")
    assert name.startswith("woo_return_loss_test_"), "Refusing non-isolated database"

    def forbidden(*args, **kwargs):
        pytest.fail("External HTTP must never run in the site-removal contract")

    monkeypatch.setattr(requests.sessions.Session, "request", forbidden)
    monkeypatch.setattr(sync_service, "publish_pending_outbox", lambda **_kw: 0)
    monkeypatch.setattr(sync_tasks, "publish_pending_outbox", lambda **_kw: 0)
    run_id = str(uuid.uuid4())
    conn = sync_service.get_connection()
    assert conn.execute("SELECT current_database()").fetchone()[0] == name
    try:
        for sid, url in ((SITE_ID, SITE_URL), (OTHER_ID, "https://site-connection-other.invalid")):
            conn.execute("INSERT INTO sites(id,url,consumer_key,consumer_secret,manager,country,cod_on_hold_is_shipped) "
                         "VALUES(?,?,?,?,?,?,?)", (sid, url, "fixture-ck", "fixture-cs", "Fixture Owner", "PL", 1))
        for oid in (f"{SITE_ID}-101", f"{SITE_ID}-102"):
            conn.execute("INSERT INTO orders(id,source,status,total,currency) VALUES(?,?,'completed',10,'PLN')", (oid, SITE_URL))
        conn.execute("INSERT INTO sync_runs(run_id,mode,status,total_sites,completed_sites) VALUES(?,'quick','success',1,1)", (run_id,))
        conn.execute("INSERT INTO sync_site_progress(run_id,site_id,status,current_page,written_count) VALUES(?,?,'success',1,2)", (run_id, SITE_ID))
        conn.execute("INSERT INTO sync_page_dispatches(run_id,site_id,page,status,content_hash) VALUES(?,?,1,'completed','history-hash')", (run_id, SITE_ID))
        conn.execute("""INSERT INTO sync_page_receipts(run_id,site_id,page,content_hash,fetched_count,
            written_count,changed_count,is_last_page,post_commit_status)
            VALUES(?,?,1,'history-hash',2,2,2,TRUE,'completed')""", (run_id, SITE_ID))
        conn.execute("INSERT INTO sync_task_outbox(dedupe_key,queue_name,task_name,payload,status) "
                     "VALUES(?,'sync_fetch','woo_sync.fetch_page',?::jsonb,'published')",
                     (f"site-contract:{run_id}", json.dumps({"run_id": run_id, "site_id": SITE_ID, "page": 1})))
        conn.commit()
    finally:
        conn.close()
    yield {"run_id": run_id, "connect": sync_service.get_connection}
    conn = sync_service.get_connection()
    try:
        conn.execute("DELETE FROM product_clone_jobs WHERE source_site_id IN (?,?) OR target_site_id IN (?,?)",
                     (SITE_ID, OTHER_ID, SITE_ID, OTHER_ID))
        conn.execute("DELETE FROM sync_task_outbox WHERE dedupe_key=?", (f"site-contract:{run_id}",))
        conn.execute("DELETE FROM sync_runs WHERE run_id=?", (run_id,))
        conn.execute("DELETE FROM orders WHERE source=?", (SITE_URL,))
        conn.execute("DELETE FROM inv_site_sync_audit WHERE site_id IN (?,?)", (SITE_ID, OTHER_ID))
        conn.execute("DELETE FROM sites WHERE id IN (?,?)", (SITE_ID, OTHER_ID))
        conn.commit()
    finally:
        conn.close()


def _snapshot(c, run_id):
    return {
        "orders": [dict(r) for r in c.execute("SELECT * FROM orders WHERE source=? ORDER BY id", (SITE_URL,))],
        **{table: [dict(r) for r in c.execute(f"SELECT * FROM {table} WHERE run_id=? ORDER BY 1", (run_id,))]
           for table in ("sync_runs", "sync_site_progress", "sync_page_dispatches", "sync_page_receipts")},
    }


def test_real_foreign_keys_and_all_history_survive_connection_removal(isolated):
    c = isolated["connect"]()
    try:
        before = _snapshot(c, isolated["run_id"])
        site_before = dict(c.execute("SELECT * FROM sites WHERE id=?", (SITE_ID,)).fetchone())
        c.commit()
        assert remove_connection(c, SITE_ID)["success"] is True
        site_after = dict(c.execute("SELECT * FROM sites WHERE id=?", (SITE_ID,)).fetchone())
        for key in ("id", "url", "manager", "country", "cod_on_hold_is_shipped"):
            assert site_after[key] == site_before[key]
        assert site_after["consumer_key"] == site_after["consumer_secret"] == ""
        assert site_after["api_status"] == "archived"
        assert _snapshot(c, isolated["run_id"]) == before
        assert len(before["orders"]) == 2
        assert sync_service._load_sites(c, [SITE_ID]) == []
        c.rollback()
        with pytest.raises(ValueError, match="unavailable"):
            sync_service.start_sync(mode="quick", created_by="fixture", site_ids=[SITE_ID], publish=False)
    finally:
        c.close()


def test_task_creation_key_share_prevents_archive_nowait(isolated):
    creator = isolated["connect"]()
    remover = isolated["connect"]()
    try:
        lock_active_sites(creator, [SITE_ID])
        with pytest.raises(SiteConnectionError) as caught:
            remove_connection(remover, SITE_ID)
        assert caught.value.code == "SITE_BUSY"
        assert caught.value.details["reasons"] == ["site_configuration_lock"]
        creator.rollback()
        assert remove_connection(remover, SITE_ID)["archived"] is True
    finally:
        creator.close()
        remover.close()


def _enqueue(c):
    return enqueue_clone_job(c, source_site_id=SITE_ID, target_site_id=OTHER_ID, product_ids=[101],
        options={}, target_url="https://site-connection-other.invalid", created_by_id="", created_by_name="Fixture")


def test_archive_row_lock_makes_enqueue_wait_then_reject_fresh_marker(isolated):
    remover = isolated["connect"]()
    observer = isolated["connect"]()
    pids = queue.Queue()
    try:
        remover.execute("SELECT id FROM sites WHERE id=? FOR UPDATE", (SITE_ID,)).fetchone()
        remover.execute("UPDATE sites SET api_status='archived',consumer_key='',consumer_secret='' WHERE id=?", (SITE_ID,))

        def enqueue_after_old_read():
            c = isolated["connect"]()
            try:
                assert c.execute("SELECT api_status FROM sites WHERE id=?", (SITE_ID,)).fetchone()[0] != "archived"
                pids.put(c.execute("SELECT pg_backend_pid()").fetchone()[0])
                return _enqueue(c)
            finally:
                c.close()

        with ThreadPoolExecutor(max_workers=1) as pool:
            future = pool.submit(enqueue_after_old_read)
            pid = pids.get(timeout=5)
            deadline = time.monotonic() + 5
            waiting = False
            while time.monotonic() < deadline:
                row = observer.execute("SELECT wait_event_type FROM pg_stat_activity WHERE pid=?", (pid,)).fetchone()
                observer.commit()
                if row and row[0] == "Lock":
                    waiting = True
                    break
                time.sleep(0.02)
            try:
                assert waiting, "Clone enqueue must wait on archive's config row lock"
            finally:
                remover.commit()
            with pytest.raises(SiteConnectionError) as caught:
                future.result(timeout=5)
            assert caught.value.code == "SITE_ARCHIVED"
        assert observer.execute("SELECT COUNT(*) FROM product_clone_jobs WHERE source_site_id=?", (SITE_ID,)).fetchone()[0] == 0
    finally:
        remover.rollback()
        remover.close()
        observer.close()


def test_fetch_session_advisory_lock_prevents_removal(isolated):
    fetcher = isolated["connect"]()
    remover = isolated["connect"]()
    try:
        assert sync_tasks._site_lock(fetcher, SITE_ID) is True
        with pytest.raises(SiteConnectionError) as caught:
            remove_connection(remover, SITE_ID)
        assert caught.value.details["reasons"] == ["order_sync_lock"]
    finally:
        sync_tasks._site_unlock(fetcher, SITE_ID)
        fetcher.close()
        remover.close()


def test_archived_messages_cannot_fetch_write_clean_or_post_commit(isolated, monkeypatch):
    c = isolated["connect"]()
    try:
        remove_connection(c, SITE_ID)
        before = _snapshot(c, isolated["run_id"])
    finally:
        c.close()
    payload = {"run_id": isolated["run_id"], "site_id": SITE_ID, "page": 2,
               "orders": [], "notes": [], "content_hash": sync_tasks._canonical_hash([], []),
               "is_last_page": True, "total_pages": 2}
    monkeypatch.setattr(sync_tasks, "upsert_orders_in_transaction", lambda *_: pytest.fail("No stale order writes"))
    assert sync_tasks._claim_fetch(payload, "stale") is None
    assert sync_tasks._claim_write(payload["run_id"], SITE_ID, 2, "stale") is False
    assert sync_tasks._write_page_transaction(payload)["archived"] is True
    assert sync_tasks._claim_post_commit(payload["run_id"], SITE_ID, 1) is None
    assert clean_sync.run_clean_site(payload)["archived"] is True
    assert sync_tasks._queue_page_write(payload, orders=[], notes=[], total_pages=2, is_last_page=True) == payload["content_hash"]
    c = isolated["connect"]()
    try:
        assert _snapshot(c, isolated["run_id"]) == before
        with pytest.raises(SiteConnectionError) as caught:
            _enqueue(c)
        assert caught.value.code == "SITE_ARCHIVED"
    finally:
        c.close()


def test_restoring_same_id_does_not_replay_terminal_error_run_messages(isolated, monkeypatch):
    c = isolated["connect"]()
    run_id = isolated["run_id"]
    try:
        c.execute("UPDATE sync_runs SET status='error' WHERE run_id=?", (run_id,))
        c.execute("UPDATE sync_site_progress SET status='error' WHERE run_id=?", (run_id,))
        c.execute("INSERT INTO sync_page_dispatches(run_id,site_id,page,status) VALUES(?,?,2,'error')", (run_id, SITE_ID))
        c.commit()
        remove_connection(c, SITE_ID)
        c.execute("UPDATE sites SET api_status='unknown',consumer_key='restored-ck',consumer_secret='restored-cs' WHERE id=?", (SITE_ID,))
        c.commit()
        before = _snapshot(c, run_id)
    finally:
        c.close()
    payload = {"run_id": run_id, "site_id": SITE_ID, "page": 2, "orders": [], "notes": [],
               "content_hash": sync_tasks._canonical_hash([], []), "total_pages": 2, "is_last_page": True}
    monkeypatch.setattr(sync_tasks, "upsert_orders_in_transaction", lambda *_: pytest.fail("No stale writes after restore"))
    assert sync_tasks._claim_fetch(payload, "old") is None
    assert sync_tasks._claim_write(run_id, SITE_ID, 2, "old") is False
    assert sync_tasks._write_page_transaction(payload)["written"] == 0
    sync_tasks._queue_page_write(payload, orders=[], notes=[], total_pages=2, is_last_page=True)
    c = isolated["connect"]()
    try:
        assert _snapshot(c, run_id) == before
    finally:
        c.close()
