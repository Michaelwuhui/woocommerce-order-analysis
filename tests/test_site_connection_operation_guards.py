"""Retained site identities cannot remain remote-operation targets.

All site credentials are synthetic and every HTTP path is replaced or rejected.
The archive marker is deliberately tested without clearing credentials first.
"""
import sqlite3
from contextlib import nullcontext
from unittest.mock import Mock

import pytest

import auto_confirm
import delivery_automation as delivery
import fulfillment_woocommerce as fulfillment
import inv_push
import product_clone_worker as clone_worker
from product_clone_jobs import claim_clone_job, enqueue_clone_job, get_clone_job, init_product_clone_jobs
from site_connection_service import SiteConnectionError
from stock_sync_fixtures import system
from test_product_manager_catalog import system as catalog_system
from test_product_manager_catalog_edit import system as edit_system, preview, update
import test_inventory_auto_push as inventory_fixture
from test_stock_sync_mapping import mapped_system, scan, payload


def archive(connect, site_id):
    conn = connect()
    columns = {column[0] for column in conn.execute("SELECT * FROM sites WHERE 1=0").description}
    if "api_status" not in columns:
        conn.execute("ALTER TABLE sites ADD COLUMN api_status TEXT")
    conn.execute("UPDATE sites SET api_status='archived' WHERE id=?", (site_id,))
    conn.commit()
    conn.close()


def test_catalog_archive_hides_only_archived_and_forced_scan_reads_nothing(catalog_system):
    s = catalog_system
    archive(s["connect"], 1)
    sites = s["client"].get("/api/product-manager/catalog-sites").get_json()["sites"]
    assert {site["id"] for site in sites} == {2, 3, 4}
    assert next(site for site in sites if site["id"] == 3)["configured"] is False
    response = s["client"].get("/api/product-manager/catalog-page?site_id=1&query_mode=all")
    assert response.status_code == 409 and response.get_json()["code"] == "site_archived"
    assert s["calls"] == []
    conn = s["connect"]()
    assert conn.execute("SELECT consumer_key FROM sites WHERE id=1").fetchone()[0]
    conn.close()


def test_catalog_editor_and_clone_reject_retained_archive_without_remote_access(edit_system):
    s = edit_system
    before = preview(s).get_json()
    s["remote"].calls.clear()
    archive(s["connect"], 1)
    config = s["client"].get("/api/product-manager/catalog-edit-config").get_json()
    assert 1 not in {site["id"] for site in config["sites"]}
    assert preview(s).get_json()["code"] == "SITE_ARCHIVED"
    assert update(s, {"regular_price": "15"}, snapshot=before).get_json()["code"] == "SITE_ARCHIVED"
    response = s["client"].post("/api/product-manager/catalog-clone", json={
        "source_site_id": 1, "target_site_id": 2, "product_ids": [10],
        "include_variations": True, "include_images": False,
        "status_on_target": "draft", "collision_mode": "skip_existing",
    }, headers={"X-PM-CSRF": s["csrf"]})
    assert response.status_code == 409
    assert s["remote"].calls == [] and s["jobs"] == []


def test_catalog_editor_rechecks_archive_after_preflight_before_put(edit_system):
    s = edit_system
    before = preview(s).get_json()
    s["remote"].preflight_hook = lambda: archive(s["connect"], 1)
    response = update(s, {"regular_price": "15"}, snapshot=before)
    assert response.status_code == 409
    assert not any(method == "PUT" for method, _, _ in s["remote"].calls)
    assert s["remote"].items["https://child.test/wp-json/wc/v3/products/10"]["regular_price"] == "12.00"


@pytest.mark.parametrize("archived_id", [1, 2])
def test_legacy_queued_clone_rejects_archived_endpoint_and_keeps_job_history(tmp_path, archived_id):
    conn = sqlite3.connect(tmp_path / "clone.db")
    conn.row_factory = sqlite3.Row
    conn.executescript("""CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT,manager TEXT,
        consumer_key TEXT,consumer_secret TEXT,product_master_id INTEGER,api_status TEXT);
        INSERT INTO sites VALUES(1,'https://source.test','Admin','ck','cs',NULL,NULL);
        INSERT INTO sites VALUES(2,'https://target.test','Admin','ck','cs',NULL,NULL);""")
    init_product_clone_jobs(conn)
    job = enqueue_clone_job(conn, source_site_id=1, target_site_id=2, target_url="https://target.test", product_ids=[10],
                            options={}, created_by_id=1, created_by_name="Admin")
    job = claim_clone_job(conn, "offline-worker")
    conn.execute("UPDATE sites SET api_status='archived' WHERE id=?", (archived_id,))
    conn.commit()
    clone_one, resolver = Mock(), Mock()
    clone_worker.process_clone_job(conn, job, clone_one=clone_one, resolve_site=resolver)
    assert not clone_one.called and not resolver.called
    saved = get_clone_job(conn, job["id"])
    assert saved["status"] == "failed" and conn.execute("SELECT COUNT(*) FROM sites").fetchone()[0] == 2
    conn.close()


@pytest.fixture
def inventory():
    fixture = inventory_fixture.InventoryAutoPushTests()
    fixture.setUp()
    fixture.db.execute("ALTER TABLE sites ADD COLUMN api_status TEXT")
    yield fixture.db
    fixture.tearDown()


def test_inventory_archive_excluded_from_quota_scheduler_and_manual_execution(inventory, monkeypatch):
    conn = inventory
    monkeypatch.setattr(inv_push.inv_allocator, "candidate_warehouses", lambda *args: [{"warehouse_id": 1}])
    for identity in (1, 2):
        inv_push.update_site_sync_config(conn, identity, {"mode": "observe"}, (1, "admin"))
    conn.execute("UPDATE sites SET api_status='archived' WHERE id=2")
    conn.commit()
    remaining = inv_push.compute_site_stock(conn, 1, use_sync_strategy=True)[0]
    assert remaining["allocated_sku"] == 9 and remaining["allocation_participants"] == 1
    assert inv_push.compute_site_stock(conn, 2) == []
    assert 2 not in inv_push.scheduler_site_ids(conn)
    remote = Mock(side_effect=AssertionError("Archived connection reached Woo"))
    monkeypatch.setattr(inv_push, "_sync_one_stock", remote)
    assert inv_push.execute_site_sync(conn, 2, trigger_type="manual")["status"] == "skipped"
    assert inv_push.push_site(conn, 2, dry_run=False)["fatal"]
    with pytest.raises(ValueError):
        inv_push.update_site_sync_config(conn, 2, {"mode": "live"}, (1, "admin"))
    with pytest.raises(SiteConnectionError):
        inv_push._acquire_site_lock(conn, 2)
    conn.rollback()
    assert not remote.called
    assert conn.execute("SELECT COUNT(*) FROM inv_push_runs").fetchone()[0] == 0
    assert conn.execute("SELECT COUNT(*) FROM inv_push_locks").fetchone()[0] == 0


def test_inventory_cached_credentials_checked_before_stock_put(inventory, monkeypatch):
    conn = inventory
    monkeypatch.setattr(inv_push.inv_allocator, "candidate_warehouses", lambda *args: [{"warehouse_id": 1}])
    # Real push orchestration passes a current-connection check into each leaf.
    def remote(*args, write_check=None, **kwargs):
        conn.execute("UPDATE sites SET api_status='archived' WHERE id=1")
        conn.commit()
        assert write_check and write_check()
        return False, "站点连接已删除", None, None
    monkeypatch.setattr(inv_push, "_sync_one_stock", remote)
    result = inv_push.push_site(conn, 1, dry_run=False, only_changed=False)
    assert result["error"] and result["ok"] == 0


def test_inventory_preflight_archive_never_reaches_actual_put(inventory, monkeypatch):
    conn = inventory
    def preflight(*args):
        conn.execute("UPDATE sites SET api_status='archived' WHERE id=1")
        conn.commit()
        return {"manage_stock": True, "stock_quantity": 1}, None
    write = Mock(side_effect=AssertionError("Archived stock preflight sent PUT"))
    monkeypatch.setattr(inv_push, "_get_stock_state", preflight)
    monkeypatch.setattr(inv_push, "_put_unmanaged_stock", write)
    monkeypatch.setattr("stock_sync_guard.legacy_write", lambda *args: nullcontext())
    status, _, _, error = inv_push._sync_one_stock("https://one.test", "ck", "cs",
        {"wc_product_id": 101, "wc_variation_id": 0, "publishable": 9},
        write_check=lambda: "archived" if inv_push.is_site_archived(conn, 1) else None)
    assert status == "error" and error == "archived" and not write.called


def test_stock_options_and_future_execution_reject_archive_but_get_history_survives(system):
    s = system
    s.seed(1, 1)
    plan = s.plan("reference_status")
    job = s.execute(plan)
    reads, writes = len(s.http.calls), len(s.http.puts)
    archive(s.connect, 2)
    options = s.call("/options").get_json()
    assert 2 not in {site["id"] for site in options["target_sites"]}
    assert 2 not in {site["id"] for site in options["reference_sites"]}
    assert s.call("/plans/" + plan["id"]).status_code == 200
    assert s.call("/jobs/" + job["id"]).status_code == 200
    snapshots = s.sql("SELECT id FROM stock_sync_catalog_snapshots")
    assert snapshots and all(s.call("/catalog-scans/" + snap["id"]).status_code == 200 for snap in snapshots)
    rejected = s.call("/jobs", "POST", {"plan_id": plan["id"], "plan_version": plan["version"],
        "idempotency_key": "archive-attempt", "accepted_item_ids": [i["id"] for i in plan["items"]]})
    assert rejected.status_code in (403, 409)
    assert s.call("/catalog-scans", "POST", {"target_scope": {"mode": "explicit_sites", "site_ids": [2]}}).status_code in (403, 409)
    assert len(s.http.calls) == reads and len(s.http.puts) == writes
    assert len(s.sql("SELECT * FROM stock_sync_jobs")) == 1
    archive(s.connect, 1)
    assert s.call("/plans/" + plan["id"]).status_code == 200
    assert s.call("/catalog-scans", "POST", {"source_site_id": 1}).status_code == 409


def test_archived_mapping_snapshot_is_readable_but_not_writable(mapped_system):
    s = mapped_system
    report = scan(s)
    archive(s.connect, 2)
    reads = len(s.http.calls)
    response = s.call("/mapping-scans/" + report["id"])
    assert response.status_code == 200
    archived_items = [item for item in response.get_json()["items"] if item["site_id"] == 2]
    assert archived_items and all(item["writable"] is False for item in archived_items)
    response = s.call("/mapping-scans/" + report["id"] + "/confirm", "POST", payload(report))
    assert response.status_code in (403, 409)
    assert len(s.http.calls) == reads and not s.http.puts


@pytest.fixture
def retained_connection():
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.executescript("""CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT,country TEXT,
        consumer_key TEXT,consumer_secret TEXT,api_status TEXT);
        INSERT INTO sites VALUES(1,'https://retained.test','HU','ck','cs',NULL);
        CREATE TABLE oms_order_fulfillment_state(order_id TEXT,completion_sync_status TEXT);""")
    conn.commit()
    yield conn
    conn.close()


def test_auto_confirm_cached_site_cannot_put_after_archive(retained_connection, monkeypatch):
    conn = retained_connection
    cached = conn.execute("SELECT * FROM sites WHERE id=1").fetchone()
    conn.execute("UPDATE sites SET api_status='archived' WHERE id=1")
    remote = Mock()
    monkeypatch.setattr(auto_confirm.requests, "put", remote)
    assert auto_confirm._load_sites(conn) == {}
    ok, reason = auto_confirm._complete_remote(cached, "1-123", conn=conn)
    assert not ok and "删除" in reason and not remote.called
    # Country and identity remain available to historical order calculations.
    assert conn.execute("SELECT country FROM sites WHERE id=1").fetchone()[0] == "HU"


@pytest.mark.parametrize("method", ["GET", "PUT", "POST"])
def test_fulfillment_cached_site_blocks_remote_and_email_paths(retained_connection, monkeypatch, method):
    conn = retained_connection
    cached = conn.execute("SELECT * FROM sites WHERE id=1").fetchone()
    conn.execute("UPDATE sites SET api_status='archived' WHERE id=1")
    remote, email = Mock(), Mock()
    monkeypatch.setattr(fulfillment.requests, "request", remote)
    monkeypatch.setattr(fulfillment.requests, "post", email)
    with pytest.raises(fulfillment.WooError) as raised:
        fulfillment._request(method, "https://retained.test/wp-json/wc/v3/orders/123", cached, conn=conn)
    assert raised.value.code == "site_archived"
    with pytest.raises(fulfillment.WooError):
        fulfillment._notify_shipment(conn, cached, {}, {}, "custom", True)
    assert not remote.called and not email.called


class DeliveryConnection:
    """SQLite plus the two PostgreSQL advisory-lock calls used by delivery."""
    def __init__(self, conn):
        self.conn = conn
    def execute(self, sql, args=()):
        if "pg_try_advisory_lock" in sql or "pg_advisory_unlock" in sql:
            return self.conn.execute("SELECT 1")
        return self.conn.execute(sql, args)
    def __getattr__(self, name):
        return getattr(self.conn, name)


@pytest.mark.parametrize("after_preflight", [False, True])
def test_delivery_archive_rejects_before_ledger_or_cached_put(retained_connection, monkeypatch, after_preflight):
    conn = retained_connection
    wrapped = DeliveryConnection(conn)
    monkeypatch.setattr(delivery, "eligible", lambda *_: {"source": "https://retained.test"})
    begin = Mock(return_value={"operation_id": "offline-op", "status": "pending", "should_execute": True})
    transition, phases = Mock(), []
    monkeypatch.setattr(delivery, "begin_operation", begin)
    monkeypatch.setattr(delivery, "_record_phase", lambda _c, _op, phase: phases.append(phase))
    monkeypatch.setattr(delivery, "_transition", transition)
    session = Mock()
    if after_preflight:
        def preflight(*args):
            conn.execute("UPDATE sites SET api_status='archived' WHERE id=1")
            conn.commit()
            return {"id": 123, "status": "on-hold", "_verified_url": "https://retained.test/wp-json/wc/v3/orders/123"}
        monkeypatch.setattr(delivery, "_remote_order", preflight)
    else:
        conn.execute("UPDATE sites SET api_status='archived' WHERE id=1")
        conn.commit()
    result = delivery.process_delivered_order(wrapped, "1-123", session)
    assert result["result"] == "site_archived" and not session.put.called
    assert "write_started" not in phases
    assert begin.called is after_preflight
    if after_preflight:
        assert transition.call_args.kwargs["error"] == "site_connection_archived_before_write"
