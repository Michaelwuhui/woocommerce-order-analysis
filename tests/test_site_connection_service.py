"""Real SQLite transactions; no application import or external I/O."""
import json
import sqlite3

import pytest

from product_clone_jobs import enqueue_clone_job, init_product_clone_jobs
from site_connection_service import (
    SiteConnectionError, archived_site_ids, filter_active_sites, is_site_archived,
    lock_active_sites, remove_connection,
)
from sync_service import _load_sites


@pytest.fixture
def conn():
    c = sqlite3.connect(":memory:")
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    c.executescript("""
        CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT,consumer_key TEXT,consumer_secret TEXT,
            manager TEXT,country TEXT,cod_on_hold_is_shipped INTEGER,product_master_id INTEGER,
            api_status TEXT,last_sync TEXT);
        INSERT INTO sites VALUES(7,'https://historical.invalid','fixture-ck','fixture-cs',
            'Fixture Owner','PL',1,3,'ok','2026-10-01');
        INSERT INTO sites VALUES(37,'https://active.invalid','other-ck','other-cs',
            'Other Owner','CZ',0,3,'unknown','2026-10-02');
        CREATE TABLE product_masters(id INTEGER PRIMARY KEY,url TEXT);
        INSERT INTO product_masters VALUES(3,'https://master.invalid');
        CREATE TABLE orders(id TEXT PRIMARY KEY,source TEXT,status TEXT,total REAL);
        INSERT INTO orders VALUES('7-101','https://historical.invalid','on-hold',100);
        INSERT INTO orders VALUES('7-102','https://historical.invalid','completed',200);
        CREATE TABLE sync_runs(run_id TEXT PRIMARY KEY,status TEXT);
        INSERT INTO sync_runs VALUES('old','success');
        CREATE TABLE sync_site_progress(run_id TEXT,site_id INTEGER,status TEXT,
            FOREIGN KEY(site_id) REFERENCES sites(id) ON DELETE RESTRICT);
        INSERT INTO sync_site_progress VALUES('old',7,'success');
        CREATE TABLE sync_page_dispatches(run_id TEXT,site_id INTEGER,status TEXT);
        INSERT INTO sync_page_dispatches VALUES('old',7,'completed');
        CREATE TABLE sync_page_receipts(run_id TEXT,site_id INTEGER,post_commit_status TEXT);
        INSERT INTO sync_page_receipts VALUES('old',7,'completed');
        CREATE TABLE sync_task_outbox(status TEXT,payload TEXT);
        INSERT INTO sync_task_outbox VALUES('published','{"site_id":7,"run_id":"old"}');
        CREATE TABLE external_operations(site_id INTEGER,status TEXT);
        CREATE TABLE inv_push_runs(site_id INTEGER,status TEXT);
        CREATE TABLE inv_push_locks(site_id INTEGER PRIMARY KEY,lock_token TEXT);
        CREATE TABLE inv_site_sync_audit(id INTEGER PRIMARY KEY,site_id INTEGER,action TEXT,
            before_json TEXT,after_json TEXT,operator_id INTEGER,operator_name TEXT);
        CREATE TABLE stock_sync_jobs(id TEXT,plan_id TEXT,status TEXT);
        CREATE TABLE stock_sync_plans(id TEXT,request_json TEXT);
        CREATE TABLE stock_sync_plan_items(plan_id TEXT,site_id INTEGER,resource_key TEXT);
        CREATE TABLE stock_sync_resource_leases(resource_key TEXT);
        CREATE TABLE stock_sync_catalog_snapshots(id TEXT,site_id INTEGER,scope_json TEXT);
        CREATE TABLE stock_sync_work(object_id TEXT,status TEXT);
        CREATE TABLE oms_integration_jobs(aggregate_type TEXT,aggregate_id TEXT,status TEXT,job_type TEXT);
        CREATE TABLE oms_fulfillments(id TEXT,order_id TEXT,status TEXT);
    """)
    init_product_clone_jobs(c)
    yield c
    c.close()


def snapshot(c):
    return {table: [dict(r) for r in c.execute(f"SELECT * FROM {table}")]
            for table in ("orders", "sync_runs", "sync_site_progress", "sync_page_dispatches",
                          "sync_page_receipts", "sync_task_outbox")}


def test_remove_preserves_identity_orders_and_restricted_sync_history(conn):
    historical = snapshot(conn)
    old = dict(conn.execute("SELECT * FROM sites WHERE id=7").fetchone())
    other = dict(conn.execute("SELECT * FROM sites WHERE id=37").fetchone())
    result = remove_connection(conn, 7, actor_id=1, actor_name="Fixture Owner")
    site = dict(conn.execute("SELECT * FROM sites WHERE id=7").fetchone())
    assert result == {"success": True, "id": 7, "url": old["url"],
                      "archived": True, "already_archived": False}
    for field in ("id", "url", "country", "manager", "cod_on_hold_is_shipped", "last_sync"):
        assert site[field] == old[field]
    assert site["consumer_key"] == site["consumer_secret"] == ""
    assert site["product_master_id"] is None
    assert site["api_status"] == "archived"
    assert snapshot(conn) == historical
    assert dict(conn.execute("SELECT * FROM sites WHERE id=37").fetchone()) == other
    audit = dict(conn.execute("SELECT * FROM inv_site_sync_audit").fetchone())
    assert audit["action"] == "connection_removed"
    serialized = json.dumps(audit)
    assert "fixture-ck" not in serialized and "fixture-cs" not in serialized
    assert json.loads(audit["before_json"])["configured"] is True
    assert json.loads(audit["after_json"])["configured"] is False
    assert remove_connection(conn, 7)["already_archived"] is True
    assert conn.execute("SELECT COUNT(*) FROM inv_site_sync_audit").fetchone()[0] == 1


@pytest.mark.parametrize("statement,reason", [
    ("UPDATE sync_runs SET status='running'", "order_sync"),
    ("UPDATE sync_site_progress SET status='fetching'", "order_sync"),
    ("UPDATE sync_page_dispatches SET status='writing'", "sync_pages"),
    ("UPDATE sync_page_receipts SET post_commit_status='error'", "sync_post_commit"),
    ("UPDATE sync_task_outbox SET status='publishing'", "sync_outbox"),
    ("INSERT INTO external_operations VALUES(7,'reconciliation_required')", "external_operation"),
    ("INSERT INTO inv_push_runs VALUES(7,'running')", "inventory_push"),
    ("INSERT INTO inv_push_locks VALUES(7,'fixture-token')", "inventory_push_lock"),
    ("INSERT INTO stock_sync_resource_leases VALUES('endpoint:historical.invalid')", "product_write_lease"),
    ("INSERT INTO oms_integration_jobs VALUES('order','7-101','retry','COMPLETE_WOOCOMMERCE_ORDER')", "order_fulfillment_job"),
    ("INSERT INTO oms_fulfillments VALUES('f','7-101','submission_unknown')", "warehouse_result_unknown"),
])
def test_busy_removal_is_rolled_back_without_losing_configuration(conn, statement, reason):
    conn.execute(statement)
    conn.commit()
    old = dict(conn.execute("SELECT * FROM sites WHERE id=7").fetchone())
    with pytest.raises(SiteConnectionError) as caught:
        remove_connection(conn, 7)
    assert caught.value.code == "SITE_BUSY" and caught.value.status == 409
    assert reason in caught.value.details["reasons"]
    assert dict(conn.execute("SELECT * FROM sites WHERE id=7").fetchone()) == old
    assert conn.execute("SELECT COUNT(*) FROM inv_site_sync_audit").fetchone()[0] == 0


def test_queued_clone_blocks_both_source_and_target(conn):
    conn.execute("INSERT INTO product_clone_jobs(id,source_site_id,target_site_id,product_ids_json,options_json) "
                 "VALUES('queued',37,7,'[10]','{}')")
    conn.commit()
    for site_id in (7, 37):
        with pytest.raises(SiteConnectionError) as caught:
            remove_connection(conn, site_id)
        assert "product_clone" in caught.value.details["reasons"]


@pytest.mark.parametrize("status", ["queued", "running", "requires_review", "cancel_requested"])
def test_stock_target_and_reference_work_blocks_removal(conn, status):
    conn.execute("INSERT INTO stock_sync_jobs VALUES('job','plan',?)", (status,))
    conn.execute("INSERT INTO stock_sync_plans VALUES('plan','{\"source_site_id\":7}')")
    conn.execute("INSERT INTO stock_sync_plan_items VALUES('plan',37,'active.invalid/products/10/variations/0')")
    conn.commit()
    with pytest.raises(SiteConnectionError) as caught:
        remove_connection(conn, 7)
    assert "stock_reference" in caught.value.details["reasons"]
    with pytest.raises(SiteConnectionError) as caught:
        remove_connection(conn, 37)
    assert "stock_sync" in caught.value.details["reasons"]


def test_actual_master_with_other_children_is_protected_but_child_is_allowed(conn):
    conn.execute("UPDATE product_masters SET url='https://historical.invalid/' WHERE id=3")
    conn.commit()
    with pytest.raises(SiteConnectionError) as caught:
        remove_connection(conn, 7)
    assert caught.value.code == "MASTER_IN_USE"
    # This site points to that Master but its own URL is unrelated.
    assert remove_connection(conn, 37)["success"] is True
    assert remove_connection(conn, 7)["success"] is True


def test_helpers_hide_only_archived_and_keep_original_rows(conn):
    rows = conn.execute("SELECT * FROM sites ORDER BY id").fetchall()
    conn.execute("UPDATE sites SET consumer_key='',consumer_secret='' WHERE id=37")
    conn.commit()
    assert filter_active_sites(conn, rows) == rows
    remove_connection(conn, 7)
    assert archived_site_ids(conn) == {7}
    assert is_site_archived(conn, 7) is True
    assert filter_active_sites(conn, rows) == [rows[1]]
    assert filter_active_sites(conn, [{"site_id": 7}, {"site_id": 37}]) == [{"site_id": 37}]
    assert [r["id"] for r in _load_sites(conn, None)] == [37]
    assert _load_sites(conn, [7]) == []
    with pytest.raises(SiteConnectionError, match="已移除"):
        lock_active_sites(conn, [7])


def test_legacy_sqlite_helper_schema_is_compatible():
    c = sqlite3.connect(":memory:")
    c.row_factory = sqlite3.Row
    assert archived_site_ids(c) == set()
    assert is_site_archived(c, 7) is False
    c.execute("CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT)")
    c.execute("INSERT INTO sites VALUES(7,'https://fixture.invalid')")
    assert archived_site_ids(c) == set()
    assert filter_active_sites(c, [{"id": 7}]) == [{"id": 7}]
    c.close()


def test_clone_queue_refuses_archived_source_or_target(conn):
    remove_connection(conn, 7)
    for source, target in ((7, 37), (37, 7)):
        with pytest.raises(SiteConnectionError) as caught:
            enqueue_clone_job(conn, source_site_id=source, target_site_id=target, product_ids=[10],
                              options={}, target_url="https://fixture.invalid", created_by_id="1", created_by_name="Fixture")
        assert caught.value.code == "SITE_ARCHIVED"
        conn.rollback()
    assert conn.execute("SELECT COUNT(*) FROM product_clone_jobs").fetchone()[0] == 0


@pytest.mark.parametrize("site_id,code,status", [(True, "INVALID_SITE_ID", 400),
                                                (0, "INVALID_SITE_ID", 400),
                                                (999, "SITE_NOT_FOUND", 404)])
def test_invalid_and_unknown_site_are_public_errors(conn, site_id, code, status):
    with pytest.raises(SiteConnectionError) as caught:
        remove_connection(conn, site_id)
    assert (caught.value.code, caught.value.status) == (code, status)
