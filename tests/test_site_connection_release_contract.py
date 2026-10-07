"""Connection removal contracts against AST-isolated application functions.

No app import, configured database, or real HTTP is allowed. The fixtures keep
historical source permissions, partner revenue, and durable sync records real.
"""
import ast
import importlib
import json
from pathlib import Path
import sqlite3
from types import SimpleNamespace

from flask import Flask, jsonify, redirect, render_template_string, request, url_for
import pytest
import requests

import db_backend as db
from partner_site_scope import init_partner_site_scope
from stock_sync_schema import migrate as init_stock_sync_tables


ROOT = Path(__file__).resolve().parents[1]
TARGET_URL = "https://historical.example.invalid"
OTHER_URL = "https://other.example.invalid"
HISTORY_TABLES = (
    "orders", "sync_runs", "sync_site_progress", "sync_page_dispatches",
    "sync_page_receipts", "sync_events", "sync_task_outbox", "external_operations",
    "user_site_permissions", "user_country_permissions", "partner_sites",
)


class Connection:
    """Own one in-memory DB while route-local close calls remain harmless."""
    def __init__(self, raw):
        self.raw = raw
        self.statements = []

    def execute(self, sql, params=()):
        self.statements.append(sql)
        if "delete from sites" in " ".join(sql.lower().split()):
            pytest.fail("Removing a connection must never physically delete a site")
        return self.raw.execute(sql, params)

    def executescript(self, sql):
        return self.raw.executescript(sql)

    def commit(self):
        self.raw.commit()

    def rollback(self):
        self.raw.rollback()

    def close(self):
        pass


def _compile_functions(scope, names):
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    # Only the new standalone connection service is allowed as a global import.
    # App startup imports and statements are deliberately never executed.
    for node in tree.body:
        if isinstance(node, ast.ImportFrom) and node.module and node.module.startswith("site_connection"):
            module = importlib.import_module(node.module)
            for alias in node.names:
                scope[alias.asname or alias.name] = getattr(module, alias.name)
    functions = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and (node.name in names or node.name.startswith("_site_connection")):
            node.decorator_list = []
            functions.append(node)
    assert names.issubset({node.name for node in functions})
    exec(compile(ast.Module(body=functions, type_ignores=[]), "app.py", "exec"), scope)
    return scope


@pytest.fixture
def contract(monkeypatch):
    def forbidden(*args, **kwargs):
        pytest.fail("Configured database and real HTTP are forbidden")

    monkeypatch.setattr(requests.sessions.Session, "request", forbidden)
    monkeypatch.setattr(db, "connect", forbidden)
    monkeypatch.setattr(db, "is_postgres_backend", lambda: False)
    raw = sqlite3.connect(":memory:")
    raw.row_factory = sqlite3.Row
    raw.execute("PRAGMA foreign_keys=ON")
    raw.executescript("""
        CREATE TABLE sites(
            id INTEGER PRIMARY KEY AUTOINCREMENT, url TEXT NOT NULL,
            consumer_key TEXT NOT NULL, consumer_secret TEXT NOT NULL,
            manager TEXT, country TEXT, mask_id TEXT, product_master_id INTEGER,
            cod_on_hold_is_shipped INTEGER, last_sync TEXT,
            api_status TEXT DEFAULT 'unknown', last_api_error TEXT,
            tracking_api_status TEXT, api_write_status TEXT
        );
        INSERT INTO sites VALUES(7,'https://historical.example.invalid','fixture-key','fixture-secret',
            'Fixture Owner','PL','history-mask',3,1,'2026-10-01','ok',NULL,'ok','ok');
        INSERT INTO sites VALUES(37,'https://other.example.invalid','other-key','other-secret',
            'Other Owner','CZ','other-mask',3,0,'2026-10-02','ok',NULL,'ok','ok');
        CREATE TABLE users(id INTEGER PRIMARY KEY,username TEXT,name TEXT,role TEXT,
            can_manage_products INTEGER,can_manage_own_products INTEGER,is_active INTEGER);
        INSERT INTO users VALUES(1,'admin','Fixture Owner','admin',1,0,1);
        INSERT INTO users VALUES(2,'explicit','History Reader','user',0,0,1);
        INSERT INTO users VALUES(3,'country','Country Reader','user',0,0,1);
        CREATE TABLE user_site_permissions(user_id INTEGER,site_id INTEGER);
        CREATE TABLE user_site_exclusions(user_id INTEGER,site_id INTEGER);
        CREATE TABLE user_country_permissions(user_id INTEGER,country TEXT);
        INSERT INTO user_site_permissions VALUES(2,7);
        INSERT INTO user_country_permissions VALUES(3,'PL');
        CREATE TABLE partners(id INTEGER PRIMARY KEY,name TEXT,currency TEXT);
        INSERT INTO partners VALUES(1,'Fixture Partner','PLN');
        CREATE TABLE partner_sites(partner_id INTEGER,site_id INTEGER,UNIQUE(partner_id,site_id));
        INSERT INTO partner_sites VALUES(1,7);
        CREATE TABLE orders(id TEXT PRIMARY KEY,source TEXT,currency TEXT,status TEXT,
            date_created TEXT,total REAL,shipping_total REAL,payment_method TEXT,
            is_undelivered INTEGER,is_problem_return INTEGER,shipping_loss_amount REAL,
            billing TEXT,shipping TEXT,line_items TEXT);
        INSERT INTO orders VALUES('7-101','https://historical.example.invalid','PLN','on-hold',
            '2026-10-01',100,10,'cod',0,0,0,'{}','{}','[]');
        INSERT INTO orders VALUES('7-102','https://historical.example.invalid','PLN','completed',
            '2026-10-02',200,20,'cod',0,0,0,'{}','{}','[]');
        INSERT INTO orders VALUES('37-101','https://other.example.invalid','CZK','completed',
            '2026-10-02',300,30,'cod',0,0,0,'{}','{}','[]');
        CREATE TABLE settings(key TEXT PRIMARY KEY,value TEXT);
        CREATE TABLE exchange_rates(id INTEGER,year_month TEXT,currency TEXT,rate_to_cny REAL,updated_at TEXT);
        CREATE TABLE product_masters(id INTEGER PRIMARY KEY,label TEXT,url TEXT,consumer_key TEXT,
            consumer_secret TEXT,api_status TEXT,last_api_error TEXT,last_tested_at TEXT,created_at TEXT,updated_at TEXT);
        INSERT INTO product_masters VALUES(3,'Other Master','https://master.example.invalid',
            'master-key','master-secret','ok',NULL,NULL,NULL,NULL);
        CREATE TABLE sync_runs(run_id TEXT PRIMARY KEY,mode TEXT,status TEXT,
            cancellation_requested INTEGER,total_sites INTEGER,completed_sites INTEGER,
            fetched_orders INTEGER,written_orders INTEGER,changed_orders INTEGER);
        INSERT INTO sync_runs VALUES('fixture-success','auto','success',0,2,2,3,3,3);
        INSERT INTO sync_runs VALUES('fixture-error','quick','error',0,1,1,0,0,0);
        CREATE TABLE sync_site_progress(run_id TEXT,site_id INTEGER,status TEXT,current_page INTEGER,
            fetched_count INTEGER,written_count INTEGER,changed_count INTEGER,retry_count INTEGER,
            heartbeat_at TEXT,started_at TEXT,finished_at TEXT,error_message TEXT,
            PRIMARY KEY(run_id,site_id),FOREIGN KEY(site_id) REFERENCES sites(id) ON DELETE RESTRICT);
        INSERT INTO sync_site_progress VALUES('fixture-success',7,'success',1,2,2,2,0,NULL,NULL,NULL,NULL);
        INSERT INTO sync_site_progress VALUES('fixture-success',37,'success',1,1,1,1,0,NULL,NULL,NULL,NULL);
        INSERT INTO sync_site_progress VALUES('fixture-error',7,'error',1,0,0,0,1,NULL,NULL,NULL,'source unavailable');
        CREATE TABLE sync_page_dispatches(run_id TEXT,site_id INTEGER,page INTEGER,status TEXT,
            PRIMARY KEY(run_id,site_id,page),FOREIGN KEY(run_id,site_id)
                REFERENCES sync_site_progress(run_id,site_id) ON DELETE CASCADE);
        INSERT INTO sync_page_dispatches VALUES('fixture-success',7,1,'completed');
        INSERT INTO sync_page_dispatches VALUES('fixture-success',37,1,'completed');
        INSERT INTO sync_page_dispatches VALUES('fixture-error',7,1,'error');
        CREATE TABLE sync_page_receipts(receipt_id INTEGER PRIMARY KEY,run_id TEXT,site_id INTEGER,
            page INTEGER,content_hash TEXT,post_commit_status TEXT,
            FOREIGN KEY(run_id,site_id) REFERENCES sync_site_progress(run_id,site_id) ON DELETE CASCADE);
        INSERT INTO sync_page_receipts VALUES(1,'fixture-success',7,1,'target-hash','completed');
        INSERT INTO sync_page_receipts VALUES(2,'fixture-success',37,1,'other-hash','completed');
        CREATE TABLE sync_events(event_id INTEGER PRIMARY KEY,run_id TEXT,site_id INTEGER,
            level TEXT,event_type TEXT,message TEXT,details TEXT,created_at TEXT);
        INSERT INTO sync_events VALUES(1,'fixture-success',7,'info','site_completed','done','{}',NULL);
        CREATE TABLE sync_task_outbox(outbox_id INTEGER PRIMARY KEY,dedupe_key TEXT,payload TEXT,
            status TEXT,task_name TEXT,queue_name TEXT,updated_at TEXT,last_error TEXT);
        INSERT INTO sync_task_outbox VALUES(1,'fetch:fixture-success:7:1',
            '{"run_id":"fixture-success","site_id":7,"page":1}','published','woo_sync.fetch_page','sync_fetch',NULL,NULL);
        CREATE TABLE external_operations(operation_id TEXT PRIMARY KEY,site_id INTEGER,status TEXT,
            FOREIGN KEY(site_id) REFERENCES sites(id) ON DELETE RESTRICT);
        CREATE TABLE inv_push_runs(id INTEGER PRIMARY KEY,site_id INTEGER,status TEXT,
            FOREIGN KEY(site_id) REFERENCES sites(id) ON DELETE RESTRICT);
        CREATE TABLE inv_push_locks(site_id INTEGER PRIMARY KEY,lock_token TEXT,acquired_at TEXT,
            FOREIGN KEY(site_id) REFERENCES sites(id) ON DELETE CASCADE);
        CREATE TABLE inv_site_sync_config(site_id INTEGER PRIMARY KEY,mode TEXT,last_error TEXT,
            paused_reason TEXT,next_run_at TEXT,updated_at TEXT,updated_by INTEGER,updated_by_name TEXT,
            FOREIGN KEY(site_id) REFERENCES sites(id) ON DELETE CASCADE);
        CREATE TABLE inv_site_sync_audit(id INTEGER PRIMARY KEY,site_id INTEGER,action TEXT,
            before_json TEXT,after_json TEXT,operator_id INTEGER,operator_name TEXT,created_at TEXT);
        CREATE TABLE product_clone_jobs(id TEXT PRIMARY KEY,source_site_id INTEGER,target_site_id INTEGER,status TEXT);
        CREATE TABLE oms_integration_jobs(id INTEGER PRIMARY KEY,aggregate_type TEXT,aggregate_id TEXT,
            payload_json TEXT,status TEXT,job_type TEXT);
        CREATE TABLE oms_fulfillments(id TEXT PRIMARY KEY,order_id TEXT,status TEXT);
        CREATE TABLE oms_order_fulfillment_state(order_id TEXT PRIMARY KEY,aggregate_status TEXT);
        CREATE TABLE inv_site_sku_map(id INTEGER PRIMARY KEY,site_id INTEGER,sku_id INTEGER,is_active INTEGER);
    """)
    init_partner_site_scope(raw)
    init_stock_sync_tables(raw)
    raw.commit()
    connection = Connection(raw)
    app = Flask("isolated-site-connection-contract")
    app.config.update(TESTING=True, SECRET_KEY="fixture-only")
    current_user = SimpleNamespace(id=1,username="admin",name="Fixture Owner",role="admin",
        is_admin=lambda: True, is_authenticated=True)
    app.context_processor(lambda: {"current_user": current_user})
    contexts = []

    def render_template(name, **kwargs):
        contexts.append({"name": name, **kwargs})
        return "offline-rendered"

    names = {"add_site", "delete_site", "settings", "admin_required", "_resolve_user_sites",
             "_calc_partner_net_sales", "_on_hold_is_shipped_clause", "_revenue_status_cond"}
    scope = _compile_functions({"app": app, "current_user": current_user,
        "get_db_connection": lambda: connection, "jsonify": jsonify, "request": request,
        "json": json, "sqlite3": SimpleNamespace(is_postgres_backend=lambda: False),
        "redirect": redirect, "url_for": url_for, "render_template_string": render_template_string,
        "render_template": render_template, "_can_manage_all_settings": lambda user: True}, names)
    app.add_url_rule("/api/sites/<int:site_id>", view_func=scope["admin_required"](scope["delete_site"]), methods=["DELETE"])
    app.add_url_rule("/api/sites", view_func=scope["admin_required"](scope["add_site"]), methods=["POST"])
    app.add_url_rule("/settings", view_func=scope["settings"])
    yield SimpleNamespace(raw=raw,connection=connection,app=app,client=app.test_client(),
                          scope=scope,contexts=contexts,current_user=current_user)
    raw.close()


def snapshot(contract, tables=HISTORY_TABLES):
    return {table: [dict(row) for row in contract.raw.execute(f"SELECT * FROM {table} ORDER BY 1,2")]
            for table in tables}


def remove(contract):
    response = contract.client.delete("/api/sites/7")
    assert response.status_code == 200, response.get_json()
    assert response.get_json()["success"] is True
    return response.get_json()


def test_connection_removal_preserves_site_identity_and_all_sync_history(contract):
    before = snapshot(contract)
    site = dict(contract.raw.execute("SELECT * FROM sites WHERE id=7").fetchone())
    other = dict(contract.raw.execute("SELECT * FROM sites WHERE id=37").fetchone())
    remove(contract)
    after = dict(contract.raw.execute("SELECT * FROM sites WHERE id=7").fetchone())
    for field in ("id", "url", "manager", "country", "mask_id", "cod_on_hold_is_shipped", "last_sync"):
        assert after[field] == site[field]
    assert after["consumer_key"] == after["consumer_secret"] == ""
    assert after["api_status"] == "archived"
    assert after["product_master_id"] is None
    assert dict(contract.raw.execute("SELECT * FROM sites WHERE id=37").fetchone()) == other
    assert snapshot(contract) == before


def test_historical_grants_country_and_partner_two_order_revenue_keep_original_id(contract):
    explicit_before = contract.scope["_resolve_user_sites"](2)
    country_before = contract.scope["_resolve_user_sites"](3)
    partner_before = contract.scope["_calc_partner_net_sales"](1, 2026, 10)
    assert explicit_before == country_before == [TARGET_URL]
    assert partner_before["total_orders"] == 2
    assert partner_before["total_gross_pln"] == 300
    assert partner_before["total_net_pln"] == 270
    remove(contract)
    assert contract.scope["_resolve_user_sites"](2) == explicit_before
    assert contract.scope["_resolve_user_sites"](3) == country_before
    assert contract.scope["_calc_partner_net_sales"](1, 2026, 10) == partner_before
    assert [row["id"] for row in contract.raw.execute(
        "SELECT id FROM orders WHERE source=? ORDER BY id", (TARGET_URL,))] == ["7-101", "7-102"]


@pytest.mark.parametrize("own_settings", [False, True])
def test_removed_connection_is_absent_from_active_settings_without_hiding_history(contract, own_settings):
    contract.scope["_can_manage_all_settings"] = lambda user: not own_settings
    remove(contract)
    response = contract.client.get("/settings")
    assert response.status_code == 200
    assert 7 not in {row["id"] for row in contract.contexts[-1]["sites"]}
    assert contract.scope["_resolve_user_sites"](2) == [TARGET_URL]


def test_restore_same_url_reuses_id_and_preserves_historical_permissions(contract):
    before = snapshot(contract)
    remove(contract)
    response = contract.client.post("/api/sites", json={"url": TARGET_URL + "/",
        "consumer_key": "replacement-key", "consumer_secret": "replacement-secret",
        "manager": "Fixture Owner", "country": "PL", "mask_id": "history-mask"})
    assert response.status_code == 200, response.get_json()
    assert response.get_json()["success"] is True
    site = dict(contract.raw.execute("SELECT * FROM sites WHERE url=?", (TARGET_URL,)).fetchone())
    assert site["id"] == 7
    assert site["consumer_key"] == "replacement-key"
    assert site["consumer_secret"] == "replacement-secret"
    assert site["api_status"] != "archived"
    assert contract.raw.execute("SELECT COUNT(*) FROM sites").fetchone()[0] == 2
    assert snapshot(contract) == before
    assert contract.scope["_resolve_user_sites"](2) == [TARGET_URL]


def test_missing_connection_returns_404_without_business_mutation(contract):
    before = snapshot(contract)
    response = contract.client.delete("/api/sites/999")
    assert response.status_code == 404
    assert snapshot(contract) == before


@pytest.mark.parametrize("profile", [{}, {"manager": "", "country": "", "mask_id": ""}])
def test_restore_with_only_credentials_preserves_history_classification(contract, profile):
    original = dict(contract.raw.execute("SELECT * FROM sites WHERE id=7").fetchone())
    remove(contract)
    response = contract.client.post("/api/sites", json={"url": TARGET_URL,
        "consumer_key": "replacement-key", "consumer_secret": "replacement-secret", **profile})
    assert response.status_code == 200, response.get_json()
    after = dict(contract.raw.execute("SELECT * FROM sites WHERE id=7").fetchone())
    for field in ("id", "url", "manager", "country", "mask_id", "cod_on_hold_is_shipped", "last_sync"):
        assert after[field] == original[field]
    assert contract.scope["_calc_partner_net_sales"](1, 2026, 10)["total_orders"] == 2


def test_non_admin_cannot_remove_connection_or_change_historical_data(contract):
    before = snapshot(contract, HISTORY_TABLES + ("sites",))
    contract.current_user.is_admin = lambda: False
    contract.current_user.username = "ordinary-reader"
    response = contract.client.delete("/api/sites/7")
    assert response.status_code == 403
    assert snapshot(contract, HISTORY_TABLES + ("sites",)) == before


def test_repeated_removal_is_idempotent_and_keeps_all_history(contract):
    before = snapshot(contract)
    remove(contract)
    archived = dict(contract.raw.execute("SELECT * FROM sites WHERE id=7").fetchone())
    result = remove(contract)
    assert result["already_archived"] is True
    assert dict(contract.raw.execute("SELECT * FROM sites WHERE id=7").fetchone()) == archived
    assert snapshot(contract) == before


@pytest.mark.parametrize("post_commit_status", ["pending", "processing", "error"])
def test_terminal_run_with_unfinished_post_commit_returns_409_without_changes(contract, post_commit_status):
    contract.raw.execute("UPDATE sync_page_receipts SET post_commit_status=? WHERE site_id=7",
                         (post_commit_status,))
    contract.raw.commit()
    before = snapshot(contract, HISTORY_TABLES + ("sites",))
    response = contract.client.delete("/api/sites/7")
    assert response.status_code == 409
    assert response.get_json()["success"] is False
    assert snapshot(contract, HISTORY_TABLES + ("sites",)) == before


def test_database_failure_rolls_back_and_does_not_disclose_internal_error(contract):
    contract.raw.execute("""CREATE TRIGGER fail_connection_archive AFTER UPDATE ON sites
        WHEN NEW.api_status='archived' BEGIN SELECT RAISE(ABORT,'fixture-internal-secret'); END""")
    contract.raw.commit()
    before = snapshot(contract, HISTORY_TABLES + ("sites", "inv_site_sync_audit"))
    response = contract.client.delete("/api/sites/7")
    assert response.status_code == 500
    assert response.get_json()["success"] is False
    assert "fixture-internal-secret" not in response.get_data(as_text=True)
    assert snapshot(contract, HISTORY_TABLES + ("sites", "inv_site_sync_audit")) == before
