"""Independent catalog scope contracts using real routes and synthetic data.

No app startup/configured database or network. Selection scope is a narrowing
constraint, never an authority grant; archived/other-owner data stays unreadable.
"""
from functools import wraps
from pathlib import Path
from types import SimpleNamespace
import shutil
import sqlite3
import subprocess

from flask import Flask, jsonify
from flask_login import LoginManager, UserMixin, current_user
import pytest
import requests

import product_manager_catalog as catalog


class ScopeUser(UserMixin):
    id = "scope-contract-user"
    name = "Alice"
    own_scoped = False
    may_manage = True

    def product_manager_own_scoped(self):
        return self.own_scoped


@pytest.fixture
def scope_system(tmp_path, monkeypatch):
    path = tmp_path / "synthetic-scope.db"

    def connect():
        conn = sqlite3.connect(path)
        conn.row_factory = sqlite3.Row
        return conn

    with connect() as conn:
        conn.execute("""CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT,
            manager TEXT,country TEXT,consumer_key TEXT,consumer_secret TEXT,
            product_master_id INTEGER,api_status TEXT)""")
        conn.executemany("INSERT INTO sites VALUES(?,?,?,?,?,?,?,?)", [
            (1, "https://one.invalid", " Alice ", " PL ", "fake-key-1", "fake-secret-1", 99, "ok"),
            (2, "https://two.invalid", "Alice", "CZ", "fake-key-2", "fake-secret-2", None, "ok"),
            (3, "https://three.invalid", "Bob", "PL", "fake-key-3", "fake-secret-3", None, "ok"),
            (4, "https://four.invalid", "", None, "fake-key-4", "fake-secret-4", None, "ok"),
            (5, "https://five.invalid", "Alice", "PL", None, None, 99, "ok"),
            (6, "https://archived.invalid", "Alice", "PL", "fake-key-6", "fake-secret-6", None, "archived"),
            (7, "https://sentinel.invalid", "all", "__unassigned__", "fake-key-7", "fake-secret-7", None, "ok"),
        ])

    def forbidden(*args, **kwargs):
        pytest.fail("Network outside the synthetic Woo reader is forbidden")

    monkeypatch.setattr(requests.sessions.Session, "request", forbidden)
    user = ScopeUser()
    app = Flask("catalog-scope-contract")
    app.config.update(TESTING=True, SECRET_KEY="synthetic-scope-only")
    LoginManager(app).user_loader(lambda _identity: user)

    def manager_required(function):
        @wraps(function)
        def wrapped(*args, **kwargs):
            if not current_user.may_manage:
                return jsonify({"error": "无商品管理权限"}), 403
            return function(*args, **kwargs)
        return wrapped

    app.register_blueprint(catalog.create_catalog_blueprint(connect, manager_required))
    client = app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = user.id
        session["_fresh"] = True
    calls = []

    def woo(site, resource, params=None):
        calls.append({"site_id": site["id"], "resource": resource, "params": params})
        assert resource == "products"
        return [{
            "id": 10, "name": "Synthetic IGET 12000 Puffs", "sku": "same-id-across-sites",
            "type": "simple", "status": "publish", "attributes": [], "brands": [], "variations": [],
            "manage_stock": True, "stock_quantity": 5, "stock_status": "instock",
            "price": "10.00", "regular_price": "10.00", "sale_price": "",
            "country": "FORGED-CLIENT-COUNTRY", "manager": "FORGED-CLIENT-MANAGER",
        }], {"X-WP-Total": "1", "X-WP-TotalPages": "1"}

    monkeypatch.setattr(catalog, "_read_wc", woo)
    return SimpleNamespace(client=client, user=user, calls=calls, connect=connect)


def metadata(system, **params):
    return system.client.get("/api/product-manager/catalog-sites", query_string=params)


def read_page(system, site_id=1, **params):
    return system.client.get("/api/product-manager/catalog-page", query_string={"site_id": site_id, **params})


def test_initial_metadata_never_scans_and_never_exposes_credentials(scope_system):
    result = metadata(scope_system)
    assert result.status_code == 200
    sites = result.get_json()["sites"]
    assert {site["id"] for site in sites} == {1, 2, 3, 4, 5, 7}
    assert scope_system.calls == []
    assert next(site for site in sites if site["id"] == 5)["configured"] is False
    assert next(site for site in sites if site["id"] == 1)["manager"] == "Alice"
    assert next(site for site in sites if site["id"] == 1)["country"] == "PL"
    assert "fake-key" not in result.get_data(as_text=True)
    assert "fake-secret" not in result.get_data(as_text=True)


@pytest.mark.parametrize("params,expected", [
    ({"manager": "Alice"}, {1, 2, 5}),
    ({"country": "PL"}, {1, 3, 5}),
    ({"manager": "Alice", "country": "PL"}, {1, 5}),
    ({"manager": "Alice", "country": "PL", "site_ids": "1,2,3"}, {1}),
    ({"manager": "Alice", "country": "PL", "site_ids": "1,1,5"}, {1, 5}),
    ({"manager": "Alice", "country": "PL", "site_ids": "2,3"}, set()),
    ({"site_ids": ""}, set()), ({"site_ids": "999999"}, set()),
    ({"site_id": "2", "site_ids": "1,2"}, {2}),
    ({"site_id": "3", "site_ids": "1,2"}, set()),
    ({"manager_mode": "unassigned"}, {4}),
    ({"country_mode": "unassigned"}, {4}),
    ({"manager_mode": "exact", "manager": "all", "country_mode": "exact", "country": "__unassigned__"}, {7}),
])
def test_metadata_scope_is_intersection_with_no_empty_fallback(scope_system, params, expected):
    result = metadata(scope_system, **params)
    assert result.status_code == 200
    assert {site["id"] for site in result.get_json()["sites"]} == expected
    assert scope_system.calls == []


@pytest.mark.parametrize("params", [
    {"site_ids": "0"}, {"site_ids": "-1"}, {"site_ids": "1,bad"},
    {"site_ids": "1.0"}, {"site_ids": "1,,2"},
    {"manager_mode": "exact"}, {"country_mode": "exact", "country": " "},
    {"manager_mode": "all", "manager": "Alice"},
    {"country_mode": "all", "country": "PL"},
    {"manager_mode": "unassigned", "manager": "Alice"},
    {"country_mode": "unassigned", "country": "PL"},
    {"manager_mode": "unsupported"},
])
def test_malformed_scope_is_rejected_before_any_product_read(scope_system, params):
    result = metadata(scope_system, **params)
    assert result.status_code == 400
    assert "error" in result.get_json()
    result = read_page(scope_system, **params)
    assert result.status_code == 400
    assert scope_system.calls == []


@pytest.mark.parametrize("params", [
    {"site_ids": ""}, {"site_ids": "2,3"}, {"manager": "Bob"},
    {"country": "CZ"}, {"manager_mode": "unassigned"},
    {"country_mode": "unassigned"},
])
def test_product_endpoint_rejects_outside_scope_before_woo(scope_system, params):
    result = read_page(scope_system, **params)
    assert result.status_code == 403
    assert result.get_json()["code"] == "scope_mismatch"
    assert scope_system.calls == []


def test_only_intersection_site_is_read_and_country_is_fresh_server_metadata(scope_system):
    params = {"manager": "Alice", "country": "PL", "site_ids": "1,2,3"}
    scope = metadata(scope_system, **params).get_json()["sites"]
    for site in scope:
        result = read_page(scope_system, site["id"], **params)
        assert result.status_code == 200 and result.get_json()["complete_page"] is True
        row = result.get_json()["rows"][0]
        assert row["country"] == "PL" and row["manager"] == "Alice"
    assert [call["site_id"] for call in scope_system.calls] == [1]


def test_stale_scope_cannot_bypass_owner_revocation_or_archive(scope_system):
    scope_system.user.own_scoped = True
    assert {site["id"] for site in metadata(scope_system).get_json()["sites"]} == {1, 2, 5}
    result = read_page(scope_system, site_id=3, manager="Bob", country="PL", site_ids="3")
    assert result.status_code == 403 and result.get_json()["code"] == "site_forbidden"
    with scope_system.connect() as conn:
        conn.execute("UPDATE sites SET manager='Bob' WHERE id=1")
    result = read_page(scope_system, manager="Alice", country="PL", site_ids="1")
    assert result.status_code == 403
    result = read_page(scope_system, site_id=6, manager="Alice", country="PL", site_ids="6")
    assert result.status_code == 409 and result.get_json()["code"] == "site_archived"
    assert scope_system.calls == []


def test_no_owner_and_missing_permissions_do_not_fallback_to_all(scope_system):
    scope_system.user.own_scoped = True
    scope_system.user.name = ""
    assert metadata(scope_system).get_json()["sites"] == []
    assert read_page(scope_system).status_code == 403
    scope_system.user.may_manage = False
    assert metadata(scope_system).status_code == 403
    assert read_page(scope_system).status_code == 403
    scope_system.client.get("/unrelated")
    with scope_system.client.session_transaction() as session:
        session.clear()
    assert metadata(scope_system).status_code == 403
    assert read_page(scope_system).status_code == 403
    assert scope_system.calls == []


@pytest.mark.skipif(shutil.which("node") is None, reason="Node.js is required")
def test_real_catalog_scope_controller_editor_and_mount_contracts():
    root = Path(__file__).resolve().parents[1]
    completed = subprocess.run(
        [shutil.which("node"), str(root / "tests/catalog_load_scope_harness.cjs")],
        cwd=root, capture_output=True, text=True, encoding="utf-8", timeout=30,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr
    assert "Independent catalog scope contracts" in completed.stdout
