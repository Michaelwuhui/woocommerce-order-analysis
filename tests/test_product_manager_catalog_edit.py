"""Direct-child editing safety contracts, isolated SQL and fake Woo requests."""
from contextlib import contextmanager
from copy import deepcopy
from functools import wraps
import ast
import json
from pathlib import Path
import sqlite3

from flask import Flask, jsonify
from flask_login import LoginManager, UserMixin, current_user
import pytest
import requests

import product_manager_catalog_edit as edit
import product_clone_worker as worker
from product_clone_jobs import claim_clone_job, enqueue_clone_job, get_clone_job, init_product_clone_jobs
from product_clone_sku import build_clone_sku, make_clone_suffix, normalize_clone_suffix
from product_manager_service import parse_wc_response


class User(UserMixin):
    id = "2"
    username = "operator"
    name = "Michael"
    allowed = True
    scoped = True
    def product_manager_own_scoped(self):
        return self.scoped


class Response:
    def __init__(self, value, status=200):
        self.value, self.status_code = deepcopy(value), status
        self.text = json.dumps(value)
    def json(self):
        return deepcopy(self.value)


class Remote:
    RequestException = requests.RequestException
    def __init__(self):
        self.items = {
            "https://child.test/wp-json/wc/v3/products/10": {
                "id": 10, "type": "simple", "status": "publish", "sku": "CHILD-10", "name": "Device",
                "attributes": [], "brands": [{"name": "Example"}], "manage_stock": True,
                "stock_quantity": 8, "stock_status": "instock", "regular_price": "12.00", "sale_price": "", "price": "12.00",
            },
        }
        self.calls = []
        self.in_lease = False
        self.preflight_hook = None
        self.ignore_meta = False
        self.force_parent = False
        self.wrong_read_id = False
        self.collection_override = None
        self.timeout_put = None
        self.bridge_timeout = None
        self.auto_stock_status = False
        self.no_stock_threshold = 0
        self.allow_backorders = False
        self.restore_from_bridge = False
        self.write_leases = []
    def get(self, url, **kwargs):
        self.calls.append(("GET", url, kwargs))
        if self.in_lease and self.preflight_hook:
            hook, self.preflight_hook = self.preflight_hook, None
            hook()
        if url.endswith("/products"):
            if self.collection_override is not None:
                return Response(self.collection_override)
            ids = set(int(identity) for identity in kwargs["params"]["include"].split(","))
            return Response([item for resource, item in self.items.items() if resource.startswith(url + "/") and item["id"] in ids])
        item = deepcopy(self.items[url])
        if self.in_lease and self.wrong_read_id:
            item["id"] = 999
        return Response(item)
    def put(self, url, **kwargs):
        self.calls.append(("PUT", url, kwargs))
        self.write_leases.append(self.in_lease)
        bridge_only = set(kwargs["json"]) == {"meta_data"}
        if bridge_only and self.bridge_timeout == "before":
            raise requests.Timeout("fake bridge write not confirmed")
        if self.timeout_put == "before":
            raise requests.Timeout("fake ambiguous timeout")
        payload = deepcopy(kwargs["json"])
        if self.ignore_meta:
            payload.pop("meta_data", None)
        self.items[url].update(payload)
        current = self.items[url]
        if self.auto_stock_status and current.get("manage_stock") is True:
            quantity = current.get("stock_quantity")
            if type(quantity) is int:
                current["stock_status"] = ("instock" if quantity > self.no_stock_threshold
                    else "onbackorder" if self.allow_backorders else "outofstock")
        if self.restore_from_bridge:
            bridge = {row["key"]: row["value"] for row in current.get("meta_data", [])}
            current.update(manage_stock=bridge["wcms_stock_manage"] == "yes",
                           stock_quantity=int(bridge["wcms_stock_qty"]), stock_status=bridge["wcms_stock_status"])
        if self.force_parent:
            self.items[url]["manage_stock"] = "parent"
        if self.timeout_put == "after":
            raise requests.Timeout("fake response lost after save")
        if bridge_only and self.bridge_timeout == "after":
            raise requests.Timeout("fake bridge saved but response lost")
        return Response(self.items[url])


@pytest.fixture
def system(tmp_path, monkeypatch):
    path = tmp_path / "catalog-edit.db"
    def connect():
        connection = sqlite3.connect(path)
        connection.row_factory = sqlite3.Row
        return connection
    connection = connect()
    connection.execute("CREATE TABLE sites(id INTEGER,url TEXT,manager TEXT,country TEXT,consumer_key TEXT,consumer_secret TEXT,product_master_id INTEGER)")
    connection.executemany("INSERT INTO sites VALUES(?,?,?,?,?,?,?)", [
        (1, "https://child.test", "Michael", "PL", "child_ck", "child_cs", 99),
        (2, "https://target.test", "Michael", "CZ", "target_ck", "target_cs", 99),
        (3, "https://other.test", "Jane", "PL", "other_ck", "other_cs", None),
        (4, "https://missing.test", "Michael", "PL", "", "", 99),
    ])
    connection.execute("CREATE TABLE users(id INTEGER,username TEXT,name TEXT,role TEXT,can_manage_products INTEGER,can_manage_own_products INTEGER)")
    connection.execute("INSERT INTO users VALUES(2,'operator','Michael','user',1,0)")
    connection.commit()
    connection.close()
    app = Flask(__name__)
    app.config.update(SECRET_KEY="offline-catalog-edit", TESTING=True)
    user = User()
    LoginManager(app).user_loader(lambda _id: user)
    def permission(function):
        @wraps(function)
        def decorated(*args, **kwargs):
            if not current_user.allowed:
                return jsonify(error="forbidden"), 403
            return function(*args, **kwargs)
        return decorated
    remote = Remote()
    audits, jobs = [], []
    def enqueue(connection, **kwargs):
        jobs.append(kwargs)
        return {"id": "job-1", "status": "queued", "total_count": len(kwargs["product_ids"])}
    app.register_blueprint(edit.create_catalog_edit_blueprint(connect, permission, lambda **kwargs: audits.append(kwargs),
        enqueue_clone_job=enqueue, make_clone_suffix=lambda: "stable-job-suffix", requests_client=remote))
    client = app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = "2"
    csrf = client.get("/api/product-manager/catalog-edit-config").get_json()["csrf_token"]
    @contextmanager
    def lease(_url, _payload):
        remote.in_lease = True
        try:
            yield
        finally:
            remote.in_lease = False
    monkeypatch.setattr("stock_sync_guard.legacy_write", lease)
    return {"client": client, "remote": remote, "csrf": csrf, "audits": audits, "jobs": jobs, "user": user, "connect": connect}


def preview(system, product_id=10, variation_id=0):
    return system["client"].get("/api/product-manager/catalog-item", query_string={"site_id": 1, "product_id": product_id, "variation_id": variation_id})


def update(system, changes, *, product_id=10, variation_id=0, snapshot=None, **overrides):
    snapshot = snapshot or preview(system, product_id, variation_id).get_json()
    body = {"site_id": 1, "product_id": product_id, "variation_id": variation_id,
            "expected_identity": snapshot["identity"], "expected_before": snapshot["before"], "changes": changes, "batch_id": "batch-1", **overrides}
    return system["client"].put("/api/product-manager/catalog-item", json=body, headers={"X-PM-CSRF": system["csrf"]})


def test_config_is_scoped_and_excludes_credentials(system):
    response = system["client"].get("/api/product-manager/catalog-edit-config")
    data = response.get_json()
    assert {site["id"] for site in data["sites"]} == {1, 2, 4}
    assert {site["id"] for site in data["clone_targets"]} == {1, 2}
    assert next(site for site in data["sites"] if site["id"] == 4)["can_edit"] is False
    assert "child_ck" not in response.get_data(as_text=True)


def test_direct_child_price_update_never_reads_or_writes_master(system):
    data = update(system, {"regular_price": "15"}).get_json()
    assert data["success"] is True and data["verification"]["status"] == "verified"
    assert data["item"]["regular_price"] == "15"
    assert all(url.startswith("https://child.test/") for _, url, _ in system["remote"].calls)
    assert data["routing"]["is_routed_to_master"] is False
    assert system["audits"][-1]["site"]["product_master_id"] is None
    assert system["audits"][-1]["trace"]["routing_mode"] == "direct"


@pytest.mark.parametrize("changes", [
    {"manage_stock": "false"}, {"stock_quantity": True}, {"stock_quantity": 1.2}, {"stock_quantity": -1},
    {"stock_quantity": "7"}, {"regular_price": "NaN"}, {"sale_price": "Infinity"}, {"regular_price": -1},
    {"regular_price": None}, {"name": "Overwrite"}, {}, {"stock_status": "unknown"},
    {"stock_quantity": 2147483648}, {"regular_price": "1e100000000"}, {"regular_price": "1e-100000000"},
    {"regular_price": "99999999999999"}, {"sale_price": "0.123456789"}, {"regular_price": "1" * 65}, {"stock_status": []},
])
def test_invalid_payload_rejected_without_any_put(system, changes):
    response = update(system, changes)
    assert response.status_code == 400
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


def test_csrf_and_owner_permission_block_without_remote_access(system):
    response = system["client"].put("/api/product-manager/catalog-item", json={})
    assert response.status_code == 403 and system["remote"].calls == []
    response = system["client"].get("/api/product-manager/catalog-item?site_id=3&product_id=10")
    assert response.status_code == 403 and system["remote"].calls == []


@pytest.mark.parametrize("change_identity", ["sku", "attributes"])
def test_identity_changed_after_preview_blocks_put(system, change_identity):
    before = preview(system).get_json()
    target = system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]
    target[change_identity] = "CHANGED" if change_identity == "sku" else [{"name": "Flavor", "options": ["New"]}]
    response = update(system, {"regular_price": "15"}, snapshot=before)
    assert response.status_code == 409 and response.get_json()["code"] == "STALE_IDENTITY"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


def test_state_changed_within_guard_after_api_get_blocks_put(system):
    before = preview(system).get_json()
    target = system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]
    system["remote"].preflight_hook = lambda: target.update(stock_quantity=9)
    response = update(system, {"stock_quantity": 10}, snapshot=before)
    assert response.status_code == 409 and response.get_json()["code"] == "STALE_STATE"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)
    assert system["audits"][-1]["status"] == "failed"


def test_unrelated_concurrent_field_does_not_block_requested_price_change(system):
    before = preview(system).get_json()
    system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]["stock_quantity"] = 9
    assert update(system, {"regular_price": "15"}, snapshot=before).get_json()["success"] is True


def test_wrong_id_on_guard_preflight_is_not_accepted(system):
    system["remote"].wrong_read_id = True
    response = update(system, {"regular_price": "15"})
    assert response.status_code == 409 and response.get_json()["code"] == "RESOURCE_MISMATCH"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


def test_wcms_metadata_is_merged_and_must_be_read_back(system):
    target = system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]
    target["meta_data"] = [{"key": "wcms_stock_manage", "value": "yes"}, {"key": "wcms_stock_qty", "value": 8}, {"key": "wcms_stock_status", "value": "instock"}]
    assert update(system, {"manage_stock": False, "stock_status": "outofstock"}).get_json()["success"] is True
    meta = {row["key"]: row["value"] for row in target["meta_data"]}
    assert meta == {"wcms_stock_manage": "no", "wcms_stock_qty": 8, "wcms_stock_status": "outofstock"}
    target.update(manage_stock=True, stock_status="instock")
    target["meta_data"] = [{"key": "wcms_stock_manage", "value": "yes"}]
    system["remote"].ignore_meta = True
    response = update(system, {"manage_stock": False, "stock_status": "outofstock"})
    assert response.status_code == 409 and response.get_json()["success"] is False
    assert response.get_json()["expected_stock_bridge"] == {"wcms_stock_manage": "no", "wcms_stock_qty": 8, "wcms_stock_status": "outofstock"}
    readback = preview(system).get_json()
    assert readback["stock_bridge"]["present"] is True
    assert readback["stock_bridge"]["values"] == {"wcms_stock_manage": "yes"}


def test_inherited_parent_management_does_not_count_as_independent_true(system):
    system["remote"].force_parent = True
    response = update(system, {"manage_stock": True, "stock_quantity": 7})
    assert response.status_code == 409 and response.get_json()["code"] == "WRITE_NOT_VERIFIED"


def test_clear_sale_price_is_explicit_and_keeps_regular_price(system):
    target = system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]
    target["sale_price"] = "10"
    assert update(system, {"sale_price": ""}).get_json()["success"] is True
    assert target["sale_price"] == "" and target["regular_price"] == "12.00"


def test_clone_enqueues_direct_job_without_any_remote_mutation(system):
    response = system["client"].post("/api/product-manager/catalog-clone", json={
        "source_site_id": 1, "target_site_id": 2, "product_ids": [10], "collision_mode": "clone_as_new", "status_on_target": "publish",
    }, headers={"X-PM-CSRF": system["csrf"]})
    assert response.status_code == 202
    job = system["jobs"][0]
    assert job["options"]["_catalog_direct_sites"] is True
    assert job["options"]["status_on_target"] == "draft"
    assert job["target_url"] == "https://target.test"
    assert set(job["options"]["_catalog_source_identities"]) == {"10"}
    assert len(system["remote"].calls) == 1
    assert system["remote"].calls[0][2]["params"]["include"] == "10"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


def test_direct_clone_worker_resolver_rechecks_live_owner_and_ignores_master(system):
    connection = system["connect"]()
    site, url, _, _ = worker.resolve_catalog_clone_site(connection, 1, 2)
    assert url == "https://child.test" and site["product_master_id"] == 99
    connection.execute("UPDATE sites SET manager='Jane' WHERE id=1")
    connection.commit()
    with pytest.raises(ValueError, match="负责人"):
        worker.resolve_catalog_clone_site(connection, 1, 2)
    connection.close()


def test_bridge_metadata_uses_fresh_guard_state_for_unmodified_stock_fields(system):
    target = system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]
    target["meta_data"] = [{"key": "wcms_stock_manage", "value": "yes"}, {"key": "wcms_stock_qty", "value": 8}]
    system["remote"].preflight_hook = lambda: target.update(stock_quantity=9)
    response = update(system, {"stock_status": "outofstock"})
    assert response.get_json()["success"] is True
    assert {row["key"]: row["value"] for row in target["meta_data"]}["wcms_stock_qty"] == 9


def test_live_permission_revocation_after_read_stops_put(system):
    def revoke():
        connection = system["connect"]()
        connection.execute("UPDATE users SET can_manage_products=0 WHERE id=2")
        connection.commit()
        connection.close()
    system["remote"].preflight_hook = revoke
    response = update(system, {"regular_price": "15"})
    assert response.status_code == 403 and response.get_json()["code"] == "PERMISSION_REVOKED"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


def add_variation(system):
    remote = system["remote"]
    remote.items["https://child.test/wp-json/wc/v3/products/20"] = {
        **deepcopy(remote.items["https://child.test/wp-json/wc/v3/products/10"]),
        "id": 20, "type": "variable", "sku": "PARENT-20", "name": "Device All Flavors",
        "attributes": [{"id": 0, "name": "Flavor", "options": ["A", "B"]}],
    }
    remote.items["https://child.test/wp-json/wc/v3/products/20/variations/21"] = {
        **deepcopy(remote.items["https://child.test/wp-json/wc/v3/products/10"]),
        "id": 21, "parent_id": 20, "type": "variation", "sku": "CHILD-21",
        "attributes": [{"id": 0, "name": "Flavor", "option": "A"}],
    }
    return remote.items["https://child.test/wp-json/wc/v3/products/20/variations/21"]


def test_variation_edits_exact_leaf_without_expanding_parent_or_siblings(system):
    leaf = add_variation(system)
    response = update(system, {"regular_price": "15"}, product_id=20, variation_id=21)
    assert response.get_json()["success"] is True
    assert leaf["regular_price"] == "15"
    assert response.get_json()["identity"]["parent_id"] == 20
    assert all(url.endswith("/variations/21") for method, url, _ in system["remote"].calls if method == "PUT")
    assert system["remote"].items["https://child.test/wp-json/wc/v3/products/20"]["regular_price"] == "12.00"


def test_variation_parent_mismatch_blocks_preview_and_write(system):
    leaf = add_variation(system)
    leaf["parent_id"] = 999
    response = preview(system, 20, 21)
    assert response.status_code == 409 and response.get_json()["code"] == "RESOURCE_MISMATCH"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


def test_standard_woo_variation_without_parent_id_uses_verified_exact_parent_route(system):
    leaf = add_variation(system)
    leaf.pop("parent_id")
    leaf["_links"] = {"up": [{"href": "https://child.test/wp-json/wc/v3/products/20"}]}
    snapshot = preview(system, 20, 21).get_json()
    assert snapshot["identity"]["parent_id"] == 20
    response = update(system, {"regular_price": "15"}, product_id=20, variation_id=21, snapshot=snapshot)
    assert response.get_json()["success"] is True
    assert all(url.endswith("/variations/21") for method, url, _ in system["remote"].calls if method == "PUT")


def test_variation_missing_parent_id_and_up_evidence_blocks_before_write(system):
    leaf = add_variation(system)
    snapshot = preview(system, 20, 21).get_json()
    leaf.pop("parent_id")
    response = update(system, {"regular_price": "15"}, product_id=20, variation_id=21, snapshot=snapshot)
    assert response.status_code == 409 and response.get_json()["code"] == "PARENT_EVIDENCE_MISSING"
    assert response.get_json()["write_started"] is False
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


@pytest.mark.parametrize("href,allowed", [("https://child.test/wp-json/wc/v3/products/20", True),
                                       ("https://child.test/wp-json/wc/v3/products/999", False),
                                       ("https://master.test/wp-json/wc/v3/products/20", False)])
def test_standard_woo_variation_parent_up_link_is_bound_to_actual_endpoint(system, href, allowed):
    leaf = add_variation(system)
    leaf.pop("parent_id")
    leaf["_links"] = {"up": [{"href": href}]}
    response = preview(system, 20, 21)
    assert response.status_code == (200 if allowed else 409)
    if not allowed:
        assert response.get_json()["code"] == "RESOURCE_MISMATCH" and response.get_json()["write_started"] is False


def test_variation_up_link_changed_inside_guard_stops_before_put(system):
    leaf = add_variation(system)
    leaf.pop("parent_id")
    leaf["_links"] = {"up": [{"href": "https://child.test/wp-json/wc/v3/products/20"}]}
    system["remote"].preflight_hook = lambda: leaf.update(_links={"up": [{"href": "https://child.test/wp-json/wc/v3/products/999"}]})
    response = update(system, {"regular_price": "15"}, product_id=20, variation_id=21)
    assert response.status_code == 409 and response.get_json()["code"] == "RESOURCE_MISMATCH"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


@pytest.mark.parametrize("changes", [{"stock_quantity": 0}, {"stock_status": "outofstock"}, {"manage_stock": False, "stock_status": "outofstock"}, {"manage_stock": True}])
def test_inherited_stock_actions_refused_before_any_parent_or_leaf_mutation(system, changes):
    leaf = add_variation(system)
    leaf["manage_stock"] = "parent"
    leaf["stock_quantity"] = None
    response = update(system, changes, product_id=20, variation_id=21)
    assert response.status_code == 409 and response.get_json()["code"] == "INHERITED_STOCK"
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)
    assert system["remote"].items["https://child.test/wp-json/wc/v3/products/20"]["stock_quantity"] == 8


def test_inherited_stock_allows_only_explicit_independent_management_and_quantity(system):
    leaf = add_variation(system)
    leaf.update(manage_stock="parent", stock_quantity=None)
    response = update(system, {"manage_stock": True, "stock_quantity": 3}, product_id=20, variation_id=21)
    assert response.get_json()["success"] is True
    assert leaf["manage_stock"] is True and leaf["stock_quantity"] == 3
    assert system["remote"].items["https://child.test/wp-json/wc/v3/products/20"]["stock_quantity"] == 8


def test_inherited_stock_price_only_remains_editable(system):
    leaf = add_variation(system)
    leaf.update(manage_stock="parent", stock_quantity=None)
    response = update(system, {"sale_price": "10.5"}, product_id=20, variation_id=21)
    assert response.get_json()["success"] is True and leaf["manage_stock"] == "parent"


@pytest.mark.parametrize("items", [[], [{"id": 10, "type": "simple", "attributes": []}, {"id": 10, "type": "simple", "attributes": []}], [{"id": 99, "type": "simple", "attributes": []}]])
def test_clone_incomplete_or_wrong_source_collection_does_not_queue(system, items):
    system["remote"].collection_override = items
    response = system["client"].post("/api/product-manager/catalog-clone", json={
        "source_site_id": 1, "target_site_id": 2, "product_ids": [10],
    }, headers={"X-PM-CSRF": system["csrf"]})
    assert response.status_code == 409 and response.get_json()["code"] == "SOURCE_NOT_VERIFIED"
    assert system["jobs"] == []


def test_direct_worker_stops_using_initial_config_when_credentials_change(system):
    connection = system["connect"]()
    init_product_clone_jobs(connection)
    queued = enqueue_clone_job(connection, source_site_id=1, target_site_id=2, product_ids=[10, 11],
        options={"_catalog_direct_sites": True, "_catalog_source_url": "https://child.test", "_catalog_target_url": "https://target.test"},
        target_url="https://target.test", created_by_id="2", created_by_name="Michael")
    job = claim_clone_job(connection, "test")
    calls = []
    def clone_one(*args):
        calls.append(args)
        connection.execute("UPDATE sites SET consumer_key='changed' WHERE id=2")
        connection.commit()
        return {"new_id": 100, "name": "Cloned"}
    results = worker.process_clone_job(connection, job, clone_one=clone_one,
        resolve_site=lambda *_: pytest.fail("Direct jobs must never resolve Master"))
    assert len(calls) == 1 and len(results["success"]) == 1 and len(results["failed"]) == 1
    assert "凭据已变化" in results["failed"][0]["error"]
    assert get_clone_job(connection, queued["id"])["status"] == "partial_failed"
    connection.close()


@pytest.mark.parametrize("point,success", [("before", False), ("after", True)])
def test_ambiguous_put_reads_back_without_replaying(system, point, success):
    system["remote"].timeout_put = point
    response = update(system, {"regular_price": "15"})
    data = response.get_json()
    assert data["success"] is success
    assert sum(method == "PUT" for method, _, _ in system["remote"].calls) == 1
    if not success:
        assert response.status_code == 409 and data["write_started"] is True
        assert data["verification_status"] == "unconfirmed"
        assert system["audits"][-1]["status"] == "unconfirmed"


def test_direct_worker_transient_callback_rechecks_right_before_external_write(system):
    connection = system["connect"]()
    init_product_clone_jobs(connection)
    queued = enqueue_clone_job(connection, source_site_id=1, target_site_id=2, product_ids=[10],
        options={"_catalog_direct_sites": True, "_catalog_source_url": "https://child.test", "_catalog_target_url": "https://target.test"},
        target_url="https://target.test", created_by_id="2", created_by_name="Michael")
    job = claim_clone_job(connection, "test")
    remote_writes = []
    def clone_one(*args):
        options = args[-1]
        connection.execute("UPDATE users SET can_manage_products=0 WHERE id=2")
        connection.commit()
        options["_catalog_write_check"]()
        remote_writes.append("POST")
        return {"new_id": 100}
    results = worker.process_clone_job(connection, job, clone_one=clone_one,
        resolve_site=lambda *_: pytest.fail("Direct jobs must never resolve Master"))
    assert remote_writes == [] and len(results["failed"]) == 1
    assert "权限已撤销" in results["failed"][0]["error"]
    saved = get_clone_job(connection, queued["id"])
    assert "_catalog_write_check" not in saved["options"]
    connection.close()


def test_direct_worker_partial_clone_is_failed_with_existing_target_evidence(system):
    connection = system["connect"]()
    init_product_clone_jobs(connection)
    queued = enqueue_clone_job(connection, source_site_id=1, target_site_id=2, product_ids=[10],
        options={"_catalog_direct_sites": True, "_catalog_source_url": "https://child.test", "_catalog_target_url": "https://target.test"},
        target_url="https://target.test", created_by_id="2", created_by_name="Michael")
    job = claim_clone_job(connection, "test")
    worker.process_clone_job(connection, job,
        clone_one=lambda *_: {"new_id": 100, "partial_clone": True, "warnings": ["1 variation failed"]},
        resolve_site=lambda *_: pytest.fail("Direct jobs must never resolve Master"))
    saved = get_clone_job(connection, queued["id"])
    assert saved["status"] == "failed" and saved["success_count"] == 0
    partial = saved["results"]["failed"][0]
    assert partial["partial_clone"] is True and partial["target_id"] == 100
    assert partial["warnings"] == ["1 variation failed"] and "不要直接重复克隆" in partial["error"]
    connection.close()


def test_preview_bridge_evidence_excludes_all_other_metadata(system):
    target = system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]
    target["meta_data"] = [{"key": "private_plugin_key", "value": "DO-NOT-EXPOSE"},
        {"key": "wcms_stock_manage", "value": "yes"}, {"key": "wcms_stock_qty", "value": 8}]
    response = preview(system)
    assert response.get_json()["stock_bridge"]["values"] == {"wcms_stock_manage": "yes", "wcms_stock_qty": 8}
    assert "DO-NOT-EXPOSE" not in response.get_data(as_text=True)


def setup_stock_bridge(system, *, quantity=0, status="outofstock"):
    item = system["remote"].items["https://child.test/wp-json/wc/v3/products/10"]
    item.update(manage_stock=True, stock_quantity=quantity, stock_status=status,
        meta_data=[{"key": "wcms_stock_manage", "value": "yes"},
                   {"key": "wcms_stock_qty", "value": quantity}, {"key": "wcms_stock_status", "value": status}])
    return item


@pytest.mark.parametrize("quantity,threshold,backorders,initial_status,final_status", [
    (10, 0, False, "outofstock", "instock"), (2, 5, False, "instock", "outofstock"),
    (0, 0, False, "instock", "outofstock"), (0, 0, True, "instock", "onbackorder"),
])
def test_bridge_uses_site_authoritative_threshold_and_backorder_status(system, quantity, threshold, backorders, initial_status, final_status):
    target = setup_stock_bridge(system, quantity=0 if initial_status == "outofstock" else threshold + 1, status=initial_status)
    remote = system["remote"]
    remote.auto_stock_status, remote.no_stock_threshold, remote.allow_backorders = True, threshold, backorders
    data = update(system, {"manage_stock": True, "stock_quantity": quantity}).get_json()
    assert data["success"] is True
    assert target["stock_status"] == final_status
    assert data["stock_bridge"]["values"]["wcms_stock_status"] == final_status
    assert remote.write_leases == [True, True]
    writes = [kwargs["json"] for method, _, kwargs in remote.calls if method == "PUT"]
    assert len(writes) == 2 and set(writes[-1]) == {"meta_data"}


def test_legacy_bridge_restoring_core_after_each_save_keeps_requested_quantity(system):
    target = setup_stock_bridge(system)
    system["remote"].restore_from_bridge = True
    response = update(system, {"manage_stock": True, "stock_quantity": 10})
    assert response.get_json()["success"] is True
    assert target["stock_quantity"] == 10 and target["manage_stock"] is True
    assert response.get_json()["stock_bridge"]["values"]["wcms_stock_qty"] == 10


@pytest.mark.parametrize("point,success", [("before", False), ("after", True)])
def test_bridge_only_timeout_is_read_back_once_and_never_replayed(system, point, success):
    target = setup_stock_bridge(system)
    system["remote"].auto_stock_status = True
    system["remote"].bridge_timeout = point
    response = update(system, {"manage_stock": True, "stock_quantity": 10})
    data = response.get_json()
    assert data["success"] is success
    writes = [kwargs["json"] for method, _, kwargs in system["remote"].calls if method == "PUT"]
    assert len(writes) == 2 and sum(set(payload) == {"meta_data"} for payload in writes) == 1
    if not success:
        assert response.status_code == 409 and data["code"] == "BRIDGE_NOT_VERIFIED"
        assert data["verification_status"] == "unconfirmed" and data["write_started"] is True
        assert data["expected_stock_bridge"]["wcms_stock_status"] == target["stock_status"] == "instock"
        assert data["stock_bridge"]["values"]["wcms_stock_status"] == "outofstock"
        assert system["audits"][-1]["status"] == "unconfirmed"


@pytest.mark.parametrize("malformed", ["duplicate", "object"])
def test_bridge_ambiguous_or_invalid_values_block_stock_write_before_put(system, malformed):
    target = setup_stock_bridge(system)
    if malformed == "duplicate":
        target["meta_data"].append({"key": "wcms_stock_status", "value": "instock"})
    else:
        target["meta_data"][1]["value"] = {"private": "never expose"}
    snapshot = preview(system).get_json()
    assert snapshot["stock_bridge"]["valid"] is False
    response = update(system, {"manage_stock": True, "stock_quantity": 10}, snapshot=snapshot)
    assert response.status_code == 409 and response.get_json()["code"] == "BRIDGE_INVALID"
    assert response.get_json()["write_started"] is False
    assert not any(method == "PUT" for method, _, _ in system["remote"].calls)


def clone_app_functions(monkeypatch, source, write_check):
    """Exercise app's real write boundaries without its schema initialization."""
    calls, writes = [], []
    def get(url, **kwargs):
        calls.append(url)
        if url.endswith("/products/10"):
            return Response(source)
        if url.endswith("/products/10/variations"):
            return Response([{"id": 11, "parent_id": 10, "sku": "LEAF-11", "attributes": []}])
        return Response([])
    def post(url, **kwargs):
        writes.append(url)
        return Response({"id": 100, **kwargs["json"]}, 201)
    monkeypatch.setattr(requests.sessions.Session, "request", lambda *_args, **_kwargs: pytest.fail("Real HTTP forbidden"))
    monkeypatch.setattr(requests, "get", get)
    monkeypatch.setattr(requests, "post", post)
    scope = {"_WC_HEADERS": {"User-Agent": "catalog-offline-test"}, "_parse_wc_response": parse_wc_response,
             "_resolve_taxonomy_on_target": lambda *_args: ([], []), "build_clone_sku": build_clone_sku,
             "make_clone_suffix": make_clone_suffix, "normalize_clone_suffix": normalize_clone_suffix}
    tree = ast.parse((Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8"))
    definitions = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in {"_clone_one_product", "_clone_variations"}]
    assert len(definitions) == 2
    exec(compile(ast.Module(body=definitions, type_ignores=[]), "app.py", "exec"), scope)
    identity = edit.catalog_item_identity({**source, "id": 10, "sku": "ORIGINAL"})
    result = scope["_clone_one_product"]("https://source.invalid", "fixture", "fixture", "https://target.invalid", "fixture", "fixture", 10,
        {"_catalog_direct_sites": True, "_catalog_source_identities": {"10": identity}, "_catalog_write_check": write_check,
         "collision_mode": "clone_as_new", "clone_sku_suffix": "NEW-TEST", "include_variations": True,
         "include_images": False, "status_on_target": "draft"})
    return result, calls, writes


@pytest.mark.parametrize("source", [{"id": 99, "type": "simple", "sku": "ORIGINAL", "attributes": []},
                                    {"id": 10, "type": "simple", "sku": "CHANGED", "attributes": []}])
def test_direct_app_clone_wrong_source_identity_never_reaches_target(monkeypatch, source):
    result, calls, writes = clone_app_functions(monkeypatch, source, lambda: None)
    assert "身份已变化" in result["error"]
    assert len(calls) == 1 and writes == []


def test_direct_app_clone_permission_callback_blocks_parent_post(monkeypatch):
    def denied():
        raise ValueError("权限已撤销")
    result, _, writes = clone_app_functions(monkeypatch, {"id": 10, "type": "simple", "sku": "ORIGINAL", "attributes": []}, denied)
    assert "权限已撤销" in result["error"] and writes == []


def test_direct_app_clone_permission_callback_blocks_variation_post(monkeypatch):
    checks = []
    def check():
        checks.append(True)
        if len(checks) > 1:
            raise ValueError("权限已撤销")
    result, _, writes = clone_app_functions(monkeypatch, {"id": 10, "type": "variable", "sku": "ORIGINAL", "attributes": []}, check)
    assert len(writes) == 1 and not writes[0].endswith("/variations")
    assert any("1 失败" in warning for warning in result["warnings"])
