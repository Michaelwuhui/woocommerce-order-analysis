"""Release contracts for the actual clone functions, without importing app.

All HTTP is synthetic and Session.request is forbidden. Compiling only the
functions prevents application startup from opening any configured database.
"""
import ast
from copy import deepcopy
import json
from pathlib import Path
from types import SimpleNamespace

from flask import Flask, jsonify
import pytest
import requests

from product_clone_sku import build_clone_sku, make_clone_suffix, normalize_clone_suffix
from product_manager_catalog_edit import catalog_item_identity
from product_manager_service import parse_wc_response
from test_product_manager_catalog_edit import system, update


class Response:
    def __init__(self, body, status=200):
        self.body = deepcopy(body)
        self.status_code = status
        self.text = json.dumps(body)

    def json(self):
        return deepcopy(self.body)


class Remote:
    def __init__(self):
        self.source = {"id": 100, "name": "Fixture", "type": "simple", "sku": "SOURCE",
                       "attributes": [{"id": 0, "name": "Flavor", "options": ["Mint"]}]}
        self.variations = [{"id": 101, "parent_id": 100, "sku": "SOURCE-MINT", "attributes": []}]
        self.created = []
        self.puts = []
        self.revoked = False
        self.revoke_on_source_read = False
        self.revoke_after_parent_write = False
        self.revoke_after_inline_put = False
        self.fail_variation_page = False
        self.check_count = 0

    def check_write(self):
        self.check_count += 1
        if self.revoked:
            raise ValueError("live permission revoked")

    def get(self, url, **kwargs):
        if url == "https://source.invalid/wp-json/wc/v3/products/100":
            if self.revoke_on_source_read:
                self.revoked = True
            return Response(self.source)
        if url == "https://source.invalid/wp-json/wc/v3/products/100/variations":
            if self.fail_variation_page:
                return Response({"code": "unavailable", "message": "Source unavailable"}, 503)
            return Response(self.variations)
        assert url == "https://target.invalid/wp-json/wc/v3/products"
        return Response([])

    def post(self, url, **kwargs):
        assert url.startswith("https://target.invalid/wp-json/wc/v3/products")
        payload = deepcopy(kwargs["json"])
        self.created.append((url, payload))
        if self.revoke_after_parent_write and not url.endswith("/variations"):
            self.revoked = True
        return Response({"id": 1000 + len(self.created), **payload})

    def put(self, url, **kwargs):
        assert url.startswith("https://target.invalid/wp-json/wc/v3/products/")
        self.puts.append((url, deepcopy(kwargs["json"])))
        if self.revoke_after_inline_put:
            self.revoked = True
        return Response({"id": 1001, "images": [{"id": 4001, "src": "https://target.invalid/media/mint.jpg"}]})


@pytest.fixture
def clone_contract(monkeypatch):
    remote = Remote()
    monkeypatch.setattr(requests.sessions.Session, "request",
                        lambda *args, **kwargs: pytest.fail("Real HTTP is forbidden"))
    monkeypatch.setattr(requests, "get", remote.get)
    monkeypatch.setattr(requests, "post", remote.post)
    monkeypatch.setattr(requests, "put", remote.put)
    scope = {"_WC_HEADERS": {"User-Agent": "offline-contract"},
             "_parse_wc_response": parse_wc_response,
             "_resolve_taxonomy_on_target": lambda *args: ([], []),
             "build_clone_sku": build_clone_sku, "make_clone_suffix": make_clone_suffix,
             "normalize_clone_suffix": normalize_clone_suffix}
    tree = ast.parse((Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8"))
    functions = [node for node in tree.body if isinstance(node, ast.FunctionDef)
                 and (node.name in {"_clone_one_product", "_clone_variations",
                                    "_migrate_inline_images_after_clone", "_extract_source_image_urls"}
                      or node.name.startswith("_catalog_clone"))]
    exec(compile(ast.Module(body=functions, type_ignores=[]), "app.py", "exec"), scope)

    def run(*, expected=None, direct=True, **options):
        defaults = {"collision_mode": "clone_as_new", "clone_sku_suffix": "NEW-CONTRACT",
                    "include_images": False, "include_variations": True, "status_on_target": "draft"}
        if direct:
            defaults.update({"_catalog_direct_sites": True,
                             "_catalog_source_identities": {"100": expected or catalog_item_identity(remote.source)},
                             "_catalog_write_check": remote.check_write})
        defaults.update(options)
        return scope["_clone_one_product"]("https://source.invalid", "offline", "offline",
                                            "https://target.invalid", "offline", "offline", 100, defaults)
    return remote, run


@pytest.mark.parametrize("mutation", [
    {"id": 999}, {"sku": "DIFFERENT"}, {"type": "variable"},
    {"attributes": [{"id": 0, "name": "Flavor", "options": ["Peach"]}]},
])
def test_direct_changed_source_identity_never_creates_target(clone_contract, mutation):
    remote, run = clone_contract
    expected = catalog_item_identity(remote.source)
    remote.source.update(mutation)
    result = run(expected=expected)
    assert result.get("error")
    assert remote.created == []


def test_direct_matching_source_identity_creates_exact_real_target(clone_contract):
    remote, run = clone_contract
    result = run()
    assert not result.get("error")
    assert len(remote.created) == 1
    assert remote.created[0][0] == "https://target.invalid/wp-json/wc/v3/products"
    assert remote.created[0][1]["status"] == "draft"


def test_direct_revoked_during_source_read_cannot_reach_parent_post(clone_contract):
    remote, run = clone_contract
    remote.revoke_on_source_read = True
    result = run()
    assert result.get("error")
    assert remote.created == []
    assert remote.check_count > 0


def test_direct_revoked_after_parent_post_cannot_create_variations(clone_contract):
    remote, run = clone_contract
    remote.source["type"] = "variable"
    remote.revoke_after_parent_write = True
    result = run()
    assert len(remote.created) == 1
    assert result.get("error") or any("失败" in warning for warning in result.get("warnings", []))
    assert remote.check_count >= 2


def test_legacy_clone_does_not_depend_on_new_live_callback(clone_contract):
    remote, run = clone_contract
    remote.revoked = True
    result = run(direct=False)
    assert not result.get("error")
    assert len(remote.created) == 1
    assert remote.check_count == 0


def test_direct_revoked_after_parent_post_cannot_sideload_inline_images(clone_contract):
    remote, run = clone_contract
    remote.source["description"] = '<img src="https://source.invalid/uploads/mint.jpg">'
    remote.revoke_after_parent_write = True
    result = run(include_images=True)
    assert len(remote.created) == 1
    assert remote.puts == []
    assert result["partial_clone"] is True
    assert any("迁移失败" in warning for warning in result.get("warnings", []))


def test_direct_rechecks_authority_before_second_inline_put(clone_contract):
    remote, run = clone_contract
    remote.source["description"] = '<img src="https://source.invalid/uploads/mint.jpg">'
    remote.revoke_after_inline_put = True
    result = run(include_images=True)
    assert len(remote.created) == 1
    assert len(remote.puts) == 1
    assert result["partial_clone"] is True
    assert any("回写失败" in warning for warning in result.get("warnings", []))


def test_direct_source_variation_page_failure_reports_partial_parent_without_leaf_posts(clone_contract):
    remote, run = clone_contract
    remote.source["type"] = "variable"
    remote.fail_variation_page = True
    result = run()
    assert len(remote.created) == 1
    assert result["partial_clone"] is True
    assert result["new_id"] == 1001
    assert any("读取失败" in warning for warning in result["warnings"])


def test_legacy_single_variation_list_keeps_its_original_routing(monkeypatch):
    monkeypatch.setattr(requests.sessions.Session, "request",
                        lambda *args, **kwargs: pytest.fail("Real HTTP is forbidden"))
    calls = []

    def get(url, **kwargs):
        calls.append((url, kwargs))
        return Response([{"id": 101, "sku": "LEGACY", "manage_stock": True,
                          "stock_quantity": 8, "stock_status": "instock", "attributes": []}])

    monkeypatch.setattr(requests, "get", get)
    flask_app = Flask("offline-legacy-list")
    connection = SimpleNamespace(close=lambda: None)
    scope = {"get_db_connection": lambda: connection,
             "_resolve_site_for_product_edit": lambda *_: ({"id": 1}, "https://master.invalid", "offline", "offline"),
             "_parse_wc_response": parse_wc_response, "jsonify": jsonify, "app": flask_app}
    tree = ast.parse((Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8"))
    function = next(node for node in tree.body if isinstance(node, ast.FunctionDef)
                    and node.name == "product_manager_list_variations")
    function.decorator_list = []
    exec(compile(ast.Module(body=[function], type_ignores=[]), "app.py", "exec"), scope)
    with flask_app.test_request_context():
        response = scope["product_manager_list_variations"](1, 100)
    assert response.get_json()["variations"][0]["id"] == 101
    assert calls[0][0] == "https://master.invalid/wp-json/wc/v3/products/100/variations"
    assert calls[0][1]["params"] == {"page": 1, "per_page": 100}


@pytest.mark.parametrize("quantity, initial_status", [(10, "outofstock"), (0, "instock")])
def test_direct_stock_bridge_cannot_verify_old_status_after_woo_normalizes_quantity(system, quantity, initial_status):
    """Woo validate_props derives status on save; a plain dict update hides this."""
    remote = system["remote"]
    target = remote.items["https://child.test/wp-json/wc/v3/products/10"]
    target.update(manage_stock=True, stock_quantity=8, stock_status=initial_status,
                  meta_data=[{"key": "wcms_stock_manage", "value": "yes"},
                             {"key": "wcms_stock_qty", "value": 8},
                             {"key": "wcms_stock_status", "value": initial_status}])
    original_put = remote.put

    def woo_put(url, **kwargs):
        original_put(url, **kwargs)
        item = remote.items[url]
        if item.get("manage_stock") is True:
            item["stock_status"] = "instock" if item.get("stock_quantity", 0) > 0 else "outofstock"
        return Response(item)

    remote.put = woo_put
    response = update(system, {"manage_stock": True, "stock_quantity": quantity})
    body = response.get_json()
    metadata = {row["key"]: row["value"] for row in target["meta_data"]}
    assert body.get("success") is not True or metadata["wcms_stock_status"] == target["stock_status"], \
        "A verified restore must not retain the previous status in the WCMS bridge"


@pytest.mark.parametrize("normalize_after_bridge", [False, True])
def test_direct_stock_restore_survives_child_hook_restoring_core_from_bridge(system, normalize_after_bridge):
    """A core-only first PUT would be undone by this existing WCMS save hook."""
    remote = system["remote"]
    target = remote.items["https://child.test/wp-json/wc/v3/products/10"]
    target.update(manage_stock=False, stock_quantity=None, stock_status="outofstock",
                  meta_data=[{"key": "wcms_stock_manage", "value": "no"},
                             {"key": "wcms_stock_qty", "value": 0},
                             {"key": "wcms_stock_status", "value": "outofstock"}])
    original_put = remote.put

    def bridge_restoring_put(url, **kwargs):
        original_put(url, **kwargs)
        item = remote.items[url]
        metadata = {row["key"]: row["value"] for row in item["meta_data"]}
        item["manage_stock"] = metadata["wcms_stock_manage"] == "yes"
        item["stock_quantity"] = int(metadata["wcms_stock_qty"]) if item["manage_stock"] else None
        item["stock_status"] = metadata["wcms_stock_status"]
        if normalize_after_bridge and item["manage_stock"]:
            item["stock_status"] = "instock" if item["stock_quantity"] > 0 else "outofstock"
        return Response(item)

    remote.put = bridge_restoring_put
    response = update(system, {"manage_stock": True, "stock_quantity": 10})
    body = response.get_json()
    assert body["success"] is True
    assert target["manage_stock"] is True and target["stock_quantity"] == 10
    metadata = {row["key"]: row["value"] for row in target["meta_data"]}
    assert metadata["wcms_stock_status"] == target["stock_status"]
    assert metadata["wcms_stock_qty"] == target["stock_quantity"]
    writes = [kwargs["json"] for method, _, kwargs in remote.calls if method == "PUT"]
    assert writes[0]["meta_data"], "First save must carry new manage/quantity for existing child hooks"
    if normalize_after_bridge:
        assert len(writes) == 2 and set(writes[1]) == {"meta_data"}
