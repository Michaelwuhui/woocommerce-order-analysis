"""Offline catalog API contract checks; every Woo read is replaced by a fake."""
from functools import wraps
import sqlite3

from flask import Flask, jsonify
from flask_login import LoginManager, UserMixin, current_user
import pytest
import requests

import product_manager_catalog as catalog


class User(UserMixin):
    id = "1"
    name = "Michael"
    own_scoped = False
    may_manage = True

    def product_manager_own_scoped(self):
        return self.own_scoped


class Response:
    def __init__(self, payload, *, total=None, pages=None, status=200, headers=None):
        self.payload = payload
        self.status_code = status
        if headers is not None:
            self.headers = headers
        else:
            self.headers = {}
            if total is not None:
                self.headers["X-WP-Total"] = str(total)
            if pages is not None:
                self.headers["X-WP-TotalPages"] = str(pages)

    def json(self):
        if isinstance(self.payload, Exception):
            raise self.payload
        return self.payload


def product(identity, **updates):
    return {
        "id": identity, "name": f"Device {identity}", "sku": f"SKU-{identity}",
        "type": "simple", "status": "publish", "attributes": [], "brands": [], "variations": [],
        "stock_status": "instock", "stock_quantity": 8, "manage_stock": True,
        "price": "14.50", "regular_price": "15.00", "sale_price": "14.50",
        "permalink": f"https://child.test/product/{identity}", **updates,
    }


@pytest.fixture
def system(tmp_path, monkeypatch):
    db_path = tmp_path / "catalog.db"

    def connect():
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        return conn

    conn = connect()
    conn.execute("""CREATE TABLE sites (
        id INTEGER PRIMARY KEY, url TEXT, manager TEXT, country TEXT,
        consumer_key TEXT, consumer_secret TEXT, product_master_id INTEGER
    )""")
    conn.executemany("INSERT INTO sites VALUES (?, ?, ?, ?, ?, ?, ?)", [
        (1, "https://child.test", " Michael ", "PL", "private_ck_1", "private_cs_1", 97),
        (2, "https://other.test", "Jane", "CZ", "private_ck_2", "private_cs_2", None),
        (3, "https://missing.test", "Michael", "PL", None, None, 97),
        (4, "https://vacant.test", "", "PL", "private_ck_4", "private_cs_4", None),
    ])
    conn.commit()
    conn.close()
    app = Flask(__name__)
    app.config.update(TESTING=True, SECRET_KEY="offline-catalog-test")
    login_manager = LoginManager(app)
    user = User()
    login_manager.user_loader(lambda _identity: user)

    def manager_required(function):
        @wraps(function)
        def decorated(*args, **kwargs):
            if not current_user.may_manage:
                return jsonify({"error": "无产品管理权限"}), 403
            return function(*args, **kwargs)
        return decorated

    recognition = {"brands": [], "series": []}
    recognition_calls = []

    def recognition_loader():
        recognition_calls.append(True)
        return recognition

    app.register_blueprint(catalog.create_catalog_blueprint(connect, manager_required, recognition_loader))
    client = app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = "1"
        session["_fresh"] = True
    calls = []
    responses = []

    def fake_get(url, **kwargs):
        calls.append((url, kwargs))
        assert responses, "Unexpected external read"
        response = responses.pop(0)
        if isinstance(response, Exception):
            raise response
        return response

    monkeypatch.setattr(catalog.requests, "get", fake_get)
    return {
        "client": client, "user": user, "responses": responses, "calls": calls,
        "connect": connect, "recognition": recognition, "recognition_calls": recognition_calls,
    }


def page(system, **params):
    return system["client"].get("/api/product-manager/catalog-page", query_string={"site_id": 1, **params})


def test_sites_preserve_unconfigured_children_and_do_not_expose_credentials(system):
    response = system["client"].get("/api/product-manager/catalog-sites")
    data = response.get_json()
    assert response.status_code == 200
    assert {site["id"] for site in data["sites"]} == {1, 2, 3, 4}
    missing = next(site for site in data["sites"] if site["id"] == 3)
    assert missing["configured"] is False
    assert missing["product_master_id"] == 97
    assert "private_ck" not in response.get_data(as_text=True)
    assert "consumer_key" not in response.get_data(as_text=True)
    assert system["calls"] == []


def test_misconfigured_url_does_not_expose_embedded_credentials(system):
    conn = system["connect"]()
    conn.execute("UPDATE sites SET url = ? WHERE id = 3", ("https://user:secret@missing.test/?consumer_key=ck_sensitive",))
    conn.commit()
    conn.close()
    response = system["client"].get("/api/product-manager/catalog-sites")
    missing = next(site for site in response.get_json()["sites"] if site["id"] == 3)
    assert missing["url"] == "https://missing.test"
    assert missing["configured"] is False
    assert "secret" not in response.get_data(as_text=True)
    assert "ck_sensitive" not in response.get_data(as_text=True)


def test_reads_actual_child_and_scans_parent_directory_without_woo_search(system):
    system["responses"].append(Response([
        product(10, name="Generic variable", type="variable", variations=[101]),
        product(11, name="Mango Device"),
    ], total=2, pages=1))
    data = page(system, search="mango").get_json()
    assert data["complete_page"] is True
    assert [row["product_id"] for row in data["rows"]] == [11]
    assert [parent["id"] for parent in data["variable_products"]] == [10]
    assert data["variable_products"][0]["variation_ids"] == [101]
    assert data["source_ids"] == [10, 11]
    assert data["scanned"] == 2
    url, kwargs = system["calls"][0]
    assert url == "https://child.test/wp-json/wc/v3/products"
    assert kwargs["auth"] == ("private_ck_1", "private_cs_1")
    assert kwargs["timeout"] == (5, 25)
    assert kwargs["params"] == {"page": 1, "per_page": 50, "status": "any", "orderby": "id", "order": "asc"}
    assert "search" not in kwargs["params"]


@pytest.mark.parametrize("attribute_name", ["口味", "Flavor", "Flavour", "Flavour Profiles", "Smak", "Příchuť", "pa_flavor", "pa_prichut", "Íz", "Ízcsoport", "pa_íz", "Ízesítés", "pa_izesites"])
def test_simple_flavor_attributes_match_even_with_generic_product_name(system, attribute_name):
    system["responses"].append(Response([
        product(1, attributes=[{"name": attribute_name, "options": ["Míxed &amp; Berries"]}]),
    ], total=1, pages=1))
    data = page(system, search=" mixed   & berries ").get_json()
    assert len(data["rows"]) == 1
    assert data["rows"][0]["flavors"] == ["Míxed & Berries"]
    assert data["has_more"] is False


def test_variation_uses_own_attribute_not_sibling_options(system):
    parent = product(10, name="Generic Device", type="variable", variations=[101, 102], attributes=[
        {"id": 5, "name": "Příchuť", "variation": True, "options": ["Cherry", "Blue Razz"]},
        {"id": 6, "name": "Brand", "options": ["Example Brand"]},
    ])
    system["responses"].extend([
        Response(parent),
        Response([
            product(101, name="", attributes=[{"id": 5, "option": "Cherry"}]),
            product(102, name="", attributes=[{"id": 5, "option": "Blue Razz"}]),
        ], total=2, pages=1),
    ])
    data = page(system, parent_id=10, search="Cherry").get_json()
    assert data["source_ids"] == [101, 102]
    assert len(data["rows"]) == 1
    leaf = data["rows"][0]
    assert leaf["product_id"] == 10 and leaf["variation_id"] == 101
    assert leaf["flavors"] == ["Cherry"]
    assert leaf["brands"] == ["Example Brand"]
    assert leaf["name"] == "Generic Device — Cherry"
    assert "Blue Razz" not in str(leaf)
    assert system["calls"][0][0] == "https://child.test/wp-json/wc/v3/products/10"
    assert system["calls"][1][0] == "https://child.test/wp-json/wc/v3/products/10/variations"


@pytest.mark.parametrize("attribute_identity", [
    {"id": 7, "name": "Ízesítés"},
    {"id": 7, "slug": "pa_izesites"},
])
def test_hungarian_flavoring_preserves_specific_variant_flavor(system, attribute_identity):
    parent = product(10, name="Merry Mi Blade 30000 Puffs", type="variable", variations=[101, 102], attributes=[
        {**attribute_identity, "variation": True, "options": ["Aloe blackcurrant", "Black Dragon Ice"]},
    ], brands=[{"name": "Merrymi"}])
    leaves = [
        product(101, name="", attributes=[{**attribute_identity, "option": "Aloe blackcurrant"}]),
        product(102, name="", attributes=[{**attribute_identity, "option": "Black Dragon Ice"}]),
    ]
    system["responses"].extend([Response(parent), Response(leaves, total=2, pages=1)])
    by_brand = page(system, parent_id=10, query_mode="brand", brand="Merrymi").get_json()
    assert [row["flavors"] for row in by_brand["rows"]] == [["Aloe blackcurrant"], ["Black Dragon Ice"]]
    system["responses"].extend([Response(parent), Response(leaves, total=2, pages=1)])
    by_flavor = page(system, parent_id=10, query_mode="flavor", search="aloe blackcurrant").get_json()
    assert [row["variation_id"] for row in by_flavor["rows"]] == [101]
    assert by_flavor["rows"][0]["flavors"] == ["Aloe blackcurrant"]


def test_parent_identity_is_read_from_site_not_client_claims(system):
    parent = product(10, name="Actual Product", type="variable", brands=[{"id": 1, "name": "Woo Brand"}])
    system["responses"].extend([Response(parent), Response([product(101, name="Actual Product - Lemon")], total=1, pages=1)])
    data = page(system, parent_id=10, parent_name="Fake Product", parent_sku="Fake SKU").get_json()
    assert data["rows"][0]["product_name"] == "Actual Product"
    assert data["rows"][0]["brands"] == ["Woo Brand"]
    assert data["rows"][0]["flavors"] == ["Lemon"]
    assert "Fake" not in str(data)


def test_empty_variation_name_does_not_inherit_parsed_parent_flavor(system):
    parent = product(10, name="Device - Cherry", type="variable")
    system["responses"].extend([
        Response(parent), Response([product(101, name="", attributes=[{"name": "Flavor", "option": "Lemon"}])], total=1, pages=1),
    ])
    row = page(system, parent_id=10).get_json()["rows"][0]
    assert row["flavors"] == ["Lemon"]


@pytest.mark.parametrize("parent_attribute,leaf_attribute", [
    ({"id": 5, "name": "Flavour"}, {"id": 5, "option": ""}),
    ({"id": 0, "name": "Flavour Profiles"}, {"id": 0, "name": "flavour_profiles", "option": ""}),
    ({"id": 0, "name": "Flavour Profiles", "slug": "pa_flavour_profiles"}, {"id": 0, "slug": "flavour-profiles", "option": ""}),
    ({"id": 0, "name": "Íz"}, {"id": 0, "name": "iz", "option": " "}),
])
def test_wildcard_flavor_matches_each_supported_option_without_expanding_specific_sibling(system, parent_attribute, leaf_attribute):
    parent_attribute = {**parent_attribute, "variation": True, "options": ["Cherry", "Blue Razz"]}
    parent = product(10, name="Generic Product", type="variable", variations=[101, 102], attributes=[parent_attribute])
    specific_attribute = {**leaf_attribute, "option": "Blue Razz"}
    leaves = [product(101, name="", attributes=[leaf_attribute]), product(102, name="", attributes=[specific_attribute])]
    for query, expected in [("Cherry", [101]), ("Blue Razz", [101, 102])]:
        system["responses"].extend([Response(parent), Response(leaves, total=2, pages=1)])
        data = page(system, parent_id=10, search=query).get_json()
        assert data["complete_page"] is True
        assert [row["variation_id"] for row in data["rows"]] == expected
        shared = data["rows"][0]
        assert shared["flavor_scope"] == "any"
        assert shared["attributes"][0]["wildcard"] is True
        assert shared["flavors"] == ["Cherry", "Blue Razz"]
        if len(data["rows"]) == 2:
            specific = data["rows"][1]
            assert specific["flavor_scope"] == "specific"
            assert specific["attributes"][0]["wildcard"] is False
            assert specific["flavors"] == ["Blue Razz"]


@pytest.mark.parametrize("parent_attributes", [[], [{"name": "Flavor", "options": []}], [
    {"name": "Flavor", "options": ["Cherry"]}, {"name": "Flavor", "options": ["Lemon"]},
]])
def test_unresolvable_wildcard_is_incomplete_instead_of_zero_matches(system, parent_attributes):
    parent = product(10, type="variable", variations=[101], attributes=parent_attributes)
    system["responses"].extend([Response(parent), Response([
        product(101, attributes=[{"id": 0, "name": "Flavor", "option": ""}]),
    ], total=1, pages=1)])
    data = page(system, parent_id=10, search="Cherry").get_json()
    assert data["complete_page"] is False
    assert data["code"] == "unresolved_flavor_wildcard"


def test_source_statuses_include_all_scanned_variations_even_unmatched_drafts(system):
    # Woo parent.variations declares publish/private children; the all-status
    # variations endpoint additionally contains draft/pending/future children.
    parent = product(10, type="variable", variations=[101, 102])
    states = ["publish", "private", "draft", "pending", "future"]
    leaves = [product(101 + index, name="Generic leaf", status=state) for index, state in enumerate(states)]
    system["responses"].extend([Response(parent), Response(leaves, total=5, pages=1)])
    data = page(system, parent_id=10, search="Absent flavor").get_json()
    assert data["complete_page"] is True
    assert data["rows"] == []
    assert data["source_ids"] == [101, 102, 103, 104, 105]
    assert data["source_statuses"] == {str(101 + index): state for index, state in enumerate(states)}
    assert data["scanned"] == data["total"] == 5


def test_missing_source_status_is_not_treated_as_an_unpublished_extra(system):
    leaf = product(101)
    del leaf["status"]
    system["responses"].extend([Response(product(10, type="variable", variations=[101])), Response([leaf], total=1, pages=1)])
    data = page(system, parent_id=10).get_json()
    assert data["complete_page"] is False and data["code"] == "invalid_schema"


def test_explicit_leaf_flavor_takes_precedence_over_flavor_in_marketing_title(system):
    parent = product(10, name="Device - Promotional Cherry", type="variable", variations=[101])
    system["responses"].extend([
        Response(parent), Response([
            product(101, name="Device - Promotional Cherry Limited Edition", attributes=[{"name": "Flavor", "option": "Lemon"}]),
        ], total=1, pages=1),
    ])
    row = page(system, parent_id=10).get_json()["rows"][0]
    assert row["flavors"] == ["Lemon"]
    system["responses"].append(Response([
        product(1, name="Device - Promotional Cherry", attributes=[{"name": "Flavor", "options": ["Lemon"]}]),
    ], total=1, pages=1))
    row = page(system).get_json()["rows"][0]
    assert row["flavors"] == ["Lemon"]


def test_rules_are_loaded_once_per_page_and_shared_parser_extracts_brand_flavor(system):
    system["recognition"]["brands"] = [{"id": 1, "name": "IGET", "patterns": ["IGET"]}]
    system["responses"].append(Response([product(1, name="IGET ONE 12000 Puffs - Mixed Berries")], total=1, pages=1))
    row = page(system).get_json()["rows"][0]
    assert row["brands"] == ["IGET"] and row["flavors"] == ["Mixed Berries"]
    assert row["puffs"] == 12000
    assert len(system["recognition_calls"]) == 1


def configure_fumot_rules(system):
    system["recognition"]["brands"] = [{
        "id": 1, "name": "Fumot", "patterns": ["FUMOT", "RANDM", "R&M"],
        "aliases": ["R and M"],
    }]


def test_brand_only_returns_every_flavor_of_brand_without_name_search(system):
    configure_fumot_rules(system)
    system["responses"].append(Response([
        product(1, name="Generic Lemon", brands=[{"name": "Fumot"}], attributes=[{"name": "Flavor", "options": ["Lemon"]}]),
        product(2, name="Generic Cherry", brands=[{"name": "Fumot"}], attributes=[{"name": "Flavor", "options": ["Cherry"]}]),
        product(3, name="Other Lemon", brands=[{"name": "Other"}], attributes=[{"name": "Flavor", "options": ["Lemon"]}]),
    ], total=3, pages=1))
    data = page(system, brand="Fumot").get_json()
    assert [row["product_id"] for row in data["rows"]] == [1, 2]
    assert {flavor for row in data["rows"] for flavor in row["flavors"]} == {"Lemon", "Cherry"}
    assert data["source_ids"] == [1, 2, 3] and data["scanned"] == data["total"] == 3
    assert "brand" not in system["calls"][0][1]["params"]
    assert "search" not in system["calls"][0][1]["params"]


@pytest.mark.parametrize("brand", ["R&M", "randm", "R and M", "FÚMOT"])
def test_known_alias_and_canonical_brand_inputs_resolve_to_same_brand(system, brand):
    configure_fumot_rules(system)
    system["responses"].append(Response([product(1, name="Generic Device", brands=[{"name": "FUMOT"}])], total=1, pages=1))
    rows = page(system, brand=brand).get_json()["rows"]
    assert len(rows) == 1 and rows[0]["brands"] == ["Fumot"]


def test_recognized_alias_in_actual_brand_label_is_canonicalized(system):
    configure_fumot_rules(system)
    system["responses"].append(Response([
        product(1, name="Generic Device", attributes=[{"name": "pa_brand", "options": ["R&M"]}]),
    ], total=1, pages=1))
    row = page(system, brand="Fumot").get_json()["rows"][0]
    assert row["brands"] == ["Fumot"]


def test_brand_and_search_are_anded(system):
    configure_fumot_rules(system)
    system["responses"].append(Response([
        product(1, name="Generic Lemon", brands=[{"name": "Fumot"}]),
        product(2, name="Generic Cherry", brands=[{"name": "Fumot"}]),
        product(3, name="Other Lemon", brands=[{"name": "Other"}]),
    ], total=3, pages=1))
    rows = page(system, brand="R&M", search="lemon").get_json()["rows"]
    assert [row["product_id"] for row in rows] == [1]


def test_brand_word_in_flavor_or_marketing_suffix_is_not_brand_identity(system):
    system["recognition"]["brands"] = [{"id": 1, "name": "IGET", "patterns": ["IGET"]}]
    system["responses"].append(Response([
        product(1, name="Generic Device - IGET Lemon", brands=[{"name": "Other"}], attributes=[{"name": "Flavor", "options": ["IGET Lemon"]}]),
        product(2, name="Generic Device - IGET Lemon", attributes=[{"name": "Flavor", "options": ["IGET Lemon"]}]),
        product(3, name="IGETastic Device", attributes=[{"name": "Flavor", "options": ["Lemon"]}]),
        product(4, name="Generic Device IGET Lemon", attributes=[{"name": "Flavor", "options": ["IGET Lemon"]}]),
    ], total=4, pages=1))
    data = page(system, brand="IGET").get_json()
    assert data["complete_page"] is True and data["rows"] == []
    assert data["scanned"] == 4


def test_variation_flavor_only_name_cannot_supply_parent_brand(system):
    system["recognition"]["brands"] = [{"id": 1, "name": "IGET", "patterns": ["IGET"]}]
    parent = product(10, name="Generic Device", type="variable", variations=[101])
    system["responses"].extend([Response(parent), Response([product(101, name="IGET Lemon")], total=1, pages=1)])
    data = page(system, parent_id=10, brand="IGET").get_json()
    assert data["complete_page"] is True and data["rows"] == []


def test_unknown_parent_brand_still_returns_descriptor_and_leaf_brand_can_match(system):
    configure_fumot_rules(system)
    parent = product(10, name="Generic Device", type="variable", variations=[101, 102])
    system["responses"].append(Response([parent], total=1, pages=1))
    root_page = page(system, brand="Fumot").get_json()
    assert [descriptor["id"] for descriptor in root_page["variable_products"]] == [10]
    system["responses"].extend([Response(parent), Response([
        product(101, name="Lemon", brands=[{"name": "Fumot"}], attributes=[{"name": "Flavor", "option": "Lemon"}]),
        product(102, name="Cherry", brands=[{"name": "Fumot"}], attributes=[{"name": "Flavor", "option": "Cherry"}]),
    ], total=2, pages=1)])
    leaves = page(system, parent_id=10, brand="Fumot").get_json()["rows"]
    assert [row["variation_id"] for row in leaves] == [101, 102]


def test_free_brand_input_matches_only_actual_unknown_brand_label(system):
    system["responses"].append(Response([
        product(1, name="Generic Device", brands=[{"name": "Éxample Brands"}]),
        product(2, name="Example Device - Example Flavor", brands=[{"name": "Other"}]),
    ], total=2, pages=1))
    rows = page(system, brand="example").get_json()["rows"]
    assert [row["product_id"] for row in rows] == [1]


def test_exact_known_brand_does_not_match_other_brand_with_longer_name(system):
    configure_fumot_rules(system)
    system["responses"].append(Response([product(1, brands=[{"name": "Fumotastic"}])], total=1, pages=1))
    assert page(system, brand="Fumot").get_json()["rows"] == []


def test_shared_alias_is_not_assigned_to_one_arbitrary_canonical_brand(system):
    system["recognition"]["brands"] = [
        {"id": 1, "name": "First", "patterns": ["FIRST", "R&M"]},
        {"id": 2, "name": "Second", "patterns": ["SECOND", "R&M"]},
    ]
    system["responses"].append(Response([product(1, name="R&M Device")], total=1, pages=1))
    data = page(system, brand="First").get_json()
    assert data["complete_page"] is True and data["rows"] == []


def test_empty_brand_and_search_preserve_full_loading_behavior(system):
    items = [product(1, brands=[{"name": "First"}]), product(2, brands=[{"name": "Second"}])]
    system["responses"].append(Response(items, total=2, pages=1))
    data = page(system, brand="", search="").get_json()
    assert [row["product_id"] for row in data["rows"]] == [1, 2]


@pytest.mark.parametrize("params,expected", [
    ({"query_mode": "brand", "brand": "Fumot", "search": "Lemon"}, [1, 2]),
    ({"query_mode": "flavor", "brand": "Fumot", "search": "Lemon"}, [1, 3]),
    ({"query_mode": "all", "brand": "Fumot", "search": "Lemon"}, [1, 2, 3]),
])
def test_explicit_query_mode_ignores_stale_fields_from_other_modes(system, params, expected):
    configure_fumot_rules(system)
    system["responses"].append(Response([
        product(1, name="Lemon Device", brands=[{"name": "Fumot"}]),
        product(2, name="Cherry Device", brands=[{"name": "Fumot"}]),
        product(3, name="Lemon Device", brands=[{"name": "Other"}]),
    ], total=3, pages=1))
    rows = page(system, **params).get_json()["rows"]
    assert [row["product_id"] for row in rows] == expected


@pytest.mark.parametrize("params", [
    {"brand": "x" * 201}, {"brand": "!!!"}, {"query_mode": "invalid"},
    {"query_mode": "brand", "brand": ""}, {"query_mode": "brand", "brand": "  "},
])
def test_invalid_brand_request_is_rejected_before_any_network_read(system, params):
    response = page(system, **params)
    assert response.status_code == 400 and response.get_json()["code"] == "invalid_arguments"
    assert system["calls"] == []


@pytest.mark.parametrize("value,query", [("SMÁK", "smak"), ("Żółty", "zolty"), ("A&amp;B", "a&b"), ("Strawberry  Ice", "strawberry ice"), ("Blueberry-Ice", "blueberry ice"), ("Blueberry_Ice", "blueberry ice")])
def test_name_and_sku_normalization(system, value, query):
    system["responses"].append(Response([product(1, sku=value)], total=1, pages=1))
    assert len(page(system, search=query).get_json()["rows"]) == 1


def test_pagination_uses_counts_before_search_filtering(system):
    system["responses"].append(Response([product(i) for i in range(1, 51)], total=51, pages=2))
    data = page(system, search="nonexistent").get_json()
    assert data["rows"] == []
    assert data["total"] == 51 and data["total_pages"] == 2
    assert data["has_more"] is True and data["next_page"] == 2
    assert data["scanned"] == 50
    system["responses"].append(Response([product(51)], total=51, pages=2))
    data = page(system, page=2).get_json()
    assert data["has_more"] is False and data["next_page"] is None


def test_missing_headers_use_full_page_continue_and_short_page_end_with_warning(system):
    system["responses"].append(Response([product(i) for i in range(1, 51)]))
    data = page(system).get_json()
    assert data["pagination_mode"] == "length-fallback"
    assert data["total"] is None and data["total_pages"] is None
    assert data["has_more"] is True and data["next_page"] == 2
    assert len(data["warnings"]) == 3
    system["responses"].append(Response([product(51)]))
    data = page(system, page=2).get_json()
    assert data["complete_page"] is True and data["has_more"] is False
    assert data["warnings"]


@pytest.mark.parametrize("headers,mode", [
    ({"X-WP-Total": "100"}, "total-header"),
    ({"X-WP-TotalPages": "2"}, "pages-header"),
])
def test_single_pagination_header_knows_exact_full_page_terminal(system, headers, mode):
    system["responses"].append(Response([product(i) for i in range(1, 51)], headers=headers))
    first = page(system).get_json()
    assert first["has_more"] is True and first["next_page"] == 2
    system["responses"].append(Response([product(i) for i in range(51, 101)], headers=headers))
    last = page(system, page=2).get_json()
    assert last["complete_page"] is True
    assert last["has_more"] is False and last["next_page"] is None
    assert last["pagination_mode"] == mode
    assert last["warnings"]


def test_parent_name_flavor_does_not_match_sibling_but_model_query_still_matches_all(system):
    parent = product(10, name="IGET Blueberry", type="variable", variations=[101, 102], attributes=[
        {"name": "Flavour Profiles", "variation": True, "options": ["Blueberry", "Strawberry"]},
    ])
    leaves = [
        product(101, name="IGET Blueberry - Blueberry", attributes=[{"name": "Flavour Profiles", "option": "Blueberry"}]),
        product(102, name="IGET Blueberry - Strawberry", attributes=[{"name": "Flavour Profiles", "option": "Strawberry"}]),
    ]
    for query, expected in [("blueberry", [101]), ("IGET", [101, 102])]:
        system["responses"].extend([Response(parent), Response(leaves, total=2, pages=1)])
        data = page(system, parent_id=10, search=query).get_json()
        assert data["complete_page"] is True
        assert [row["variation_id"] for row in data["rows"]] == expected


def test_parent_rules_supply_brand_series_and_puffs_when_leaf_name_contains_only_flavor(system):
    system["recognition"]["brands"] = [{"id": 1, "name": "Merrymi", "patterns": ["MERRY MI"]}]
    system["recognition"]["series"] = [{"id": 2, "brand_id": 1, "name": "Blade"}]
    parent = product(10, name="Merry Mi Blade 30000 Puffs", type="variable", variations=[101])
    system["responses"].extend([Response(parent), Response([product(101, name="Aperol")], total=1, pages=1)])
    row = page(system, parent_id=10, search="Aperol").get_json()["rows"][0]
    assert row["brands"] == ["Merrymi"]
    assert row["series"] == "Blade"
    assert row["puffs"] == 30000


@pytest.mark.parametrize("parent_status", ["draft", "private", "pending"])
def test_parent_publication_status_constrains_variation_effective_status(system, parent_status):
    parent = product(10, type="variable", status=parent_status, variations=[101])
    system["responses"].extend([Response(parent), Response([product(101, status="publish")], total=1, pages=1)])
    row = page(system, parent_id=10).get_json()["rows"][0]
    assert row["status"] == row["parent_status"] == parent_status
    assert row["variation_status"] == "publish"


def test_variable_catalog_missing_identity_list_is_an_error(system):
    parent = product(10, type="variable")
    del parent["variations"]
    system["responses"].append(Response([parent], total=1, pages=1))
    data = page(system).get_json()
    assert data["code"] == "invalid_schema" and data["complete_page"] is False


@pytest.mark.parametrize("headers,items,code", [
    ({"X-WP-Total": "bad", "X-WP-TotalPages": "1"}, [product(1)], "invalid_pagination"),
    ({"X-WP-Total": "1", "X-WP-TotalPages": "2"}, [product(1)], "invalid_pagination"),
    ({"X-WP-Total": "51", "X-WP-TotalPages": "2"}, [product(1)], "incomplete_page"),
    ({"X-WP-Total": "1", "X-WP-TotalPages": "1"}, [], "incomplete_page"),
    ({"X-WP-Total": "51"}, [product(1)], "incomplete_page"),
])
def test_inconsistent_pagination_is_not_a_zero_result(system, headers, items, code):
    system["responses"].append(Response(items, headers=headers))
    response = page(system)
    assert response.status_code == 200
    assert response.get_json()["code"] == code
    assert response.get_json()["complete_page"] is False


def test_empty_catalog_is_confirmed_only_when_page_metadata_agrees(system):
    system["responses"].append(Response([], total=0, pages=0))
    data = page(system).get_json()
    assert data["complete_page"] is True and data["has_more"] is False
    assert data["rows"] == [] and data["total"] == 0


def test_repeated_ids_rejected_even_when_neither_matches_search(system):
    system["responses"].append(Response([product(1), product(1)], total=2, pages=1))
    data = page(system, search="Absent").get_json()
    assert data["code"] == "duplicate_id" and data["complete_page"] is False


@pytest.mark.parametrize("payload", [{"code": "rest_error"}, [None], [{"id": True, "type": "simple"}], [product(1, attributes="bad")], [product(1, attributes=[{"name": "Flavor", "options": "bad"}])]])
def test_schema_failure_is_not_a_zero_result(system, payload):
    system["responses"].append(Response(payload, total=1, pages=1))
    data = page(system).get_json()
    assert data["code"] == "invalid_schema" and data["complete_page"] is False


@pytest.mark.parametrize("response,code", [
    (requests.Timeout("credential=private_secret"), "connection_failed"),
    (Response(ValueError("<html>private_ck_1</html>")), "invalid_json"),
    (Response({"message": "credential private_secret"}, status=403), "upstream_http_error"),
])
def test_external_failures_remain_json_and_do_not_expose_response_or_exception_secrets(system, response, code):
    system["responses"].append(response)
    result = page(system)
    assert result.status_code == 200
    assert result.is_json
    assert result.get_json()["code"] == code
    assert result.get_json()["complete_page"] is False
    assert "private_" not in result.get_data(as_text=True)


def test_failed_variation_page_is_incomplete_despite_parent_success(system):
    system["responses"].extend([Response(product(10, type="variable")), requests.Timeout()])
    data = page(system, parent_id=10).get_json()
    assert data["code"] == "connection_failed" and data["complete_page"] is False
    assert data["rows"] == []


@pytest.mark.parametrize("parent", [product(11, type="variable"), product(10, type="simple")])
def test_invalid_parent_is_not_used_to_read_variations(system, parent):
    system["responses"].append(Response(parent))
    data = page(system, parent_id=10).get_json()
    assert data["code"] == "invalid_parent" and data["complete_page"] is False
    assert len(system["calls"]) == 1


def test_unconfigured_site_is_visible_but_not_readable(system):
    response = page(system, site_id=3)
    assert response.status_code == 200
    assert response.get_json()["code"] == "site_unconfigured"
    assert response.get_json()["complete_page"] is False
    assert system["calls"] == []


def test_own_scope_site_list_and_page_match_live_named_manager(system):
    system["user"].own_scoped = True
    sites = system["client"].get("/api/product-manager/catalog-sites").get_json()["sites"]
    assert {site["id"] for site in sites} == {1, 3}
    forbidden = page(system, site_id=2)
    assert forbidden.status_code == 403 and forbidden.get_json()["code"] == "site_forbidden"
    assert system["calls"] == []
    conn = system["connect"]()
    conn.execute("UPDATE sites SET manager = 'Jane' WHERE id = 1")
    conn.commit()
    conn.close()
    assert page(system).status_code == 403


def test_empty_owner_cannot_read_unassigned_site(system):
    system["user"].own_scoped = True
    system["user"].name = ""
    assert system["client"].get("/api/product-manager/catalog-sites").get_json()["sites"] == []
    assert page(system, site_id=4).status_code == 403


def test_product_permission_is_required_for_both_routes(system):
    system["user"].may_manage = False
    assert page(system).status_code == 403
    assert system["client"].get("/api/product-manager/catalog-sites").status_code == 403
    assert system["calls"] == []


def test_unauthenticated_apis_return_json_instead_of_login_html(system):
    with system["client"].session_transaction() as session:
        session.clear()
    for path in ("/api/product-manager/catalog-sites", "/api/product-manager/catalog-page?site_id=1"):
        response = system["client"].get(path)
        assert response.status_code == 403 and response.is_json
        assert response.get_json()["code"] == "authentication_required"
    assert system["calls"] == []


@pytest.mark.parametrize("params", [{"site_id": "abc"}, {"site_id": 0}, {"page": 0}, {"parent_id": -1}, {"search": "x" * 201}])
def test_invalid_request_arguments_return_json_400(system, params):
    response = page(system, **params)
    assert response.status_code == 400
    assert response.get_json()["code"] == "invalid_arguments"
    assert system["calls"] == []


def test_nonexistent_site_returns_json_404(system):
    response = page(system, site_id=999)
    assert response.status_code == 404 and response.get_json()["code"] == "site_not_found"
    assert system["calls"] == []


def test_unknown_database_error_is_json_and_does_not_leak_internal_details(system):
    conn = system["connect"]()
    conn.execute("DROP TABLE sites")
    conn.commit()
    conn.close()
    for response in (page(system), system["client"].get("/api/product-manager/catalog-sites")):
        assert response.status_code == 200 and response.is_json
        assert response.get_json()["code"] == "catalog_read_failed"
        assert "sites" not in response.get_data(as_text=True)
