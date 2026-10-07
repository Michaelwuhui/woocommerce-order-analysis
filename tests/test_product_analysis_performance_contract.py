"""Synthetic release contracts for the product analysis performance change.

Only selected app function ASTs execute. There is no app startup, configured DB,
real order/customer data or network. These golden values were independently
checked against main 4c3f2ff before the performance change. The product page has
sales revenue, not cost/margin fields; no profit is inferred from net revenue.
Some old chart/filter differences are intentional compatibility assertions.
"""
import ast
from datetime import date, datetime, timedelta
import html
import json
from pathlib import Path
import re
import socket
import sqlite3
from types import SimpleNamespace

from flask import Flask, jsonify, request
from jinja2 import ChoiceLoader, DictLoader, FileSystemLoader
import pytest
import requests

import db_backend as db
from product_analysis_drilldown import finalize_product_drilldown, record_product_drilldown


ROOT = Path(__file__).resolve().parents[1]
PL = "https://www.historical.example.invalid"
CZ = "https://czech.example.invalid"
AE = "https://emirates.example.invalid"
MINT = "IGET Alpha 12000 Puffs - Mint"
MAPPED = "Mapped &#8222;device&#8220;"


class FrozenDate(date):
    @classmethod
    def today(cls):
        return cls(2026, 10, 7)


class FrozenDatetime(datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 10, 7, 12, 0, 0, tzinfo=tz)


class IsolatedClock(ast.NodeTransformer):
    def visit_ImportFrom(self, node):
        if node.module != "datetime":
            return node
        # Freeze only these function-local imports. No global datetime patch
        # can affect Flask or another concurrently running test.
        return [ast.copy_location(ast.Assign(
            targets=[ast.Name(id=alias.asname or alias.name, ctx=ast.Store())],
            value=ast.Name(id="_clock_" + alias.name, ctx=ast.Load())), node)
            for alias in node.names]


class ReadOnlyConnection:
    def __init__(self, raw):
        self.raw = raw
        self.statements = []

    def execute(self, sql, params=()):
        assert sql.lstrip().lower().startswith("select"), "Analysis may only read fixture data"
        self.statements.append((sql, tuple(params)))
        return self.raw.execute(sql, params)

    def commit(self):
        self.raw.commit()

    def close(self):
        pass


def _compile(scope, source):
    names = {
        "products", "get_unknown_products", "get_product_stats", "get_all_managers",
        "parse_json_field", "extract_flavor_from_meta", "extract_puffs_from_meta",
        "get_full_product_name", "normalize_flavor", "normalize_raw_name",
        "parse_product_name", "get_user_allowed_sources", "_resolve_user_sites",
        "_on_hold_is_shipped_clause", "_revenue_status_cond", "_active_status_cond",
    }
    functions = []
    for node in ast.parse(source).body:
        if isinstance(node, ast.FunctionDef) and node.name in names:
            node.decorator_list = []
            functions.append(IsolatedClock().visit(node))
    assert names == {node.name for node in functions}
    tree = ast.fix_missing_locations(ast.Module(body=functions, type_ignores=[]))
    exec(compile(tree, "app.py (isolated contract)", "exec"), scope)


def _item(name=MINT, qty=1, total=10, flavor="Mint", price=None, puffs=None):
    value = {"name": name, "quantity": qty, "total": str(total), "meta_data": []}
    if price is not None:
        value["price"] = str(price)
    if flavor is not None:
        value["meta_data"].append({"key": "pa_smak", "value": flavor})
    if puffs is not None:
        value["meta_data"].append({"key": "pa_puffs", "value": str(puffs)})
    return value


def _order(raw, ident, source=PL, currency="PLN", created="2026-10-02T10:00:00",
           items=None, total=None, shipping=0, status="completed", payment="cod",
           undelivered=0, problem=0, loss=0, modified="2026-10-07T15:00:00"):
    if items is None:
        items = [_item()]
    if total is None:
        total = sum(float(item.get("total", 0)) for item in items) + shipping
    raw.execute("INSERT INTO orders VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
        ident, "fixture-" + ident, source, currency, created, modified,
        json.dumps(items), total, shipping, status, payment, undelivered, problem, loss))


def _build_contract(monkeypatch, source=None):
    def forbidden(*args, **kwargs):
        pytest.fail("Configured DB and network are forbidden in this contract")

    monkeypatch.setattr(db, "connect", forbidden)
    monkeypatch.setattr(requests.sessions.Session, "request", forbidden)
    monkeypatch.setattr(socket, "create_connection", forbidden)
    raw = sqlite3.connect(":memory:")
    raw.row_factory = sqlite3.Row
    raw.executescript("""
        CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT,manager TEXT,country TEXT,
            cod_on_hold_is_shipped INTEGER,api_status TEXT);
        INSERT INTO sites VALUES(7,'https://www.historical.example.invalid','Owner A','PL',1,'archived');
        INSERT INTO sites VALUES(37,'https://czech.example.invalid','Owner B','CZ',0,'ok');
        INSERT INTO sites VALUES(20,'https://emirates.example.invalid','Owner B','AE',0,'ok');
        INSERT INTO sites VALUES(9,'https://excluded.example.invalid','Owner B','PL',1,'ok');
        CREATE TABLE user_country_permissions(user_id INTEGER,country TEXT);
        CREATE TABLE user_site_permissions(user_id INTEGER,site_id INTEGER);
        CREATE TABLE user_site_exclusions(user_id INTEGER,site_id INTEGER);
        INSERT INTO user_site_permissions VALUES(2,7);
        INSERT INTO user_country_permissions VALUES(3,'PL');
        INSERT INTO user_site_permissions VALUES(3,37);
        INSERT INTO user_site_exclusions VALUES(3,9);
        CREATE TABLE brands(id INTEGER PRIMARY KEY,name TEXT,aliases TEXT);
        INSERT INTO brands VALUES(1,'IGET','["IG"]');
        INSERT INTO brands VALUES(2,'Merrymi','["MERRY MI"]');
        CREATE TABLE series(id INTEGER PRIMARY KEY,brand_id INTEGER,name TEXT);
        INSERT INTO series VALUES(1,1,'Alpha'),(2,1,'Beta'),(3,2,'Blade');
        CREATE TABLE product_mappings(id INTEGER PRIMARY KEY,raw_name TEXT,source TEXT,
            brand_id INTEGER,puff_count INTEGER,flavor TEXT,series_id INTEGER,is_manual INTEGER);
        INSERT INTO product_mappings VALUES(1,'Mapped „device“ - Orange','',2,30000,'LOVE_66',3,1);
        INSERT INTO product_mappings VALUES(2,'BaseMapped','',1,15000,'Raw_Flavor',1,1);
        INSERT INTO product_mappings VALUES(3,'BaseMapped - Cherry','',2,25000,'CherryMapped',3,1);
        CREATE TABLE orders(id TEXT PRIMARY KEY,number TEXT,source TEXT,currency TEXT,
            date_created TEXT,date_modified TEXT,line_items TEXT,total REAL,shipping_total REAL,
            status TEXT,payment_method TEXT,is_undelivered INTEGER,is_problem_return INTEGER,
            shipping_loss_amount REAL);
    """)
    _order(raw, "7-101", created="2026-10-01T10:00:00", status="on-hold", shipping=20, total=180, items=[
        _item(qty=2, total=100, price=50), _item(qty=1, total=50, price=45),
        _item("Merrymi Blade 30000 Puffs - Kiwi", total=50, flavor="Kiwi", price=50)])
    _order(raw, "37-101", source=CZ, currency="CZK", created="2026-10-04T10:00:00",
        shipping=20, items=[_item(qty=2, total=200, price=100)])
    _order(raw, "7-103", created="2026-10-06T10:00:00", items=[
        _item("IG Beta 12000 Puffs - Mint", qty=3, total=90, price=30),
        _item(MAPPED, qty=2, total=20, flavor="Orange", price=10),
        _item("BaseMapped", total=10, flavor="Cherry", price=10),
        _item("BaseMapped", total=10, flavor="Lime", price=10),
        _item("Mystery device", total=5, flavor=None, price=5), _item(total=0, price=0)])
    _order(raw, "20-101", source=AE, currency="AED", created="2026-10-03T10:00:00",
        shipping=2, items=[_item("Mystery device", qty=2, total=8, flavor=None)])
    for ident, created, qty in [
        ("7-old", "2026-08-10T10:00:00", 5), ("7-eight", "2026-08-13T10:00:00", 6),
        ("7-sept", "2026-09-01T10:00:00", 7), ("7-future", "2026-10-08T10:00:00", 11)]:
        _order(raw, ident, created=created, items=[_item(qty=qty, total=qty * 10)])
    for ident, site, cur, loss, created in [
        ("loss-pl", PL, "PLN", 12.5, "2026-10-03T10:00:00"),
        ("loss-cz", CZ, "CZK", 3, "2026-10-03T10:00:00"),
        ("loss-only", PL, "USD", 9, "2026-10-03T10:00:00"),
        ("loss-old", PL, "PLN", 11, "2026-08-13T10:00:00"),
        ("loss-future", PL, "PLN", 20, "2026-10-08T10:00:00")]:
        _order(raw, ident, source=site, currency=cur, created=created,
            status="failed", undelivered=1, loss=loss)
    raw.execute("INSERT INTO orders VALUES('invalid','invalid',?,?,?,?,?,?,?,?,?,?,?,?)", (
        PL, "PLN", "2026-10-01", "2026-10-07", '{"not":"a list"}',
        50, 0, "completed", "cod", 0, 0, 0))
    raw.commit()
    conn = ReadOnlyConnection(raw)
    clock_calls = []

    def rate(currency, month):
        clock_calls.append((currency, month))
        return {"PLN": 2, "CZK": .3, "USD": 7}.get(currency), month

    user = SimpleNamespace(id=1, admin=True, viewer=False)
    user.is_admin = lambda: user.admin
    user.is_viewer = lambda: user.viewer
    scope = {
        "json": json, "html": html, "re": re, "request": request, "jsonify": jsonify,
        "current_user": user, "get_db_connection": lambda: conn, "get_cny_rate": rate,
        "record_product_drilldown": record_product_drilldown,
        "finalize_product_drilldown": finalize_product_drilldown,
        "render_template": lambda name, **values: values,
        "_clock_date": FrozenDate, "_clock_datetime": FrozenDatetime, "_clock_timedelta": timedelta,
    }
    _compile(scope, source or (ROOT / "app.py").read_text(encoding="utf-8"))
    app = Flask("product-analysis-contract")
    app.add_url_rule("/products", endpoint="products", view_func=scope["products"])

    def run(args="", endpoint="products"):
        with app.test_request_context("/products" + ("?" + args if args else "")):
            result = scope[endpoint]()
            return result.get_json() if hasattr(result, "get_json") else result

    return SimpleNamespace(raw=raw, conn=conn, scope=scope, user=user,
        run=run, app=app, rate_calls=clock_calls)


@pytest.fixture
def contract(monkeypatch):
    value = _build_contract(monkeypatch)
    yield value
    value.raw.close()


def _product(value, brand="IGET", series="Alpha", puffs=12000, flavor="MINT"):
    return next(row for row in value["top_products"] if
        (row["brand"], row["series"], row["puffs"], row["flavor"]) == (brand, series, puffs, flavor))


def test_multicurrency_discount_shipping_loss_and_missing_rate_golden(contract):
    value = contract.run()
    assert value["totals"] == {
        "total_quantity": 17,
        "total_revenue_by_currency": {"PLN": 282.5, "CZK": 197.0, "AED": 8.0},
        "total_gross_revenue_by_currency": {"PLN": 315.0, "CZK": 220.0, "AED": 10.0},
        "total_revenue_cny": 624.1, "total_gross_revenue_cny": 696.0,
        "shipping_loss_by_currency": {"PLN": 12.5, "CZK": 3.0, "USD": 9.0},
        "brand_count": 3, "product_count": 7,
    }
    mint = _product(value)
    assert mint["quantity"] == 6
    assert mint["order_count"] == 4  # Four matching line items, three distinct orders.
    assert mint["revenue_by_currency"] == {"PLN": 120.0, "CZK": 200.0}
    assert mint["gross_revenue_by_currency"] == {"PLN": 135.0, "CZK": 220.0}
    assert {month for _, month in contract.rate_calls} == {"2026-10"}
    assert any(currency == "AED" for currency, _ in contract.rate_calls)


def test_manual_full_name_raw_name_entities_series_and_unknown_remain_distinct(contract):
    value = contract.run()
    assert [(row["brand"], row["series"], row["puffs"], row["flavor"], row["quantity"])
        for row in value["top_products"]] == [
        ("IGET", "Alpha", 12000, "MINT", 6), ("IGET", "Beta", 12000, "MINT", 3),
        ("Unknown", "", None, "", 3), ("Merrymi", "Blade", 30000, "LOVE 66", 2),
        ("Merrymi", "Blade", 30000, "KIWI", 1),
        ("Merrymi", "Blade", 25000, "CHERRYMAPPED", 1),
        ("IGET", "Alpha", 15000, "RAW FLAVOR", 1),
    ]
    assert value["puff_options"] == ["12000", "15000", "25000", "30000"]
    assert [(row["name"], row["quantity"]) for row in value["brand_ranking"]] == [
        ("IGET", 10), ("Merrymi", 4), ("Unknown", 3)]
    assert [(row["puffs"], row["quantity"]) for row in value["puff_ranking"]] == [
        ("12000", 9), ("30000", 3), ("Unknown", 3), ("25000", 1), ("15000", 1)]


def test_preloaded_drilldown_unit_price_distinct_orders_and_archived_owner(contract):
    mint = _product(contract.run())
    assert mint["source_prices"] == [
        {"source": CZ, "site": "czech.example.invalid", "manager": "Owner B", "currency": "CZK",
         "latest_price": "100.00", "min_price": "100.00", "max_price": "100.00",
         "latest_date": "2026-10-04", "order_count": 1, "quantity": 2},
        {"source": PL, "site": "historical.example.invalid", "manager": "Owner A", "currency": "PLN",
         "latest_price": "0.00", "min_price": "0.00", "max_price": "50.00",
         "latest_date": "2026-10-06", "order_count": 2, "quantity": 4},
    ]
    assert mint["recent_orders"] == [
        {"order_number": "fixture-7-103", "source": "historical.example.invalid", "manager": "Owner A", "date": "2026-10-06"},
        {"order_number": "fixture-37-101", "source": "czech.example.invalid", "manager": "Owner B", "date": "2026-10-04"},
        {"order_number": "fixture-7-101", "source": "historical.example.invalid", "manager": "Owner A", "date": "2026-10-01"},
    ]
    assert not any(key.startswith("_") for key in mint)


def test_brand_puff_topn_filter_compatibility_and_loss_window(contract):
    brand = contract.run("brand=IGET")
    assert brand["totals"]["total_quantity"] == 10
    assert brand["totals"]["total_revenue_by_currency"] == {"PLN": 207.5, "CZK": 197.0}
    assert brand["totals"]["shipping_loss_by_currency"] == {"PLN": 12.5, "CZK": 3.0, "USD": 9.0}
    puff = contract.run("puffs=12000")
    assert puff["totals"]["total_quantity"] == 12  # Unknown puffs remain visible.
    assert puff["totals"]["total_revenue_by_currency"] == {"PLN": 202.5, "CZK": 197.0, "AED": 8.0}
    no_known = contract.run("puffs=999")
    assert no_known["totals"]["total_revenue_by_currency"] == {"PLN": -7.5, "AED": 8.0}
    assert contract.run("top_n=1")["totals"] == contract.run()["totals"]
    assert len(contract.run("top_n=1")["top_products"]) == 1


def test_weekly_chart_golden_keeps_separate_classifier_and_series_key(contract):
    value = contract.run()
    assert value["weekly_trend"] == {
        "weeks": ["08/10", "08/31", "09/28", "10/05"],
        "flavors": ["MINT", "LOVE 66", "KIWI", "CHERRYMAPPED", "RAW FLAVOR"],
        "datasets": [
            {"flavor": "MINT", "data": [6, 7, 5, 15], "pageTotal": 9},
            {"flavor": "LOVE 66", "data": [0, 0, 0, 0], "pageTotal": 2},
            {"flavor": "KIWI", "data": [0, 0, 1, 0], "pageTotal": 1},
            {"flavor": "CHERRYMAPPED", "data": [0, 0, 0, 1], "pageTotal": 1},
            {"flavor": "RAW FLAVOR", "data": [0, 0, 0, 0], "pageTotal": 1},
        ],
    }
    assert value["product_trend"]["datasets"] == [
        {"label": "IGET 12000 MINT", "data": [6, 7, 5, 15], "pageTotal": 3},
        {"label": "Unknown N/A ", "data": [0, 0, 2, 1], "pageTotal": 3},
        {"label": "Merrymi 30000 LOVE 66", "data": [0, 0, 0, 2], "pageTotal": 2},
        {"label": "Merrymi 30000 KIWI", "data": [0, 0, 1, 0], "pageTotal": 1},
        {"label": "Merrymi 25000 CHERRYMAPPED", "data": [0, 0, 0, 1], "pageTotal": 1},
        {"label": "IGET 15000 RAW FLAVOR", "data": [0, 0, 0, 1], "pageTotal": 1},
    ]


@pytest.mark.parametrize("args,quantity,net", [
    ("country=CZ", 2, {"CZK": 197.0}),
    ("manager=Owner+A", 13, {"PLN": 282.5}),
    ("source=" + PL, 13, {"PLN": 282.5}),
    ("manager=Absent", 0, {}), ("country=Absent", 0, {}),
    ("quick_date=last_month", 7, {"PLN": 70.0}),
    ("quick_date=all", 46, {"PLN": 541.5, "CZK": 197.0, "AED": 8.0}),
])
def test_filter_and_created_date_golden(contract, args, quantity, net):
    value = contract.run(args)
    assert value["totals"]["total_quantity"] == quantity
    assert value["totals"]["total_revenue_by_currency"] == net
    if args == "country=CZ":
        # Preserve historical trend range; country only scopes the page ranking.
        assert value["weekly_trend"]["datasets"][0] == {
            "flavor": "MINT", "data": [6, 7, 5, 15], "pageTotal": 2}


@pytest.mark.parametrize("user_id,admin,viewer,quantity,expected_sources", [
    (1, True, False, 17, {PL, CZ, AE}), (2, False, False, 13, {PL}),
    (3, False, False, 15, {PL, CZ}), (4, False, False, 0, set()),
    (4, False, True, 17, {PL, CZ, AE}),
])
def test_live_permission_scope_includes_archived_history(contract, user_id, admin, viewer, quantity, expected_sources):
    contract.user.id, contract.user.admin, contract.user.viewer = user_id, admin, viewer
    value = contract.run("source=" + AE if user_id in (2, 3) else "")
    assert value["totals"]["total_quantity"] == quantity
    assert {row["source"] for row in value["sources"]} == expected_sources
    observed = {row["source"] for p in value["top_products"] for row in p["source_prices"]}
    assert observed <= expected_sources
    if user_id in (2, 3):
        assert value["current_filters"]["source"] == ""
        assert PL in observed


@pytest.mark.parametrize("status,payment,site,undelivered,problem,included", [
    ("on-hold", "cod", PL, 0, 0, True), ("on-hold", "cod", AE, 0, 0, False),
    ("on-hold", "bacs", PL, 0, 0, False), ("pending", "stripe", PL, 0, 0, False),
    ("pending", "cod", PL, 0, 0, True), ("processing", "stripe", CZ, 0, 0, True),
    ("completed", "cod", PL, 1, 0, False), ("completed", "cod", PL, 0, 1, False),
    ("refunded", "stripe", PL, 0, 0, False), ("checkout-draft", "cod", PL, 0, 0, False),
])
def test_revenue_status_country_cod_rules_do_not_change(contract, status, payment, site, undelivered, problem, included):
    _order(contract.raw, "status-fixture", source=site, status=status, payment=payment,
        undelivered=undelivered, problem=problem, items=[_item(qty=123, total=1230)])
    assert contract.run()["totals"]["total_quantity"] == 17 + (123 if included else 0)


def test_template_quantity_share_and_drilldown_have_existing_values(contract):
    value = contract.run()
    # Replace only the unrelated application shell; render the real product view.
    contract.app.jinja_loader = ChoiceLoader([
        DictLoader({"base.html": "{% block content %}{% endblock %}{% block scripts %}{% endblock %}"}),
        FileSystemLoader(str(ROOT / "templates")),
    ])
    with contract.app.test_request_context("/products"):
        rendered = contract.app.jinja_env.get_template("products.html").render(**value)
    assert "52.9%" in rendered  # 12000 puffs: 9 / 17, not revenue share.
    assert "17.6%" in rendered
    assert "624.10" in rendered and "696.00" in rendered
    assert "historical.example.invalid" in rendered
    assert "Owner A" in rendered
    assert "source_prices" in rendered and "recent_orders" in rendered
    assert "/api/products/samples?" not in rendered


def test_analysis_only_reads_and_reloads_mapping_rules_next_request(contract):
    before = contract.raw.iterdump()
    before = list(before)
    first = contract.run()
    assert list(contract.raw.iterdump()) == before
    contract.raw.execute("UPDATE product_mappings SET flavor='Changed' WHERE id=3")
    second = contract.run()
    assert _product(first, "Merrymi", "Blade", 25000, "CHERRYMAPPED")["quantity"] == 1
    assert _product(second, "Merrymi", "Blade", 25000, "CHANGED")["quantity"] == 1
    assert all(sql.lstrip().lower().startswith("select") for sql, _ in contract.conn.statements)


@pytest.mark.parametrize("user_id,admin,viewer,quantity,sources", [
    (1, True, False, 3, [CZ, PL, AE]),
    (2, False, False, 1, [PL]), (3, False, False, 1, [PL]),
    (4, False, False, 0, []), (4, False, True, 3, [CZ, PL, AE]),
])
def test_unknown_endpoint_applies_live_permission_without_losing_manual_mappings(contract, user_id, admin, viewer, quantity, sources):
    contract.user.id, contract.user.admin, contract.user.viewer = user_id, admin, viewer
    unknown = contract.run(endpoint="get_unknown_products")
    # There are no Mystery orders on CZ; sources lists report observed data only.
    expected_sources = sorted(source for source in sources if source != CZ)
    assert unknown == ([{
        "name": "Mystery device", "sample_full_name": "Mystery device", "puffs": None,
        "quantity": quantity, "sources": expected_sources,
    }] if quantity else [])


def test_unknown_mapping_source_scope_remains_authoritative(contract):
    contract.raw.execute("INSERT INTO product_mappings VALUES(4,'Generic mapped','https://www.historical.example.invalid',1,9000,'',NULL,1)")
    _order(contract.raw, "mapping-pl", items=[_item("Generic mapped", qty=3, flavor="Lemon")])
    _order(contract.raw, "mapping-cz", source=CZ, currency="CZK",
        items=[_item("Generic mapped", qty=4, flavor="Lemon", puffs=10000)])
    rows = contract.run(endpoint="get_unknown_products")
    generic = next(row for row in rows if row["name"] == "Generic mapped")
    assert generic == {"name": "Generic mapped", "sample_full_name": "Generic mapped - Lemon",
        "puffs": 10000, "quantity": 4, "sources": [CZ]}


def test_parser_work_is_bounded_by_unique_names_and_series_rules_loaded_once(contract):
    _order(contract.raw, "many-identical", items=[_item(qty=1, total=1) for _ in range(1000)])
    calls = []
    original = contract.scope["parse_product_name"]

    def counted(name, brands=None, series=None):
        assert series is not None, "No parser may reload series inside a line loop"
        calls.append(name)
        return original(name, brands, series)

    contract.scope["parse_product_name"] = counted
    contract.run()
    assert len(calls) == len(set(calls)) == 4
    assert sum("from series" in sql.lower() for sql, _ in contract.conn.statements) == 1
    assert len(contract.rate_calls) == len(set(contract.rate_calls)) == 3


def test_cached_name_recognition_keeps_each_variation_attribute(contract):
    _order(contract.raw, "different-attributes", items=[
        _item(qty=2, total=10, flavor="Fresh_Berry", puffs=16000),
        _item(qty=3, total=15, flavor="Orange", puffs=17000),
    ])
    value = contract.run()
    assert _product(value, puffs=16000, flavor="FRESH BERRY")["quantity"] == 2
    assert _product(value, puffs=17000, flavor="ORANGE")["quantity"] == 3
    assert _product(value)["quantity"] == 6
