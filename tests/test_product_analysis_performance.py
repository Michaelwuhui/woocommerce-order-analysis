"""Semantic and bounded-work checks without importing app startup or live DBs."""
import ast
from datetime import date, timedelta
import html
import json
from pathlib import Path
import re
import sqlite3
from types import SimpleNamespace

import pytest

from product_analysis_data import RequestProductParser, build_weekly_trends
from product_analysis_drilldown import finalize_product_drilldown, record_product_drilldown
from product_recognition import parse_product_name


SOURCE = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)


def _functions(names, scope):
    nodes = [node for node in TREE.body if isinstance(node, ast.FunctionDef) and node.name in names]
    for node in nodes:
        node.decorator_list = []
    exec(compile(ast.Module(body=nodes, type_ignores=[]), "isolated-products", "exec"), scope)
    return scope


def _helpers():
    return _functions({"get_full_product_name", "extract_flavor_from_meta", "extract_puffs_from_meta",
                       "normalize_raw_name", "normalize_flavor", "parse_json_field"},
                      {"json": json, "html": html, "re": re})


def _item(name="IGET One 12000 Puffs", quantity=1, flavor="Blueberry Ice", **overrides):
    return {"name": name, "quantity": quantity, "total": "20", "price": "20",
            "meta_data": [{"key": "pa_flavour", "value": flavor}], **overrides}


def test_name_cache_is_request_local_and_does_not_merge_variation_metadata():
    brands = [{"id": 1, "name": "IGET", "patterns": ["IGET"]}]
    calls = []

    def parser(name, rules, series):
        calls.append((name, rules, series))
        return parse_product_name(name, rules, series)

    cached = RequestProductParser(parser, brands, [])
    for _ in range(500):
        assert cached("IGET One 12000 Puffs")["brand"] == "IGET"
    assert len(calls) == 1
    changed = RequestProductParser(parser, [{"id": 2, "name": "Other", "patterns": ["IGET"]}], [])
    assert changed("IGET One 12000 Puffs")["brand"] == "Other"
    assert len(calls) == 2
    # The cached object contains name recognition only; attributes stay per leaf.
    helpers = _helpers()
    assert helpers["get_full_product_name"](_item(flavor="Lemon"))[1] == "Lemon"
    assert helpers["get_full_product_name"](_item(flavor="Apple"))[1] == "Apple"


def test_one_trend_pass_preserves_distinct_flavor_and_product_mapping_rules():
    h = _helpers()
    decode_calls, full_calls = [], []
    brands = [{"id": 1, "name": "IGET", "patterns": ["IGET"]}]
    parser = RequestProductParser(parse_product_name, brands, [])
    meta_item = _item(quantity=2)
    base_item = _item(name="Generic", quantity=3, flavor="Original")
    full_item = _item(name="FullMapped", quantity=4, flavor="Old")
    parsed_only = _item(name="IGET One 12000 Puffs - Blueberry Ice", quantity=5, meta_data=[])
    orders = [{"date_created": "2026-10-06T10:00:00", "line_items": json.dumps(
        [meta_item, base_item, full_item, parsed_only])},
        {"date_created": "2026-10-13T10:00:00", "line_items": "[]"}]
    top = [{"brand": "IGET", "puffs": 12000, "flavor": "BLUEBERRY ICE", "quantity": 14}]
    mappings = {
        "Generic": {"brand": "IGET", "puffs": 12000, "flavor": "Blueberry Ice"},
        "FullMapped - Old": {"brand": "IGET", "puffs": 12000, "flavor": "Blueberry Ice"},
    }

    def decode(value):
        decode_calls.append(value)
        return json.loads(value)

    def full(item):
        full_calls.append(item)
        return h["get_full_product_name"](item)

    weekly, products = build_weekly_trends(
        orders, top, ["BLUEBERRY ICE"], {"BLUEBERRY ICE": 14}, mappings,
        parse_items=decode, full_product_name=full, normalize_raw_name=h["normalize_raw_name"],
        normalize_flavor=h["normalize_flavor"], parse_product=parser)
    # Flavor chart ignores base-name mappings and parsed-only flavors, as before.
    assert weekly["datasets"] == [{"flavor": "BLUEBERRY ICE", "data": [6, 0], "pageTotal": 14}]
    assert products["datasets"] == [{"label": "IGET 12000 BLUEBERRY ICE", "data": [14, 0], "pageTotal": 14}]
    assert weekly["weeks"] == products["weeks"] == ["10/05", "10/12"]
    assert len(decode_calls) == len(orders)
    assert len(full_calls) == 4


def test_trend_malformed_line_stops_each_former_pass_independently():
    h = _helpers()
    # A non-string name breaks only the parser in the former product pass. The
    # flavor pass reads valid metadata, so it still counts later valid lines.
    mappings = {}
    top = [{"brand": "IGET", "puffs": 12000, "flavor": "LEMON", "quantity": 3}]
    orders = [{"date_created": "invalid", "line_items": "[]"},
              {"date_created": "2026-10-06", "line_items": json.dumps([
                  _item(name=123, flavor="Lemon"), _item(quantity=2, flavor="Lemon")])}]
    weekly, products = build_weekly_trends(
        orders, top, ["LEMON"], {"LEMON": 3}, mappings, parse_items=json.loads,
        full_product_name=h["get_full_product_name"], normalize_raw_name=h["normalize_raw_name"],
        normalize_flavor=h["normalize_flavor"],
        parse_product=RequestProductParser(parse_product_name, [{"id": 1, "name": "IGET", "patterns": ["IGET"]}], []))
    assert weekly["datasets"][0]["data"] == [3]
    assert products["datasets"][0]["data"] == [0]


class Connection:
    def __init__(self, raw):
        self.raw = raw
        self.queries = []

    def execute(self, statement, params=()):
        self.queries.append((statement, params))
        return self.raw.execute(statement, params)

    def commit(self):
        self.raw.commit()

    def close(self):
        pass


@pytest.fixture
def endpoint():
    raw = sqlite3.connect(":memory:")
    raw.row_factory = sqlite3.Row
    raw.executescript("""
        CREATE TABLE sites(id INTEGER PRIMARY KEY,url TEXT,manager TEXT,country TEXT);
        INSERT INTO sites VALUES(1,'https://one.invalid','A','PL');
        INSERT INTO sites VALUES(2,'https://two.invalid','B','CZ');
        CREATE TABLE orders(id TEXT,number TEXT,line_items TEXT,source TEXT,currency TEXT,total REAL,
            shipping_total REAL,date_created TEXT,status TEXT,is_undelivered INTEGER,shipping_loss_amount REAL);
        CREATE TABLE brands(id INTEGER,name TEXT,aliases TEXT);
        INSERT INTO brands VALUES(1,'IGET','[]');
        CREATE TABLE series(id INTEGER,brand_id INTEGER,name TEXT);
        INSERT INTO series VALUES(1,1,'One');
        CREATE TABLE product_mappings(raw_name TEXT,source TEXT,puff_count INTEGER,flavor TEXT,
            series_id INTEGER,brand_id INTEGER,is_manual INTEGER);
    """)
    now = date.today().isoformat() + "T10:00:00"
    items = json.dumps([_item(quantity=2), _item(quantity=1, flavor="Lemon")])
    raw.execute("INSERT INTO orders VALUES('1-101','101',?,'https://one.invalid','PLN',55,10,?,'completed',0,0)", (items, now))
    raw.execute("INSERT INTO orders VALUES('2-102','102',?,'https://two.invalid','CZK',55,10,?,'completed',0,0)", (items, now))
    raw.execute("INSERT INTO orders VALUES('1-103','103','[]','https://one.invalid','PLN',0,0,?,'completed',1,7)", (now,))
    raw.commit()
    c = Connection(raw)
    scope = _helpers()
    scope.update({"get_db_connection": lambda: c,
                  "current_user": SimpleNamespace(id=1, is_admin=lambda: True, is_viewer=lambda: False),
                  "get_user_allowed_sources": lambda *_: None,
                  "get_all_managers": lambda: ["A", "B"],
                  "request": SimpleNamespace(args={"quick_date": "all"}),
                  "_revenue_status_cond": lambda: "status='completed' AND is_undelivered=0",
                  "_active_status_cond": lambda: "status='completed'",
                  "get_cny_rate": lambda currency, month: ({"PLN": 2, "CZK": 0.3}.get(currency), month),
                  "record_product_drilldown": record_product_drilldown,
                  "finalize_product_drilldown": finalize_product_drilldown,
                  "render_template": lambda name, **context: context,
                  "jsonify": lambda value: value})
    parse_calls = []

    def parse(name, brands=None, series=None):
        assert brands is not None and series is not None, "No per-line fallback DB lookup"
        parse_calls.append(name)
        return parse_product_name(name, brands, series)

    scope["parse_product_name"] = parse
    _functions({"products", "get_product_stats", "get_unknown_products"}, scope)
    yield SimpleNamespace(raw=raw, conn=c, scope=scope, calls=parse_calls)
    raw.close()


def test_products_reuses_recognition_across_ranking_trend_and_keeps_money_samples(endpoint):
    result = endpoint.scope["products"]()
    assert endpoint.calls == ["IGET One 12000 Puffs"]
    assert result["totals"]["total_quantity"] == 6
    assert result["totals"]["total_revenue_by_currency"] == {"PLN": 38, "CZK": 45}
    assert result["totals"]["total_gross_revenue_by_currency"] == {"PLN": 55, "CZK": 55}
    assert result["totals"]["total_revenue_cny"] == 89.5
    assert result["totals"]["total_gross_revenue_cny"] == 126.5
    assert result["totals"]["shipping_loss_by_currency"] == {"PLN": 7}
    assert [p["quantity"] for p in result["top_products"]] == [4, 2]
    first = result["top_products"][0]
    assert first["series"] == "One"
    assert first["order_count"] == 2
    assert [p["order_count"] for p in first["source_prices"]] == [1, 1]
    assert len(first["recent_orders"]) == 2
    assert not any(key.startswith("_") for key in first)
    assert len([sql for sql, _ in endpoint.conn.queries if "FROM series" in sql]) == 1


@pytest.mark.parametrize("allowed,expected", [(None, 6), (["https://one.invalid"], 3), ([], 0)])
def test_stats_keeps_source_permissions_and_bounded_name_parsing(endpoint, allowed, expected):
    endpoint.scope["get_user_allowed_sources"] = lambda *_: allowed
    result = endpoint.scope["get_product_stats"]()
    assert sum(p["quantity"] for p in result) == expected
    assert len(endpoint.calls) == (1 if expected else 0)
    assert all("FROM series" not in sql for sql, _ in endpoint.conn.queries)


@pytest.mark.parametrize("allowed,expected_sources", [
    (None, ["https://one.invalid", "https://two.invalid"]),
    (["https://one.invalid"], ["https://one.invalid"]), ([], []),
])
def test_unknown_api_cannot_return_other_sites_and_respects_source_manual_mapping(endpoint, allowed, expected_sources):
    name = "Mystery"
    now = date.today().isoformat()
    items = json.dumps([_item(name=name, quantity=2)])
    for sid in (1, 2):
        source = "https://one.invalid" if sid == 1 else "https://two.invalid"
        endpoint.raw.execute("INSERT INTO orders VALUES(?,?,?,?,'PLN',20,0,?,'completed',0,0)",
                             (f"{sid}-200", "200", items, source, now))
    endpoint.raw.commit()
    endpoint.scope["get_user_allowed_sources"] = lambda *_: allowed
    result = endpoint.scope["get_unknown_products"]()
    assert (result[0]["sources"] if result else []) == expected_sources
    endpoint.raw.execute("INSERT INTO product_mappings VALUES('Mystery','https://one.invalid',30000,'',NULL,1,1)")
    endpoint.raw.commit()
    result = endpoint.scope["get_unknown_products"]()
    remaining = [source for source in expected_sources if source != "https://one.invalid"]
    assert (result[0]["sources"] if result else []) == remaining
