import ast
import json
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace

from flask import Flask, jsonify, request

ROOT = Path(__file__).resolve().parents[1]


def _create_shipping_db(path):
    conn = sqlite3.connect(path)
    conn.executescript(
        """
        CREATE TABLE orders (
            id TEXT PRIMARY KEY,
            number TEXT,
            status TEXT,
            total REAL,
            currency TEXT,
            date_created TEXT,
            date_modified TEXT,
            source TEXT,
            billing TEXT,
            shipping TEXT,
            line_items TEXT,
            meta_data TEXT,
            shipping_lines TEXT,
            shipping_total REAL,
            customer_note TEXT,
            warehouse_id INTEGER,
            is_undelivered INTEGER DEFAULT 0,
            shipping_loss_amount REAL DEFAULT 0,
            undelivered_at TEXT,
            undelivered_note TEXT,
            undelivered_by INTEGER,
            is_problem_return INTEGER DEFAULT 0,
            problem_return_type TEXT,
            product_loss_amount REAL DEFAULT 0,
            problem_return_at TEXT,
            carrier_status TEXT,
            carrier_status_at TEXT
        );
        CREATE TABLE sites (url TEXT PRIMARY KEY, country TEXT, manager TEXT);
        CREATE TABLE warehouses (id INTEGER PRIMARY KEY, name TEXT);
        CREATE TABLE shipping_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id TEXT,
            tracking_number TEXT,
            carrier_slug TEXT,
            shipped_at TEXT,
            is_reship INTEGER DEFAULT 0,
            reship_reason TEXT
        );
        CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT);
        CREATE TABLE order_notes (
            order_id TEXT,
            note TEXT,
            date_created TEXT,
            author TEXT,
            customer_note INTEGER DEFAULT 0
        );
        CREATE TABLE shipping_carriers (slug TEXT PRIMARY KEY, name TEXT, tracking_url TEXT);
        """
    )
    conn.execute(
        "INSERT INTO sites(url, country, manager) VALUES (?, ?, ?)",
        ("https://shop.example", "PL", "Alice"),
    )
    conn.execute(
        "INSERT INTO shipping_carriers(slug, name, tracking_url) VALUES (?, ?, ?)",
        ("dpd", "DPD", "https://tracking.example/{tracking}"),
    )

    start = datetime(2026, 8, 1)
    for index in range(1, 56):
        stamp = (start + timedelta(minutes=index)).strftime("%Y-%m-%d %H:%M:%S")
        billing = json.dumps({"first_name": "Test", "last_name": str(index), "email": f"buyer{index}@example.com"})
        shipping = json.dumps({"first_name": "Test", "last_name": str(index), "address_1": f"Street {index}"})
        products = json.dumps([{"name": "Product", "quantity": 1, "total": "10.00"}])
        conn.execute(
            """INSERT INTO orders(
                   id, number, status, total, currency, date_created, date_modified,
                   source, billing, shipping, line_items, meta_data, shipping_lines,
                   shipping_total, customer_note
               ) VALUES (?, ?, 'shipped', 15, 'PLN', ?, ?, ?, ?, ?, ?, '[]', '[]', 5, '')""",
            (
                f"order-{index:02d}",
                f"ORDER{index:04d}",
                stamp,
                stamp,
                "https://shop.example",
                billing,
                shipping,
                products,
            ),
        )
        conn.execute(
            """INSERT INTO shipping_logs(order_id, tracking_number, carrier_slug, shipped_at)
               VALUES (?, ?, 'dpd', ?)""",
            (f"order-{index:02d}", f"TRACK{index:04d}", stamp),
        )

    conn.execute(
        """INSERT INTO orders(
               id, number, status, total, currency, date_created, date_modified,
               source, billing, shipping, line_items, meta_data, shipping_lines,
               shipping_total, customer_note
           ) VALUES ('pending-1', 'PENDING1', 'processing', 10, 'PLN', ?, ?, ?, '{}', '{}', '[]', '[]', '[]', 0, '')""",
        (start.isoformat(), start.isoformat(), "https://shop.example"),
    )
    conn.commit()
    conn.close()


def _load_shipped_endpoint(db_path, allowed_sources):
    def connect():
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        return conn

    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    endpoint_node = next(
        node
        for node in tree.body
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
        and node.name == "get_shipped_orders"
    )
    endpoint_node.decorator_list = []
    namespace = {
        "get_db_connection": connect,
        "get_user_allowed_sources": lambda *_args: allowed_sources,
        "current_user": SimpleNamespace(
            id=1,
            is_admin=lambda: True,
            is_viewer=lambda: False,
        ),
        "request": request,
        "jsonify": jsonify,
        "_build_risk_index": lambda _conn: {},
        "process_shipped_order": lambda row, *_args: {
            "id": row["id"],
            "number": row["number"],
        },
        "parse_json_field": lambda value: json.loads(value or "{}"),
        "_assess_customer_risk": lambda *_args, **_kwargs: None,
    }
    exec(compile(ast.Module(body=[endpoint_node], type_ignores=[]), "app.py", "exec"), namespace)
    return namespace["get_shipped_orders"]


def _call_shipped_endpoint(db_path, query_string, allowed_sources=None):
    app = Flask("shipping-pagination-test")
    endpoint = _load_shipped_endpoint(db_path, allowed_sources)
    with app.test_request_context(f"/api/shipping/shipped?{query_string}"):
        return endpoint().get_json()


def test_shipped_orders_are_bounded_to_fifty_per_page(tmp_path):
    db_path = tmp_path / "shipping.db"
    _create_shipping_db(db_path)

    first = _call_shipped_endpoint(db_path, "page=1")
    second = _call_shipped_endpoint(db_path, "page=2")

    assert first["pagination"] == {
        "page": 1,
        "per_page": 50,
        "total": 55,
        "pages": 2,
        "has_previous": False,
        "has_next": True,
    }
    assert len(first["orders"]) == 50
    assert len(second["orders"]) == 5
    assert {row["id"] for row in first["orders"]}.isdisjoint(
        {row["id"] for row in second["orders"]}
    )


def test_shipped_page_and_total_respect_search_and_permissions(tmp_path):
    db_path = tmp_path / "shipping.db"
    _create_shipping_db(db_path)

    filtered = _call_shipped_endpoint(db_path, "page=9&search=ORDER0001")
    assert filtered["pagination"]["page"] == 1
    assert filtered["pagination"]["total"] == 1
    assert [row["number"] for row in filtered["orders"]] == ["ORDER0001"]

    denied = _call_shipped_endpoint(db_path, "page=1", allowed_sources=[])
    assert denied["pagination"]["total"] == 0
    assert denied["orders"] == []


def test_shipping_template_renders_server_pagination_controls():
    template = (ROOT / "templates" / "shipping.html").read_text(encoding="utf-8")

    assert 'id="shippedPagination"' in template
    assert 'id="shippedPageSummary"' in template
    assert 'id="shippedPageButtons"' in template
    assert "page: shippedPage" in template
    assert "const orders = data.orders || []" in template
    assert "loadOrders(page, true)" in template


def test_high_risk_shipping_postcode_normalizes_common_formats():
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    helper_node = next(
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef)
        and node.name == "_is_high_risk_shipping_postcode"
    )
    namespace = {}
    exec(compile(ast.Module(body=[helper_node], type_ignores=[]), "app.py", "exec"), namespace)
    is_high_risk = namespace["_is_high_risk_shipping_postcode"]

    assert is_high_risk("66-600")
    assert is_high_risk("66600")
    assert is_high_risk("66 600")
    assert is_high_risk(" 66–600 ")
    assert not is_high_risk("66-601")
    assert not is_high_risk("66-60")
    assert not is_high_risk(None)


def test_shipping_template_renders_high_risk_postcode_warning():
    template = (ROOT / "templates" / "shipping.html").read_text(encoding="utf-8")

    assert "o.high_risk_postcode" in template
    assert "高危邮编 ${o.high_risk_postcode}" in template
    assert "白嫖风险地区 · 邮编 ${o.high_risk_postcode}" in template
    assert "postcodeRiskBanner" in template
