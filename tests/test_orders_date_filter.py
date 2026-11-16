"""订单列表与客户分析页面的快捷日期及自定义日期区间筛选回归测试。"""

import ast
import json
import sqlite3
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from werkzeug.datastructures import MultiDict

from flask import Flask, request

ROOT = Path(__file__).resolve().parents[1]


def _load_date_helpers_and_orders(db_path):
    """从 app.py 加载真实的日期解析辅助函数与 orders 视图函数。"""
    source = (ROOT / "app.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    target_names = {"_quick_date_bounds", "_resolve_date_filter_params", "orders"}
    nodes = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in target_names:
            copied = ast.parse(ast.unparse(node)).body[0]
            copied.decorator_list = []
            nodes.append(copied)

    def connect():
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        return conn

    captured = {}

    def fake_render_template(template_name, **context):
        captured["template"] = template_name
        captured["context"] = context
        return json.dumps({"template": template_name, "total": context.get("total")})

    namespace = {
        "get_db_connection": connect,
        "get_user_allowed_sources": lambda *_args: None,
        "current_user": SimpleNamespace(
            id=1,
            is_admin=lambda: True,
            is_viewer=lambda: False,
            can_manage_own_site_sync=lambda: True,
        ),
        "request": request,
        "get_all_managers": lambda: ["Alice"],
        "_revenue_status_cond": lambda: "status IN ('completed', 'processing')",
        "_revenue_amount_case": lambda col: f"SUM(CASE WHEN status IN ('completed', 'processing') THEN {col} ELSE 0 END)",
        "parse_json_field": lambda raw: json.loads(raw or "[]"),
        "_build_risk_index": lambda _conn: {},
        "get_display_product_name": lambda item: item.get("name", ""),
        "extract_custom_billing_fields": lambda _meta: {},
        "_assess_customer_risk": lambda *_args, **_kwargs: None,
        "visible_fulfillment_state": lambda _status, state: state,
        "get_cny_rate": lambda _cur, month: (1.8, month),
        "_success_status_case": lambda: "CASE WHEN status = 'completed' THEN 1 ELSE 0 END",
        "_success_amount_case": lambda col: f"SUM(CASE WHEN status = 'completed' THEN {col} ELSE 0 END)",
        "_bad_order_case": lambda: "0",
        "_meaningful_order_case": lambda: "1",
        "calculate_customer_tier": lambda *_args, **_kwargs: "normal",
        "_can_manage_all_settings": lambda _user: True,
        "render_template": fake_render_template,
    }
    exec(compile(ast.Module(body=nodes, type_ignores=[]), "app.py", "exec"), namespace)
    return namespace, captured


def _init_orders_db(db_path):
    """创建包含 10 月与 11 月不同状态订单的隔离测试库。"""
    conn = sqlite3.connect(db_path)
    conn.executescript(
        """
        CREATE TABLE sites (url TEXT PRIMARY KEY, country TEXT, manager TEXT);
        CREATE TABLE customer_settings (email TEXT PRIMARY KEY, quality_tier TEXT);
        CREATE TABLE shipping_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id TEXT,
            tracking_number TEXT
        );
        CREATE TABLE orders (
            id TEXT PRIMARY KEY,
            number TEXT,
            status TEXT,
            total REAL,
            shipping_total REAL,
            currency TEXT,
            date_created TEXT,
            source TEXT,
            billing TEXT,
            shipping TEXT,
            line_items TEXT,
            meta_data TEXT,
            shipping_lines TEXT,
            payment_method TEXT,
            is_undelivered INTEGER DEFAULT 0,
            is_problem_return INTEGER DEFAULT 0,
            shipping_loss_amount REAL DEFAULT 0,
            product_loss_amount REAL DEFAULT 0
        );
        """
    )
    conn.execute(
        "INSERT INTO sites(url, country, manager) VALUES (?, ?, ?)",
        ("https://shop.example", "PL", "Alice"),
    )
    sample_orders = [
        ("oct-comp-1", "1001", "completed", 120.0, 10.0, "2026-10-01T08:30:00"),
        ("oct-comp-2", "1002", "completed", 200.0, 15.0, "2026-10-31T21:15:00"),
        ("oct-proc-1", "1003", "processing", 90.0, 10.0, "2026-10-15T12:00:00"),
        ("nov-comp-1", "1101", "completed", 150.0, 10.0, "2026-11-10T09:00:00"),
    ]
    billing = json.dumps({"first_name": "Jan", "last_name": "Kowalski", "email": "jan@example.com"})
    items = json.dumps([{"name": "Item A", "quantity": 2, "total": "100.00"}])
    for oid, number, status, total, ship, created in sample_orders:
        conn.execute(
            """INSERT INTO orders(
                   id, number, status, total, shipping_total, currency, date_created,
                   source, billing, shipping, line_items, meta_data, shipping_lines, payment_method
               ) VALUES (?, ?, ?, ?, ?, 'PLN', ?, 'https://shop.example', ?, '{}', ?, '[]', '[]', 'cod')""",
            (oid, number, status, total, ship, created, billing, items),
        )
    conn.commit()
    conn.close()


def test_resolve_date_filter_params_prioritizes_modified_custom_range():
    """表单隐藏域携带旧 quick_date=this_month 时，手动选择 10 月日期区间必须覆盖快捷预设。"""
    ns, _ = _load_date_helpers_and_orders(":memory:")
    resolve = ns["_resolve_date_filter_params"]
    nov_today = date(2026, 11, 16)

    # 首次无参进入：默认本月至今
    assert resolve(MultiDict(), default_quick_date="this_month", today=nov_today) == (
        "this_month",
        "2026-11-01",
        "2026-11-16",
    )

    # 复现线上缺陷：隐藏域仍为 this_month，但用户已选择 10月1日 到 10月31日
    args_custom_oct = MultiDict(
        [
            ("quick_date", "this_month"),
            ("status", "completed"),
            ("date_from", "2026-10-01"),
            ("date_to", "2026-10-31"),
        ]
    )
    assert resolve(args_custom_oct, default_quick_date="this_month", today=nov_today) == (
        "",
        "2026-10-01",
        "2026-10-31",
    )

    # 前端清空隐藏域后提交 10月1日 到 10月31日
    args_cleared_qd = MultiDict(
        [
            ("quick_date", ""),
            ("status", "completed"),
            ("date_from", "2026-10-01"),
            ("date_to", "2026-10-31"),
        ]
    )
    assert resolve(args_cleared_qd, default_quick_date="this_month", today=nov_today) == (
        "",
        "2026-10-01",
        "2026-10-31",
    )

    # 未改动日期、仅切换订单状态时，保留 this_month 高亮状态
    args_unchanged_dates = MultiDict(
        [
            ("quick_date", "this_month"),
            ("status", "completed"),
            ("date_from", "2026-11-01"),
            ("date_to", "2026-11-16"),
        ]
    )
    assert resolve(args_unchanged_dates, default_quick_date="this_month", today=nov_today) == (
        "this_month",
        "2026-11-01",
        "2026-11-16",
    )

    # 用户主动清空两个日期框提交表单时，查询全部日期而非跳回本月
    args_cleared_dates = MultiDict(
        [
            ("quick_date", "this_month"),
            ("status", "completed"),
            ("date_from", ""),
            ("date_to", ""),
        ]
    )
    assert resolve(args_cleared_dates, default_quick_date="this_month", today=nov_today) == (
        "all",
        "",
        "",
    )

    # 点击顶部快捷筛选链接（无 date_from / date_to 参数）正常切回预设范围
    args_quick_link = MultiDict([("quick_date", "last_month"), ("status", "completed")])
    assert resolve(args_quick_link, default_quick_date="this_month", today=nov_today) == (
        "last_month",
        "2026-10-01",
        "2026-10-31",
    )

    # 客户分析默认口径为 all，切换 this_month 后改选 10 月同样清除过期 quick_date
    assert resolve(MultiDict(), default_quick_date="all", today=nov_today) == ("all", "", "")
    assert resolve(args_custom_oct, default_quick_date="all", today=nov_today) == (
        "",
        "2026-10-01",
        "2026-10-31",
    )


def test_orders_endpoint_filters_october_completed_without_jumping_to_november(tmp_path):
    """验证 /orders 路由在 11 月筛选 10月1日～10月31日 已完成订单时准确返回 10 月数据。"""
    db_path = tmp_path / "orders_filter.db"
    _init_orders_db(db_path)
    ns, captured = _load_date_helpers_and_orders(db_path)

    app = Flask("orders-date-filter-test")
    with app.test_request_context(
        "/orders?quick_date=this_month&status=completed&date_from=2026-10-01&date_to=2026-10-31"
    ):
        ns["orders"]()

    ctx = captured["context"]
    assert ctx["current_filters"]["quick_date"] == ""
    assert ctx["current_filters"]["date_from"] == "2026-10-01"
    assert ctx["current_filters"]["date_to"] == "2026-10-31"
    assert ctx["current_filters"]["status"] == "completed"
    assert ctx["current_month"] == "2026-10"
    assert ctx["available_months"] == ["2026-10"]
    assert [o["id"] for o in ctx["orders"]] == ["oct-comp-2", "oct-comp-1"]
    assert ctx["listing_total"] == 2
