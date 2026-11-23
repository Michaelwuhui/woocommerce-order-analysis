"""客户等级评定、问题退货/拒收风控熔断、跨身份联动及客户详情统计单元测试。"""

from __future__ import annotations

import ast
import json
from pathlib import Path
import sqlite3
from types import SimpleNamespace

from flask import Flask
import pytest

from customer_table_data import customer_table_context
import customer_tier_service as cts

ROOT = Path(__file__).resolve().parents[1]


def test_marta_kubiak_real_case_never_gets_good_or_vip():
    """复现线上真实案例 Marta Kubiak：存在问题退货（调包/扣留货损），一票否决判定为劣质（bad），严禁评为优质或 VIP。"""
    tier = cts.calculate_customer_tier(
        successful_orders=4,
        total_spending=945.10,
        avg_days_between=17.5,
        bad_orders=1,
        meaningful_orders=5,
        problem_return_orders=1,
        undelivered_orders=0,
    )
    assert tier == "bad"


def test_repeat_problem_returns_or_combined_bad_orders_become_bad_tier():
    """验证 2 次及以上问题退货、问题退货叠加拒收、高比例拒收或已拉黑客户直接判定为劣质（bad）。"""
    # 8 单中 6 单成功但有 2 单问题退货 -> 直接劣质
    assert (
        cts.calculate_customer_tier(
            6, 1800.0, 15.0, bad_orders=2, meaningful_orders=8, problem_return_orders=2
        )
        == "bad"
    )
    # 7 单中 4 单成功 + 1 单问题退货 + 2 单未送达拒收 -> 直接劣质
    assert (
        cts.calculate_customer_tier(
            4, 900.0, 20.0, bad_orders=3, meaningful_orders=7, problem_return_orders=1, undelivered_orders=2
        )
        == "bad"
    )
    # 新客 2 单中 1 单成功 + 1 单问题退货 -> 直接劣质
    assert (
        cts.calculate_customer_tier(
            1, 300.0, 10.0, bad_orders=1, meaningful_orders=2, problem_return_orders=1
        )
        == "bad"
    )
    # 3 单拒收 2 单（无问题退货，拒收率 66.7%） -> 直接劣质
    assert (
        cts.calculate_customer_tier(
            1, 200.0, 12.0, bad_orders=2, meaningful_orders=3, undelivered_orders=2
        )
        == "bad"
    )
    # 已被手机号黑名单拦截，哪怕曾有 10 单成功也强制为劣质
    assert (
        cts.calculate_customer_tier(
            10, 5000.0, 10.0, bad_orders=0, meaningful_orders=10, is_blocked=True
        )
        == "bad"
    )


def test_clean_loyal_customers_still_earn_good_and_vip():
    """验证无不良退货记录的正常复购客户依然正常获评优质（good）和 VIP。"""
    assert cts.calculate_customer_tier(3, 600.0, 25.0, bad_orders=0, meaningful_orders=3) == "good"
    assert cts.calculate_customer_tier(5, 3500.0, 20.0, bad_orders=0, meaningful_orders=5) == "vip"
    # 10 单老客偶发 1 次未取件（不良率 10%，无问题退货）：保留优质但不升 VIP
    assert (
        cts.calculate_customer_tier(
            9, 4000.0, 20.0, bad_orders=1, meaningful_orders=10, undelivered_orders=1
        )
        == "good"
    )


def test_smart_actions_and_top_customers_exclude_problem_returners():
    """验证客户管理页的建议行动与 Top10 高价值榜单不会优待问题退货或劣质客户。"""
    act_prob = cts.build_smart_actions(
        tier_cn="普通",
        successful_orders=4,
        problem_return_orders=1,
        undelivered_orders=0,
        is_blocked=False,
        first_order_date="2026-09-01",
        last_order_date="2026-06-01",
        thirty_days_ago="2026-10-24",
        ninety_days_ago="2026-08-25",
    )
    assert act_prob == [{"type": "danger", "icon": "exclamation-octagon", "text": "问题退货·严审COD"}]

    ctx = customer_table_context(
        [
            {"name": "Problem Spender", "email": "p@x.pl", "tier": "普通", "problem_return_orders": 1, "total_spent_cny": 9999},
            {"name": "Bad Spender", "email": "b@x.pl", "tier": "劣质", "problem_return_orders": 0, "total_spent_cny": 8888},
            {"name": "Clean VIP", "email": "v@x.pl", "tier": "VIP", "problem_return_orders": 0, "total_spent_cny": 5000},
        ]
    )
    assert [c["name"] for c in ctx["top_customers"]] == ["Clean VIP"]


def test_get_customer_details_excludes_problem_return_from_success_and_spending(tmp_path):
    """验证客户详情 API 不会把问题退货订单算进成功订单数或总消费，且等级不再虚标优质。"""
    db_file = str(tmp_path / "cust_detail.db")
    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    conn.execute("CREATE TABLE sites (id INTEGER PRIMARY KEY, url TEXT, cod_on_hold_is_shipped INTEGER DEFAULT 1)")
    conn.execute("INSERT INTO sites VALUES (13, 'https://vapepl.pl', 1)")
    conn.execute("CREATE TABLE customer_settings (email TEXT PRIMARY KEY, quality_tier TEXT, updated_at TEXT)")
    conn.execute(
        "CREATE TABLE blocked_customers (phone TEXT PRIMARY KEY, raw_phone TEXT, name TEXT, email TEXT, "
        "reason TEXT, auto_cancel INTEGER, created_by TEXT, created_at TEXT, cancelled_count INTEGER, last_cancelled_at TEXT)"
    )
    conn.execute(
        """
        CREATE TABLE orders (
            id TEXT PRIMARY KEY, number TEXT, status TEXT, total REAL, currency TEXT,
            shipping_total REAL, date_created TEXT, source TEXT, line_items TEXT,
            billing TEXT, shipping TEXT, payment_method TEXT,
            is_undelivered INTEGER DEFAULT 0, shipping_loss_amount REAL DEFAULT 0,
            is_problem_return INTEGER DEFAULT 0, problem_return_type TEXT,
            product_loss_amount REAL DEFAULT 0, problem_return_note TEXT,
            delivery_confirmed INTEGER DEFAULT 0, undelivered_note TEXT, carrier_status TEXT
        )
        """
    )
    billing_json = json.dumps({"first_name": "Marta", "last_name": "Kubiak", "email": "marta-kubiak@o2.pl", "phone": "+48669324762"})
    items_json = json.dumps([{"name": "Vape", "quantity": 1, "total": "100.00"}])
    sample_rows = [
        ("13-8314", "8314", "shipped", 190.8, "2026-11-17T22:04:49", 0, 0.0, 0.0),
        ("13-7852", "7852", "completed", 107.9, "2026-10-12T09:37:30", 0, 0.0, 0.0),
        ("13-7474", "7474", "completed", 289.8, "2026-10-01T07:13:25", 0, 0.0, 0.0),
        ("13-7360", "7360", "shipped", 538.5, "2026-09-23T13:09:31", 1, 25.0, 204.0),
        ("13-6836", "6836", "completed", 356.6, "2026-09-08T07:07:18", 0, 0.0, 0.0),
    ]
    for oid, num, st, tot, dt, is_prob, ship_loss, prod_loss in sample_rows:
        conn.execute(
            """
            INSERT INTO orders
                (id, number, status, total, currency, shipping_total, date_created, source,
                 line_items, billing, shipping, payment_method, is_problem_return,
                 problem_return_type, shipping_loss_amount, product_loss_amount)
            VALUES (?, ?, ?, ?, 'PLN', 15.0, ?, 'https://vapepl.pl', ?, ?, '{}', 'cod', ?, ?, ?, ?)
            """,
            (oid, num, st, tot, dt, items_json, billing_json, is_prob, "swap" if is_prob else None, ship_loss, prod_loss),
        )
    conn.commit()
    conn.close()

    source = (ROOT / "app.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    needed = {
        "_normalize_email", "_is_placeholder_phone", "_normalize_phone",
        "_normalize_address", "_addr_for_order", "_resolve_identity_clusters",
        "extract_flavor_from_meta", "extract_puffs_from_meta", "get_full_product_name",
        "calculate_customer_tier", "get_customer_details",
    }
    fn_nodes = []
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in needed:
            copied = ast.parse(ast.unparse(node)).body[0]
            copied.decorator_list = []
            fn_nodes.append(copied)

    def connect():
        c = sqlite3.connect(db_file)
        c.row_factory = sqlite3.Row
        return c

    app = Flask(__name__)
    from flask import jsonify
    import html as _html

    ns = {
        "_IDENTITY_SHARED_KEY_LIMIT": 4,
        "get_db_connection": connect,
        "get_user_allowed_sources": lambda *_a: None,
        "current_user": SimpleNamespace(id=1, is_admin=lambda: True, is_viewer=lambda: False),
        "parse_json_field": lambda raw: json.loads(raw or "{}"),
        "get_cny_rate": lambda _cur, m: (1.8, m),
        "jsonify": jsonify,
        "html": _html,
    }
    mod = ast.Module(body=fn_nodes, type_ignores=[])
    ast.fix_missing_locations(mod)
    exec(compile(mod, str(ROOT / "app.py"), "exec"), ns)

    with app.test_request_context("/api/customer/marta-kubiak@o2.pl"):
        resp = ns["get_customer_details"]("marta-kubiak@o2.pl")
        data = resp.get_json()

    assert data["total_orders"] == 5
    assert data["successful_orders"] == 4
    assert data["problem_return_orders"] == 1
    assert round(data["total_spending"], 2) == 945.10
    assert data["product_loss_total"] == 204.0
    assert data["shipping_loss_total"] == 25.0
    assert data["customer_tier"]["level"] == "劣质"
