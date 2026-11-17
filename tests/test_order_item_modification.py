"""订单详情更换口味/更换产品全链路回归测试（权限、拦截、拆分金额、源站同步及多仓履约重算）。"""

from __future__ import annotations

import copy
import json
import sqlite3
from decimal import Decimal
from types import SimpleNamespace
from urllib.parse import parse_qs, urlsplit

import pytest

from fulfillment_service import plan_order, transition_fulfillment
from inv_migrations import (
    up_001,
    up_006,
    up_007,
    up_008,
    up_009,
    up_010,
    up_011,
    up_012,
)
from order_item_modification import (
    compute_line_item_replacement_amounts,
    evaluate_order_modification_eligibility,
    get_product_options_for_order,
    modify_order_line_item,
    search_site_products_for_order,
)


def _make_user(uid: int, username: str, name: str, role: str = "admin"):
    return SimpleNamespace(
        id=uid,
        username=username,
        name=name,
        role=role,
        is_authenticated=True,
        is_admin=lambda: role == "admin",
        is_viewer=lambda: role == "viewer",
    )


class FakeResponse:
    def __init__(self, status_code: int, body, headers: dict | None = None):
        self.status_code = status_code
        self._body = body
        self.headers = headers or {}
        self.text = json.dumps(body, ensure_ascii=False)

    def json(self):
        return copy.deepcopy(self._body)


class FakeWooSession:
    """模拟真实 WooCommerce REST API v3 的订单读取、商品变体查询、订单行增删改及备注创建。"""

    def __init__(self):
        self.next_item_id = 50000
        self.next_note_id = 9000
        self.put_calls: list[dict] = []
        self.notes_created: list[dict] = []

        self.products = {
            100: {
                "id": 100,
                "name": "Fumot RandM Tornado 15000",
                "sku": "FUMOT-15K",
                "type": "variable",
                "status": "publish",
                "price": "64.90",
                "regular_price": "64.90",
                "stock_status": "instock",
                "stock_quantity": None,
                "attributes": [
                    {
                        "id": 1,
                        "name": "Smaki",
                        "slug": "pa_smaki",
                        "variation": True,
                        "options": ["MIXED BERRY", "WATERMELON ICE"],
                    }
                ],
                "variations": [1001, 1002],
            },
            200: {
                "id": 200,
                "name": "MerryMi Blade 30000",
                "sku": "BLADE-30K",
                "type": "variable",
                "status": "publish",
                "price": "89.90",
                "regular_price": "89.90",
                "stock_status": "instock",
                "stock_quantity": None,
                "attributes": [
                    {
                        "id": 1,
                        "name": "Smaki",
                        "slug": "pa_smaki",
                        "variation": True,
                        "options": ["BLUE RAZZ"],
                    }
                ],
                "variations": [2001],
            },
            300: {
                "id": 300,
                "name": "Wildcard Any Flavor Pod",
                "sku": "ANY-POD",
                "type": "variable",
                "status": "publish",
                "price": "49.90",
                "regular_price": "49.90",
                "stock_status": "instock",
                "stock_quantity": 15,
                "attributes": [
                    {
                        "id": 2,
                        "name": "Smaki",
                        "slug": "smaki",
                        "variation": True,
                        "options": ["Peach Ice", "Cola Ice"],
                    }
                ],
                "variations": [3001],
            },
        }

        self.variations = {
            100: [
                {
                    "id": 1001,
                    "name": "MIXED BERRY",
                    "sku": "FUMOT-15K-MIXED",
                    "status": "publish",
                    "price": "64.90",
                    "regular_price": "64.90",
                    "stock_status": "instock",
                    "stock_quantity": 12,
                    "attributes": [{"id": 1, "name": "Smaki", "slug": "pa_smaki", "option": "MIXED BERRY"}],
                },
                {
                    "id": 1002,
                    "name": "WATERMELON ICE",
                    "sku": "FUMOT-15K-WATERMELON",
                    "status": "publish",
                    "price": "69.90",  # 故意设高价以验证同款换口味仍保持原价不变
                    "regular_price": "69.90",
                    "stock_status": "instock",
                    "stock_quantity": 8,
                    "attributes": [{"id": 1, "name": "Smaki", "slug": "pa_smaki", "option": "WATERMELON ICE"}],
                },
            ],
            200: [
                {
                    "id": 2001,
                    "name": "BLUE RAZZ",
                    "sku": "BLADE-30K-BLUERAZZ",
                    "status": "publish",
                    "price": "89.90",
                    "regular_price": "89.90",
                    "stock_status": "instock",
                    "stock_quantity": 20,
                    "attributes": [{"id": 1, "name": "Smaki", "slug": "pa_smaki", "option": "BLUE RAZZ"}],
                }
            ],
            300: [
                {
                    "id": 3001,
                    "name": "Any Flavor",
                    "sku": "ANY-POD-VAR",
                    "status": "publish",
                    "price": "49.90",
                    "regular_price": "49.90",
                    "stock_status": "instock",
                    "stock_quantity": 15,
                    "attributes": [],  # 任意口味变体：省略具体 option
                }
            ],
        }

        self.orders = {
            9001: {
                "id": 9001,
                "number": "9001",
                "status": "processing",
                "currency": "PLN",
                "shipping_total": "25.00",
                "discount_total": "0.00",
                "total_tax": "0.00",
                "cart_tax": "0.00",
                "total": "89.90",
                "date_modified": "2026-11-17T09:00:00",
                "date_modified_gmt": "2026-11-17T08:00:00",
                "meta_data": [],
                "shipping_lines": [{"id": 1, "method_title": "InPost", "total": "25.00"}],
                "fee_lines": [],
                "coupon_lines": [],
                "tax_lines": [],
                "line_items": [
                    {
                        "id": 701,
                        "name": "Fumot RandM Tornado 15000 - MIXED BERRY",
                        "product_id": 100,
                        "variation_id": 1001,
                        "sku": "FUMOT-15K-MIXED",
                        "quantity": 1,
                        "price": 64.90,
                        "subtotal": "64.90",
                        "total": "64.90",
                        "meta_data": [
                            {
                                "id": 1,
                                "key": "pa_smaki",
                                "value": "mixed-berry",
                                "display_key": "Smaki",
                                "display_value": "MIXED BERRY",
                            }
                        ],
                    }
                ],
            },
            9002: {
                "id": 9002,
                "number": "9002",
                "status": "processing",
                "currency": "PLN",
                "shipping_total": "20.00",
                "discount_total": "0.00",
                "total_tax": "0.00",
                "cart_tax": "0.00",
                "total": "120.00",
                "date_modified": "2026-11-17T09:00:00",
                "date_modified_gmt": "2026-11-17T08:00:00",
                "meta_data": [],
                "shipping_lines": [{"id": 2, "method_title": "DPD", "total": "20.00"}],
                "fee_lines": [],
                "coupon_lines": [],
                "tax_lines": [],
                "line_items": [
                    {
                        "id": 801,
                        "name": "Fumot RandM Tornado 15000 - MIXED BERRY",
                        "product_id": 100,
                        "variation_id": 1001,
                        "sku": "FUMOT-15K-MIXED",
                        "quantity": 3,
                        "price": 33.33,
                        "subtotal": "100.00",
                        "total": "100.00",
                        "meta_data": [
                            {
                                "id": 2,
                                "key": "pa_smaki",
                                "value": "mixed-berry",
                                "display_key": "Smaki",
                                "display_value": "MIXED BERRY",
                            }
                        ],
                    }
                ],
            },
        }

    def request(self, *, method, url, params=None, json=None, auth=None, headers=None, timeout=None, allow_redirects=False):
        parsed = urlsplit(url)
        path = parsed.path.split("/wp-json/wc/v3/", 1)[1].strip("/")
        parts = path.split("/")
        qs = parse_qs(parsed.query)
        merged_params = {k: v[0] for k, v in qs.items()}
        if params:
            merged_params.update(params)

        if method == "GET" and parts[0] == "orders" and len(parts) == 2:
            oid = int(parts[1])
            order = self.orders.get(oid)
            return FakeResponse(200, order) if order else FakeResponse(404, {"message": "Not found"})

        if method == "PUT" and parts[0] == "orders" and len(parts) == 2:
            oid = int(parts[1])
            order = self.orders[oid]
            self.put_calls.append({"order_id": oid, "body": copy.deepcopy(json)})
            items = list(order["line_items"])
            for patch in json.get("line_items") or []:
                if patch.get("id"):
                    target_id = int(patch["id"])
                    if int(patch.get("quantity", 1)) == 0:
                        items = [x for x in items if int(x["id"]) != target_id]
                    else:
                        for x in items:
                            if int(x["id"]) == target_id:
                                x["quantity"] = int(patch["quantity"])
                                if "subtotal" in patch:
                                    x["subtotal"] = str(patch["subtotal"])
                                if "total" in patch:
                                    x["total"] = str(patch["total"])
                                x["price"] = float(Decimal(x["total"]) / Decimal(x["quantity"]))
                else:
                    pid = int(patch["product_id"])
                    vid = int(patch.get("variation_id") or 0)
                    parent = self.products[pid]
                    var_obj = next((v for v in self.variations.get(pid, []) if v["id"] == vid), None)
                    self.next_item_id += 1
                    meta_list = []
                    if var_obj and var_obj.get("attributes"):
                        for a in var_obj["attributes"]:
                            meta_list.append(
                                {
                                    "key": a.get("slug") or a.get("name"),
                                    "value": a.get("option"),
                                    "display_key": a.get("name"),
                                    "display_value": a.get("option"),
                                }
                            )
                    for m in patch.get("meta_data") or []:
                        meta_list.append(
                            {
                                "key": m["key"],
                                "value": m["value"],
                                "display_key": m["key"],
                                "display_value": m["value"],
                            }
                        )
                    name = parent["name"]
                    if var_obj and var_obj.get("attributes"):
                        name = f"{parent['name']} - {var_obj['attributes'][0]['option']}"
                    qty = int(patch["quantity"])
                    tot = str(patch["total"])
                    items.append(
                        {
                            "id": self.next_item_id,
                            "name": name,
                            "product_id": pid,
                            "variation_id": vid,
                            "sku": (var_obj or parent)["sku"],
                            "quantity": qty,
                            "price": float(Decimal(tot) / Decimal(qty)),
                            "subtotal": str(patch["subtotal"]),
                            "total": tot,
                            "meta_data": meta_list,
                        }
                    )
            order["line_items"] = items
            merch_sum = sum(Decimal(str(i["total"])) for i in items)
            ship_sum = Decimal(str(order.get("shipping_total") or "0.00"))
            order["total"] = f"{(merch_sum + ship_sum):.2f}"
            order["date_modified"] = "2026-11-17T10:30:00"
            order["date_modified_gmt"] = "2026-11-17T09:30:00"
            return FakeResponse(200, order)

        if method == "POST" and parts[0] == "orders" and len(parts) == 3 and parts[2] == "notes":
            self.next_note_id += 1
            note_obj = {
                "id": self.next_note_id,
                "note": json["note"],
                "customer_note": bool(json.get("customer_note", False)),
                "date_created": "2026-11-17T10:30:01",
                "author": "WooCommerce",
            }
            self.notes_created.append(note_obj)
            return FakeResponse(201, note_obj)

        if method == "GET" and parts[0] == "products" and len(parts) == 1:
            q = str(merged_params.get("search") or "").lower()
            sku_q = str(merged_params.get("sku") or "").lower()
            prods = list(self.products.values())
            if q:
                prods = [p for p in prods if q in p["name"].lower()]
            elif sku_q:
                prods = [p for p in prods if sku_q == p["sku"].lower()]
            return FakeResponse(200, prods, headers={"X-WP-TotalPages": "1"})

        if method == "GET" and parts[0] == "products" and len(parts) == 2:
            pid = int(parts[1])
            prod = self.products.get(pid)
            return FakeResponse(200, prod) if prod else FakeResponse(404, {"message": "Not found"})

        if method == "GET" and parts[0] == "products" and len(parts) == 3 and parts[2] == "variations":
            pid = int(parts[1])
            return FakeResponse(200, self.variations.get(pid, []), headers={"X-WP-TotalPages": "1"})

        if method == "GET" and parts[0] == "products" and len(parts) == 4 and parts[2] == "variations":
            pid = int(parts[1])
            vid = int(parts[3])
            var_obj = next((v for v in self.variations.get(pid, []) if v["id"] == vid), None)
            return FakeResponse(200, var_obj) if var_obj else FakeResponse(404, {"message": "Not found"})

        raise AssertionError(f"未模拟的请求: {method} {url}")


@pytest.fixture
def env_setup():
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(
        """
        CREATE TABLE users (
            id INTEGER PRIMARY KEY,
            username TEXT,
            name TEXT,
            role TEXT
        );
        CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT);
        CREATE TABLE sites (
            id INTEGER PRIMARY KEY,
            url TEXT UNIQUE,
            country TEXT,
            manager TEXT,
            consumer_key TEXT,
            consumer_secret TEXT,
            api_status TEXT DEFAULT 'ok',
            api_write_status TEXT DEFAULT 'ok'
        );
        CREATE TABLE user_site_exclusions (
            user_id INTEGER,
            site_id INTEGER,
            PRIMARY KEY (user_id, site_id)
        );
        CREATE TABLE warehouses (
            id INTEGER PRIMARY KEY,
            code TEXT,
            name TEXT,
            country TEXT,
            is_active INTEGER DEFAULT 1
        );
        CREATE TABLE shipping_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id TEXT,
            tracking_number TEXT
        );
        CREATE TABLE order_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            wc_note_id INTEGER,
            order_id TEXT,
            note TEXT,
            date_created TEXT,
            customer_note INTEGER DEFAULT 0,
            author TEXT,
            added_by_user INTEGER DEFAULT 0,
            UNIQUE(order_id, wc_note_id)
        );
        CREATE TABLE orders (
            id TEXT PRIMARY KEY,
            woo_id INTEGER,
            number TEXT,
            status TEXT,
            source TEXT,
            line_items TEXT,
            billing TEXT,
            shipping TEXT,
            shipping_lines TEXT,
            fee_lines TEXT,
            coupon_lines TEXT,
            tax_lines TEXT,
            meta_data TEXT,
            total REAL,
            shipping_total REAL,
            discount_total REAL,
            total_tax REAL,
            cart_tax REAL,
            currency TEXT,
            payment_method TEXT,
            date_created TEXT,
            date_modified TEXT,
            date_modified_gmt TEXT,
            cart_hash TEXT,
            is_undelivered INTEGER DEFAULT 0,
            is_problem_return INTEGER DEFAULT 0,
            updated_at TEXT
        );
        """
    )

    conn.executemany(
        "INSERT INTO users(id, username, name, role) VALUES (?, ?, ?, ?)",
        [
            (1, "admin", "管理员", "admin"),
            (2, "michael", "吴辉", "viewer"),
            (17, "Jane", "张晴", "admin"),
            (3, "fiona", "付肖肖", "admin"),
            (10, "jinyi", "金毅", "user"),
        ],
    )
    conn.executemany(
        "INSERT INTO sites(id, url, country, manager, consumer_key, consumer_secret) VALUES (?, ?, ?, ?, ?, ?)",
        [
            (1, "https://jane-shop.pl", "PL", "张晴", "ck_jane", "cs_jane"),
            (2, "https://fiona-shop.pl", "PL", "付肖肖", "ck_fiona", "cs_fiona"),
            (3, "https://michael-shop.pl", "PL", "吴辉", "ck_m", "cs_m"),
        ],
    )
    conn.execute("INSERT INTO warehouses(id, code, name, country, is_active) VALUES (1, 'PL', '波兰主仓', 'PL', 1)")

    up_001(conn)
    up_006(conn)
    up_007(conn)
    up_008(conn)
    conn.execute("INSERT OR REPLACE INTO settings(key, value) VALUES ('oms_fulfillment_enabled', '1')")

    # 配置 3 个库存 SKU 与站点映射，验证多仓履约预留自动切换
    conn.executemany(
        "INSERT INTO inv_skus(id, sku_code, name, barcode, is_active) VALUES (?, ?, ?, ?, 1)",
        [
            (1, "SKU-15K-MIXED", "Fumot 15K Mixed Berry", "BAR-1001"),
            (2, "SKU-15K-WATERMELON", "Fumot 15K Watermelon Ice", "BAR-1002"),
            (3, "SKU-BLADE-BLUERAZZ", "Blade 30K Blue Razz", "BAR-2001"),
        ],
    )
    conn.executemany(
        """INSERT INTO inv_site_sku_map
           (site_id, wc_product_id, wc_variation_id, wc_sku, raw_name, sku_id, qty_per_item, is_active)
           VALUES (1, ?, ?, ?, ?, ?, 1, 1)""",
        [
            (100, 1001, "FUMOT-15K-MIXED", "Mixed Berry", 1),
            (100, 1002, "FUMOT-15K-WATERMELON", "Watermelon Ice", 2),
            (200, 2001, "BLADE-30K-BLUERAZZ", "Blue Razz", 3),
        ],
    )
    conn.executemany(
        """INSERT INTO oms_sku_warehouses
           (sku_id, warehouse_id, is_primary, is_enabled, wms_product_name_zh, wms_product_name_en, product_type)
           VALUES (?, 1, 1, 1, '测试烟', 'Vape', 'P')""",
        [(1,), (2,), (3,)],
    )
    conn.executemany(
        "INSERT INTO inv_stock(warehouse_id, sku_id, on_hand, reserved) VALUES (1, ?, 20, 0)",
        [(1,), (2,), (3,)],
    )

    woo = FakeWooSession()
    for local_id, woo_id, site_url in [
        ("1-9001", 9001, "https://jane-shop.pl"),
        ("1-9002", 9002, "https://jane-shop.pl"),
    ]:
        ro = woo.orders[woo_id]
        conn.execute(
            """INSERT INTO orders(
                   id, woo_id, number, status, source, line_items, billing, shipping,
                   total, shipping_total, discount_total, currency, payment_method, date_created, date_modified
               ) VALUES (?, ?, ?, ?, ?, ?, '{"country":"PL"}', '{"country":"PL"}', ?, ?, 0, 'PLN', 'cod', '2026-11-17T08:00:00', ?)""",
            (
                local_id,
                woo_id,
                ro["number"],
                ro["status"],
                site_url,
                json.dumps(ro["line_items"], ensure_ascii=False),
                float(ro["total"]),
                float(ro["shipping_total"]),
                ro["date_modified"],
            ),
        )
    conn.commit()
    yield conn, woo
    conn.close()


def test_site_owner_permission_strictly_isolated(env_setup):
    """站点负责人只能修改自己站点的订单；他人站点、发货员、只读查看员及排除站点均被拒绝。"""
    conn, woo = env_setup
    jane = _make_user(17, "Jane", "张晴", "admin")
    fiona = _make_user(3, "fiona", "付肖肖", "admin")
    shipper = _make_user(10, "jinyi", "金毅", "user")
    viewer_owner = _make_user(2, "michael", "吴辉", "viewer")
    super_admin = _make_user(1, "admin", "管理员", "admin")

    # 付肖肖尝试读取或修改张晴站点的订单 -> 403
    res_opt, st_opt = get_product_options_for_order(
        conn, "1-9001", user=fiona, product_id=100, session=woo
    )
    assert st_opt == 403
    assert res_opt["code"] == "site_owner_forbidden"

    res_mod, st_mod = modify_order_line_item(
        conn,
        "1-9001",
        {"line_item_id": 701, "replace_qty": 1, "target_product_id": 100, "target_variation_id": 1002},
        user=fiona,
        session=woo,
    )
    assert st_mod == 403
    assert len(woo.put_calls) == 0

    # 发货员金毅 -> 403
    _, st_ship = search_site_products_for_order(conn, "1-9001", user=shipper, session=woo)
    assert st_ship == 403

    # 只读角色 viewer（即使名字匹配负责人） -> 不可修改
    conn.execute(
        "INSERT INTO orders(id, woo_id, number, status, source, line_items, total, shipping_total, currency) VALUES ('3-1', 1, '1', 'processing', 'https://michael-shop.pl', '[]', 50, 10, 'PLN')"
    )
    m_order = conn.execute("SELECT * FROM orders WHERE id='3-1'").fetchone()
    assert evaluate_order_modification_eligibility(conn, m_order, viewer_owner)["can_modify"] is False
    assert evaluate_order_modification_eligibility(conn, m_order, super_admin)["can_modify"] is True

    # 张晴本人 -> 允许读取口味列表
    ok_res, ok_st = get_product_options_for_order(
        conn, "1-9001", user=jane, product_id=100, session=woo
    )
    assert ok_st == 200
    assert len(ok_res["options"]) == 2

    # 被 user_site_exclusions 排除后立即失去该站修改权
    conn.execute("INSERT INTO user_site_exclusions(user_id, site_id) VALUES (17, 1)")
    conn.commit()
    _, st_excl = get_product_options_for_order(
        conn, "1-9001", user=jane, product_id=100, session=woo
    )
    assert st_excl == 403


def test_shipped_or_picking_orders_are_blocked(env_setup):
    """已发货（on-hold / shipping_logs）或仓库已开始拣货（picking）的订单禁止修改商品。"""
    conn, woo = env_setup
    jane = _make_user(17, "Jane", "张晴", "admin")

    # 1. 状态为 on-hold（已发货）时拦截
    conn.execute("UPDATE orders SET status='on-hold' WHERE id='1-9001'")
    order_row = conn.execute("SELECT * FROM orders WHERE id='1-9001'").fetchone()
    elig = evaluate_order_modification_eligibility(conn, order_row, jane)
    assert elig["can_modify"] is False
    assert elig["is_owner"] is True
    assert "已发货" in elig["block_reason"]

    # 恢复 processing，但插入发货日志 -> 拦截
    conn.execute("UPDATE orders SET status='processing' WHERE id='1-9001'")
    conn.execute("INSERT INTO shipping_logs(order_id, tracking_number) VALUES ('1-9001', 'TRK123')")
    order_row = conn.execute("SELECT * FROM orders WHERE id='1-9001'").fetchone()
    assert evaluate_order_modification_eligibility(conn, order_row, jane)["can_modify"] is False
    conn.execute("DELETE FROM shipping_logs WHERE order_id='1-9001'")

    # 生成履约单并推进到 picking（拣货中） -> 拦截
    plan_res = plan_order(conn, "1-9001", actor={"type": "system"})
    fid = plan_res["fulfillment_ids"][0]
    transition_fulfillment(conn, fid, "picking", actor={"type": "user", "id": 10, "name": "金毅"})
    conn.commit()
    res, st = modify_order_line_item(
        conn,
        "1-9001",
        {"line_item_id": 701, "replace_qty": 1, "target_product_id": 100, "target_variation_id": 1002},
        user=jane,
        session=woo,
    )
    assert st == 409
    assert "picking" in res["error"]


def test_same_product_flavor_change_locks_price_and_replans_fulfillment(env_setup):
    """同款换口味：强制保持原成交金额不变，并自动将多仓履约预留从旧口味切至新口味且不重复分仓。"""
    conn, woo = env_setup
    jane = _make_user(17, "Jane", "张晴", "admin")

    # 初始分仓：revision 1 分配 1 件 SKU 1
    first_plan = plan_order(conn, "1-9001", actor={"type": "system"})
    assert first_plan["revision"] == 1

    # 提交同款换口味（即使请求体传了 price_mode='sync'，且新口味标价 69.90 高于原价 64.90，仍须锁定 64.90 不变）
    res, st = modify_order_line_item(
        conn,
        "1-9001",
        {
            "line_item_id": 701,
            "expected_original_qty": 1,
            "replace_qty": 1,
            "target_product_id": 100,
            "target_variation_id": 1002,
            "price_mode": "sync",
        },
        user=jane,
        session=woo,
    )
    assert st == 200
    assert res["price_mode"] == "keep"
    assert res["old_total"] == "89.90"
    assert res["new_total"] == "89.90"
    assert res["fulfillment"]["revision"] == 2

    # 核验本地订单行与备注
    updated_order = conn.execute("SELECT * FROM orders WHERE id='1-9001'").fetchone()
    assert float(updated_order["total"]) == pytest.approx(89.90)
    local_items = json.loads(updated_order["line_items"])
    assert len(local_items) == 1
    assert local_items[0]["variation_id"] == 1002
    assert local_items[0]["total"] == "64.90"

    note_row = conn.execute("SELECT * FROM order_notes WHERE order_id='1-9001'").fetchone()
    assert note_row is not None
    assert "张晴" in note_row["note"]
    assert "WATERMELON ICE" in note_row["note"]

    # 核验多仓履约：revision 1 已 superseded，revision 2 仅包含 1 件 SKU 2（绝不重复分配旧行）
    old_ff = conn.execute("SELECT status FROM oms_fulfillments WHERE order_id='1-9001' AND revision=1").fetchone()
    assert old_ff["status"] == "superseded"

    new_ff = conn.execute("SELECT id, status FROM oms_fulfillments WHERE order_id='1-9001' AND revision=2").fetchone()
    assert new_ff["status"] == "ready_to_pick"
    ff_items = conn.execute(
        "SELECT sku_id, allocated_qty FROM oms_fulfillment_items WHERE fulfillment_id=?",
        (new_ff["id"],),
    ).fetchall()
    assert [(r["sku_id"], r["allocated_qty"]) for r in ff_items] == [(2, 1)]

    # 后续仓库开始拣货后，定时同步再次触发 plan_order 时应平滑返回 locked_noop，不抛不一致异常
    transition_fulfillment(conn, new_ff["id"], "picking", actor={"type": "user", "id": 10, "name": "金毅"})
    conn.commit()
    noop_res = plan_order(conn, "1-9001", actor={"type": "system"})
    assert noop_res["action"] == "locked_noop"


def test_partial_quantity_split_preserves_exact_total_without_rounding_drift(env_setup):
    """买 3 支（100.00 PLN）拆分更换 1 支：保留原口味 2 支（66.67）+ 新口味 1 支（33.33），总计严格 100.00 零尾差。"""
    conn, woo = env_setup
    jane = _make_user(17, "Jane", "张晴", "admin")
    plan_order(conn, "1-9002", actor={"type": "system"})

    res, st = modify_order_line_item(
        conn,
        "1-9002",
        {
            "line_item_id": 801,
            "expected_original_qty": 3,
            "replace_qty": 1,
            "target_product_id": 100,
            "target_variation_id": 1002,
            "price_mode": "keep",
        },
        user=jane,
        session=woo,
    )
    assert st == 200
    assert res["new_total"] == "120.00"

    updated_order = conn.execute("SELECT * FROM orders WHERE id='1-9002'").fetchone()
    items = json.loads(updated_order["line_items"])
    assert len(items) == 2
    by_var = {i["variation_id"]: i for i in items}
    assert by_var[1001]["quantity"] == 2
    assert by_var[1001]["total"] == "66.67"
    assert by_var[1002]["quantity"] == 1
    assert by_var[1002]["total"] == "33.33"

    # 履约单应包含 2 件 SKU 1 + 1 件 SKU 2
    ff = conn.execute("SELECT id FROM oms_fulfillments WHERE order_id='1-9002' AND revision=2").fetchone()
    alloc = {
        r["sku_id"]: r["allocated_qty"]
        for r in conn.execute(
            "SELECT sku_id, allocated_qty FROM oms_fulfillment_items WHERE fulfillment_id=?",
            (ff["id"],),
        ).fetchall()
    }
    assert alloc == {1: 2, 2: 1}


def test_cross_product_supports_both_keep_price_and_sync_price_modes(env_setup):
    """跨款更换产品：既支持保持原金额不变（福利赠送），也支持按新产品站点价同步更新订单总额与 COD 代收额。"""
    conn, woo = env_setup
    jane = _make_user(17, "Jane", "张晴", "admin")
    plan_order(conn, "1-9001", actor={"type": "system"})

    # 1. 先在订单 1-9002 上测试跨款拆分 1 支且选择 keep（保持原价）
    res_keep, st_keep = modify_order_line_item(
        conn,
        "1-9002",
        {
            "line_item_id": 801,
            "replace_qty": 1,
            "target_product_id": 200,
            "target_variation_id": 2001,
            "price_mode": "keep",
        },
        user=jane,
        session=woo,
    )
    assert st_keep == 200
    assert res_keep["price_mode"] == "keep"
    assert res_keep["new_total"] == "120.00"

    # 2. 在订单 1-9001 上测试跨款整行替换为 89.90 PLN 的 Blade 30K，且选择 sync（同步新售价）
    # 原商品 64.90 + 运费 25.00 = 89.90 -> 新商品 89.90 + 运费 25.00 = 114.90 PLN
    res_sync, st_sync = modify_order_line_item(
        conn,
        "1-9001",
        {
            "line_item_id": 701,
            "replace_qty": 1,
            "target_product_id": 200,
            "target_variation_id": 2001,
            "price_mode": "sync",
        },
        user=jane,
        session=woo,
    )
    assert st_sync == 200
    assert res_sync["price_mode"] == "sync"
    assert res_sync["old_total"] == "89.90"
    assert res_sync["new_total"] == "114.90"

    # 校验履约财务表中的 COD 代收金额已自动更新为 114.90 PLN
    ff = conn.execute("SELECT id FROM oms_fulfillments WHERE order_id='1-9001' AND revision=2").fetchone()
    fin = conn.execute(
        "SELECT cod_amount, merchandise_amount FROM oms_fulfillment_financials WHERE fulfillment_id=?",
        (ff["id"],),
    ).fetchone()
    assert Decimal(str(fin["cod_amount"])) == Decimal("114.90")
    assert Decimal(str(fin["merchandise_amount"])) == Decimal("89.90")


def test_wildcard_any_flavor_variation_expands_options_and_writes_meta(env_setup):
    """任意口味（Any flavor）共用变体应自动展开父商品口味候选，并在提交时向源站写入对应口味 meta_data。"""
    conn, woo = env_setup
    jane = _make_user(17, "Jane", "张晴", "admin")

    opt_res, opt_st = get_product_options_for_order(
        conn, "1-9001", user=jane, product_id=300, session=woo
    )
    assert opt_st == 200
    labels = [o["flavor_label"] for o in opt_res["options"]]
    assert labels == ["Peach Ice", "Cola Ice"]

    res, st = modify_order_line_item(
        conn,
        "1-9001",
        {
            "line_item_id": 701,
            "replace_qty": 1,
            "target_product_id": 300,
            "target_variation_id": 3001,
            "target_flavor": "Cola Ice",
            "price_mode": "keep",
        },
        user=jane,
        session=woo,
    )
    assert st == 200
    sent_patch = woo.put_calls[-1]["body"]["line_items"]
    created_line = next(x for x in sent_patch if not x.get("id"))
    assert created_line["meta_data"] == [{"key": "pa_smaki", "value": "Cola Ice"}]


def test_compute_amounts_rejects_invalid_quantities():
    """金额拆分函数边界校验。"""
    with pytest.raises(Exception):
        compute_line_item_replacement_amounts(
            orig_qty=2, replace_qty=3, orig_subtotal="50.00", orig_total="50.00", price_mode="keep"
        )
