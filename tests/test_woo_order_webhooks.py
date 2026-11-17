"""WooCommerce 实时订单 Webhook 同步模块（含验签、Ping 握手、自愈部署、权威回读落库与权限控制）单元测试。"""

from __future__ import annotations

import json
from typing import Any
from urllib.parse import urlsplit

from flask import Flask
from flask_login import LoginManager, UserMixin, login_user
import pytest
import requests

from sync_utils import WC_ORDER_FIELDS
import woo_order_webhooks as wh


class DummyUser(UserMixin):
    """用于测试权限隔离的轻量用户对象。"""

    def __init__(
        self,
        uid: int,
        name: str,
        *,
        role: str = "user",
        manage_users: bool = False,
        own_sync: bool = False,
    ):
        self.id = str(uid)
        self.name = name
        self.role = role
        self._manage_users = manage_users
        self._own_sync = own_sync

    def is_admin(self) -> bool:
        return self.role == "admin"

    def can_manage_users(self) -> bool:
        return self._manage_users

    def can_manage_own_site_sync(self) -> bool:
        return self._own_sync


class FakeResponse:
    """模拟 requests.Response 对象以测试 REST API 客户端与裸域跳转。"""

    def __init__(
        self,
        status_code: int = 200,
        json_data: Any = None,
        headers: dict[str, str] | None = None,
        text: str = "",
    ):
        self.status_code = status_code
        self._json_data = json_data
        self.headers = {"Content-Type": "application/json", **(headers or {})}
        self.text = text or (json.dumps(json_data) if json_data is not None else "")

    def json(self):
        if self._json_data is None:
            raise ValueError("empty json")
        return self._json_data


def _init_minimal_schema(conn) -> None:
    """创建隔离 SQLite 测试所需的最小核心表结构。"""
    conn.execute(
        """
        CREATE TABLE sites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            url TEXT NOT NULL UNIQUE,
            consumer_key TEXT NOT NULL,
            consumer_secret TEXT NOT NULL,
            manager TEXT DEFAULT '',
            country TEXT DEFAULT 'PL',
            api_status TEXT DEFAULT 'ok'
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            name TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'user'
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )
    order_cols = ", ".join(
        "id TEXT PRIMARY KEY" if f == "id" else f"{f} TEXT"
        for f in WC_ORDER_FIELDS
    )
    conn.execute(
        f"CREATE TABLE orders ({order_cols}, woo_id INTEGER, updated_at TEXT)"
    )
    conn.execute(
        """
        CREATE TABLE order_notes (
            wc_note_id INTEGER NOT NULL,
            order_id TEXT NOT NULL,
            note TEXT,
            date_created TEXT,
            customer_note INTEGER DEFAULT 0,
            author TEXT,
            added_by_user INTEGER DEFAULT 0,
            PRIMARY KEY (order_id, wc_note_id)
        )
        """
    )
    wh.init_woo_order_webhook_tables(conn)


@pytest.fixture
def webhook_env(tmp_path, monkeypatch):
    """提供绑定临时 SQLite 数据库的 Flask 应用与预置站点数据。"""
    db_file = str(tmp_path / "test_webhooks.db")
    monkeypatch.setenv("WOO_DB_BACKEND", "sqlite")
    monkeypatch.setenv("WOO_SQLITE_PATH", db_file)
    monkeypatch.setenv("WOO_WEBHOOK_PUBLIC_BASE_URL", "https://cfworks.vip")

    conn = wh.get_db_conn(db_file)
    try:
        _init_minimal_schema(conn)
        conn.execute(
            "INSERT INTO sites (id, url, consumer_key, consumer_secret, manager, country, api_status) "
            "VALUES (1, 'https://shop1.example', 'ck_1', 'cs_1', '张三', 'PL', 'ok')"
        )
        conn.execute(
            "INSERT INTO sites (id, url, consumer_key, consumer_secret, manager, country, api_status) "
            "VALUES (2, 'https://shop2.example', 'ck_2', 'cs_2', '李四', 'AU', 'ok')"
        )
        conn.execute(
            "INSERT INTO sites (id, url, consumer_key, consumer_secret, manager, country, api_status) "
            "VALUES (3, 'https://retired.example', '', '', '张三', 'PL', 'archived')"
        )
        conn.execute(
            "INSERT INTO users (id, username, name, role) VALUES (1, 'admin', '超级管理员', 'admin')"
        )
        conn.execute(
            "INSERT INTO users (id, username, name, role) VALUES (2, 'zhangsan', '张三', 'user')"
        )
        conn.execute(
            "INSERT INTO users (id, username, name, role) VALUES (3, 'guest', '路人', 'viewer')"
        )
        conn.commit()
    finally:
        conn.close()

    app = Flask(__name__)
    app.secret_key = "test-secret"
    login_manager = LoginManager()
    login_manager.init_app(app)

    users_map = {
        "1": DummyUser(1, "超级管理员", role="admin", manage_users=True, own_sync=True),
        "2": DummyUser(2, "张三", role="user", manage_users=False, own_sync=True),
        "3": DummyUser(3, "路人", role="viewer", manage_users=False, own_sync=False),
    }

    @login_manager.user_loader
    def load_user(uid):
        return users_map.get(str(uid))

    @app.route("/_test_login/<uid>", methods=["POST"])
    def _test_login(uid):
        login_user(users_map[str(uid)])
        return "ok"

    app.register_blueprint(wh.woo_order_webhook_bp)
    return {"app": app, "db_file": db_file}


def test_ping_handshake_and_signature_verification(webhook_env, monkeypatch):
    """验证 Ping 连通性握手、HMAC-SHA256 验签、密钥平滑轮换、重放去重与草稿单过滤。"""
    app = webhook_env["app"]
    db_file = webhook_env["db_file"]
    client = app.test_client()

    conn = wh.get_db_conn(db_file)
    secret, _ = wh._ensure_local_site_secrets(
        conn, 1, wh.build_delivery_url(1)
    )
    conn.close()

    # 1. Ping 握手（x-www-form-urlencoded webhook_id=88）
    ping_resp = client.post(
        "/api/webhooks/woocommerce/orders/1",
        data="webhook_id=88",
        content_type="application/x-www-form-urlencoded",
        headers={"User-Agent": "WooCommerce/9.3 Hookshot (WordPress/6.7)"},
    )
    assert ping_resp.status_code == 200
    assert ping_resp.get_json()["ping"] is True

    # 2. 错误签名被拒绝（HTTP 401）
    order_body = json.dumps({"id": 5001, "status": "processing", "total": "99.00"}).encode("utf-8")
    bad_resp = client.post(
        "/api/webhooks/woocommerce/orders/1",
        data=order_body,
        content_type="application/json",
        headers={
            "X-WC-Webhook-Topic": "order.created",
            "X-WC-Webhook-Signature": "invalid-sig",
        },
    )
    assert bad_resp.status_code == 401

    # 3. 正确签名秒级入队 + 重放幂等去重
    dispatched_calls = []
    monkeypatch.setattr(
        wh,
        "dispatch_webhook_order_fetch",
        lambda s_id, o_id, **kw: dispatched_calls.append((s_id, o_id)) or "dispatched",
    )
    valid_sig = wh.compute_wc_signature(order_body, secret)
    ok_resp = client.post(
        "/api/webhooks/woocommerce/orders/1",
        data=order_body,
        content_type="application/json",
        headers={
            "X-WC-Webhook-Topic": "order.created",
            "X-WC-Webhook-Event": "created",
            "X-WC-Webhook-Delivery-ID": "deliv-001",
            "X-WC-Webhook-Signature": valid_sig,
        },
    )
    assert ok_resp.status_code == 200
    assert ok_resp.get_json()["queued"] is True
    assert dispatched_calls == [(1, 5001)]

    dup_resp = client.post(
        "/api/webhooks/woocommerce/orders/1",
        data=order_body,
        content_type="application/json",
        headers={
            "X-WC-Webhook-Topic": "order.created",
            "X-WC-Webhook-Event": "created",
            "X-WC-Webhook-Delivery-ID": "deliv-001",
            "X-WC-Webhook-Signature": valid_sig,
        },
    )
    assert dup_resp.status_code == 200
    assert dup_resp.get_json()["duplicate"] is True
    assert len(dispatched_calls) == 1

    # 4. 密钥轮换后旧在途请求（使用 previous_secret 签名）依然可验证通过
    conn = wh.get_db_conn(db_file)
    new_secret, _ = wh._ensure_local_site_secrets(
        conn, 1, wh.build_delivery_url(1), rotate_secret=True
    )
    conn.close()
    assert new_secret != secret

    upd_body = json.dumps({"id": 5002, "status": "completed"}).encode("utf-8")
    old_sig = wh.compute_wc_signature(upd_body, secret)
    rot_resp = client.post(
        "/api/webhooks/woocommerce/orders/1",
        data=upd_body,
        content_type="application/json",
        headers={
            "X-WC-Webhook-Topic": "order.updated",
            "X-WC-Webhook-Delivery-ID": "deliv-002",
            "X-WC-Webhook-Signature": old_sig,
        },
    )
    assert rot_resp.status_code == 200
    assert rot_resp.get_json()["queued"] is True
    assert dispatched_calls[-1] == (1, 5002)

    # 5. 草稿单（checkout-draft）不触发回读队列
    draft_body = json.dumps({"id": 5003, "status": "checkout-draft"}).encode("utf-8")
    draft_sig = wh.compute_wc_signature(draft_body, new_secret)
    draft_resp = client.post(
        "/api/webhooks/woocommerce/orders/1",
        data=draft_body,
        content_type="application/json",
        headers={
            "X-WC-Webhook-Topic": "order.created",
            "X-WC-Webhook-Delivery-ID": "deliv-003",
            "X-WC-Webhook-Signature": draft_sig,
        },
    )
    assert draft_resp.status_code == 200
    assert draft_resp.get_json()["ignored"] == "draft_order"
    assert len(dispatched_calls) == 2


def test_archived_site_and_global_switch_return_200_without_enqueue(webhook_env, monkeypatch):
    """验证归档站点或全局暂停时返回 HTTP 200（防源站自动禁用）且绝不入队修改数据。"""
    app = webhook_env["app"]
    db_file = webhook_env["db_file"]
    client = app.test_client()

    conn = wh.get_db_conn(db_file)
    s1, _ = wh._ensure_local_site_secrets(conn, 1, wh.build_delivery_url(1))
    s3, _ = wh._ensure_local_site_secrets(conn, 3, wh.build_delivery_url(3))
    wh.set_webhook_sync_enabled(conn, False)
    conn.close()

    enqueued = []
    monkeypatch.setattr(
        wh, "dispatch_webhook_order_fetch", lambda *a, **kw: enqueued.append(a) or "dispatched"
    )

    b1 = json.dumps({"id": 6001, "status": "processing"}).encode("utf-8")
    r1 = client.post(
        "/api/webhooks/woocommerce/orders/1",
        data=b1,
        content_type="application/json",
        headers={
            "X-WC-Webhook-Topic": "order.created",
            "X-WC-Webhook-Signature": wh.compute_wc_signature(b1, s1),
        },
    )
    assert r1.status_code == 200
    assert r1.get_json()["ignored"] == "webhook_sync_disabled"

    conn = wh.get_db_conn(db_file)
    wh.set_webhook_sync_enabled(conn, True)
    conn.close()

    b3 = json.dumps({"id": 6002, "status": "processing"}).encode("utf-8")
    r3 = client.post(
        "/api/webhooks/woocommerce/orders/3",
        data=b3,
        content_type="application/json",
        headers={
            "X-WC-Webhook-Topic": "order.created",
            "X-WC-Webhook-Signature": wh.compute_wc_signature(b3, s3),
        },
    )
    assert r3.status_code == 200
    assert r3.get_json()["ignored"] == "site_archived"
    assert enqueued == []


def test_provision_site_webhooks_creates_heals_dedupes_and_handles_www_redirect(webhook_env):
    """验证受管 Webhook 的创建、停用自愈、重复项清理以及裸域到 www 的 301 保方法跳转。"""
    db_file = webhook_env["db_file"]

    remote_store: list[dict[str, Any]] = [
        # 模拟源站上已有一个被自动停用的 order.created 以及一个多余重复项
        {
            "id": 11,
            "name": "[CFWorks OMS] 订单创建实时同步",
            "topic": "order.created",
            "delivery_url": "https://cfworks.vip/api/webhooks/woocommerce/orders/1",
            "status": "disabled",
        },
        {
            "id": 12,
            "name": "[CFWorks OMS] 订单创建实时同步",
            "topic": "order.created",
            "delivery_url": "https://cfworks.vip/api/webhooks/woocommerce/orders/1",
            "status": "disabled",
        },
    ]
    recorded_calls: list[tuple[str, str, Any]] = []

    class MockSession:
        def request(self, *, method, url, auth, params=None, json=None, headers=None, timeout=None, allow_redirects=False):
            parsed = urlsplit(url)
            # 模拟裸域 shop1.example 强制 301 到 www.shop1.example
            if parsed.hostname == "shop1.example":
                new_url = parsed._replace(netloc="www.shop1.example").geturl()
                return FakeResponse(301, headers={"Location": new_url})
            assert auth == ("ck_1", "cs_1")
            recorded_calls.append((method, parsed.path, json))
            if method == "GET" and parsed.path.endswith("/webhooks"):
                return FakeResponse(200, list(remote_store))
            if method == "DELETE":
                del_id = int(parsed.path.rsplit("/", 1)[-1])
                remote_store[:] = [w for w in remote_store if w["id"] != del_id]
                return FakeResponse(200, {"id": del_id})
            if method == "PUT":
                put_id = int(parsed.path.rsplit("/", 1)[-1])
                for w in remote_store:
                    if w["id"] == put_id:
                        w.update(json or {})
                        return FakeResponse(200, dict(w))
            if method == "POST" and parsed.path.endswith("/webhooks"):
                new_item = {"id": 99, **(json or {})}
                remote_store.append(new_item)
                return FakeResponse(201, new_item)
            raise AssertionError(f"unexpected call: {method} {url}")

        def close(self):
            pass

    res = wh.provision_site_webhooks(1, force_verify=True, session=MockSession(), db_path=db_file)
    assert res["ok"] is True
    assert res["topics"]["order.created"]["action"] == "updated"
    assert res["topics"]["order.created"]["remote_id"] == 11
    assert res["topics"]["order.updated"]["action"] == "created"
    assert res["topics"]["order.updated"]["remote_id"] == 99
    # 重复项 id=12 已被自动删除，最终只剩 2 条 active 受管 Webhook
    assert sorted(w["id"] for w in remote_store) == [11, 99]
    assert all(w["status"] == "active" for w in remote_store)

    # 紧接着非强制巡检（6 小时缓存期内）应命中健康缓存，不再向源站发网络请求
    calls_before = len(recorded_calls)
    cached_res = wh.provision_site_webhooks(1, force_verify=False, session=MockSession(), db_path=db_file)
    assert cached_res["ok"] is True
    assert cached_res.get("cached") is True
    assert len(recorded_calls) == calls_before


def test_single_order_fetch_and_serial_write_triggers_post_commit(webhook_env, monkeypatch):
    """验证单笔权威回读、订单与备注落库、投递状态结算以及下游履约/通知联动触发。"""
    db_file = webhook_env["db_file"]
    conn = wh.get_db_conn(db_file)
    conn.execute(
        """
        INSERT INTO woo_webhook_deliveries
            (delivery_key, site_id, woo_order_id, topic, payload_hash, status)
        VALUES ('1:order.created:d900:abc', 1, 9001, 'order.created', 'abc', 'fetching')
        """
    )
    conn.commit()
    conn.close()

    class OrderMockSession:
        def request(self, *, method, url, auth, params=None, json=None, headers=None, timeout=None, allow_redirects=False):
            if url.endswith("/orders/9001"):
                return FakeResponse(
                    200,
                    {
                        "id": 9001,
                        "number": "9001",
                        "status": "processing",
                        "currency": "PLN",
                        "total": "159.00",
                        "date_created": "2026-11-17T10:00:00",
                        "date_modified": "2026-11-17T10:00:02",
                        "billing": {"first_name": "Jan", "phone": "+48123456789"},
                        "shipping": {"first_name": "Jan"},
                        "line_items": [{"id": 1, "name": "Vape Berry", "quantity": 2, "total": "159.00"}],
                    },
                )
            if url.endswith("/orders/9001/notes"):
                return FakeResponse(
                    200,
                    [
                        {
                            "id": 301,
                            "note": "Przelewy24 付款已确认",
                            "date_created": "2026-11-17T10:00:02",
                            "customer_note": False,
                            "author": "WooCommerce",
                            "added_by_user": False,
                        }
                    ],
                )
            raise AssertionError(url)

        def close(self):
            pass

    post_commit_calls = []
    monkeypatch.setattr(
        wh,
        "run_post_commit_sync_actions",
        lambda cands, strict=False: post_commit_calls.append(list(cands)),
    )

    fetched = wh.fetch_single_order_from_woo(
        "https://shop1.example", "ck_1", "cs_1", 1, 9001, session=OrderMockSession()
    )
    assert fetched["not_found"] is False
    write_out = wh.apply_webhook_order_write(
        1, 9001, fetched["order"], fetched["notes"], db_path=db_file
    )
    assert write_out["written"] == 1
    assert write_out["inserted"] == 1
    assert write_out["order_oid"] == "1-9001"
    assert len(post_commit_calls) == 1
    assert post_commit_calls[0][0]["order_id"] == "1-9001"

    conn = wh.get_db_conn(db_file)
    saved_order = conn.execute("SELECT id, status, total, source FROM orders WHERE id='1-9001'").fetchone()
    saved_note = conn.execute("SELECT note FROM order_notes WHERE order_id='1-9001' AND wc_note_id=301").fetchone()
    deliv = conn.execute("SELECT status, order_oid, order_status, latency_ms FROM woo_webhook_deliveries WHERE woo_order_id=9001").fetchone()
    conn.close()

    assert saved_order["status"] == "processing"
    assert saved_order["source"] == "https://shop1.example"
    assert saved_note["note"] == "Przelewy24 付款已确认"
    assert deliv["status"] == "synced"
    assert deliv["order_oid"] == "1-9001"
    assert deliv["latency_ms"] is not None and deliv["latency_ms"] >= 0


def test_dashboard_permission_scoping(webhook_env, monkeypatch):
    """验证管理面板接口的权限控制（只读用户 403、站点负责人仅限本人站点、超管可管理全站）。"""
    app = webhook_env["app"]
    client = app.test_client()

    # 1. 无权限用户被拒绝
    client.post("/_test_login/3")
    assert client.get("/api/settings/webhooks/status").status_code == 403
    assert client.post("/api/settings/webhooks/toggle", json={"enabled": False}).status_code == 403

    # 2. 站点负责人（张三，负责站点 1）只能看到站点 1，不能修改他人的站点 2 或执行全站批量操作
    client.post("/_test_login/2")
    st_resp = client.get("/api/settings/webhooks/status")
    assert st_resp.status_code == 200
    st_data = st_resp.get_json()
    assert [s["site_id"] for s in st_data["sites"]] == [1]
    assert client.post("/api/settings/webhooks/provision", json={"site_id": 2}).status_code == 403
    assert client.post("/api/settings/webhooks/provision", json={}).status_code == 403

    # 3. 超级管理员可查看全部活跃站点（1 和 2，不含归档站点 3）并切换总开关
    client.post("/_test_login/1")
    adm_resp = client.get("/api/settings/webhooks/status")
    assert adm_resp.status_code == 200
    adm_data = adm_resp.get_json()
    assert [s["site_id"] for s in adm_data["sites"]] == [1, 2]
    tog = client.post("/api/settings/webhooks/toggle", json={"enabled": False})
    assert tog.status_code == 200 and tog.get_json()["enabled"] is False
