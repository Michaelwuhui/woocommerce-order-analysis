from __future__ import annotations

import os
import sqlite3

import db_backend as db

import pytest
from flask import Flask

from mail_center_readonly_api import mail_center_readonly_bp


TOKEN = "synthetic-mail-center-order-token-000000000001"


@pytest.fixture(params=["sqlite", "postgres"])
def client(tmp_path, monkeypatch, request):
    database = tmp_path / "orders.sqlite3"
    if request.param == "postgres":
        if not db.is_postgres_backend():
            pytest.skip("isolated PostgreSQL required")
        test_database = os.environ.get("WOO_DB_NAME_OVERRIDE", "")
        assert test_database.startswith("woo_return_loss_test_"), "Use an empty isolated test database"
        conn = db.connect()
        assert conn.execute("SELECT current_database()").fetchone()[0] == test_database
        conn.execute("DELETE FROM orders WHERE id IN ('21001','21002')")
    else:
        monkeypatch.setenv("WOO_DB_BACKEND", "sqlite")
        conn = sqlite3.connect(database)
        conn.execute("""
            CREATE TABLE orders (
                id TEXT PRIMARY KEY, number TEXT, source TEXT, status TEXT,
                date_modified_gmt TEXT, updated_at TEXT, billing TEXT,
                shipping TEXT, line_items TEXT, total REAL
            )
        """)
    conn.execute(
        """
        INSERT INTO orders (id,number,source,status,date_modified_gmt,updated_at,
                            billing,shipping,line_items,total) VALUES (
            '21001', 'VG-21001', 'https://vapego.example.invalid', 'processing',
            '2026-08-15T01:02:03Z', '2026-08-15T09:02:03+08:00',
            '{"email":"private@example.invalid"}', '{"phone":"+48123456789"}',
            '[{"name":"Synthetic Product"}]', 999.0
        )
        """
    )
    conn.execute(
        """
        INSERT INTO orders (id,number,source,status,date_modified_gmt,updated_at,
                            billing,shipping,line_items,total) VALUES (
            '21002', 'VG-21002', 'https://forbidden.example.invalid', 'completed',
            '2026-08-15T02:00:00Z', '2026-08-15T10:00:00+08:00',
            '{}', '{}', '[]', 1.0
        )
        """
    )
    conn.commit()
    conn.close()

    token_file = tmp_path / "order-api-token"
    token_file.write_text(TOKEN, encoding="utf-8")
    if os.name != "nt":
        token_file.chmod(0o600)
    monkeypatch.setenv("MAIL_CENTER_ORDER_API_TOKEN_FILE", str(token_file))
    monkeypatch.setenv("MAIL_CENTER_ORDER_DB_PATH", str(database))
    monkeypatch.setenv("MAIL_CENTER_ORDER_ALLOWED_SITES", "vapego.example.invalid")
    monkeypatch.setenv(
        "MAIL_CENTER_ORDER_PUBLIC_BASE_URL", "https://orders.example.invalid"
    )

    app = Flask(__name__)
    app.register_blueprint(mail_center_readonly_bp)
    with app.test_client() as test_client:
        yield test_client


def _auth() -> dict[str, str]:
    return {"Authorization": f"Bearer {TOKEN}"}


def test_readonly_order_endpoint_returns_only_whitelisted_metadata(client):
    response = client.get("/internal/customer-service/orders/21001", headers=_auth())
    assert response.status_code == 200
    payload = response.get_json()
    assert set(payload) == {
        "fetchedAt",
        "internalUrl",
        "orderId",
        "siteKey",
        "sourceVersion",
        "status",
        "updatedAt",
    }
    assert payload["orderId"] == "VG-21001"
    assert payload["siteKey"] == "vapego.example.invalid"
    assert payload["status"] == "processing"
    assert payload["internalUrl"].endswith("/orders?search=VG-21001")
    combined = response.get_data(as_text=True)
    for forbidden in (
        "private@example.invalid",
        "+48123456789",
        "Synthetic Product",
        "billing",
        "shipping",
        "line_items",
        "total",
    ):
        assert forbidden not in combined


def test_readonly_order_endpoint_fails_closed(client):
    assert client.get("/internal/customer-service/orders/21001").status_code == 403
    assert client.get(
        "/internal/customer-service/orders/21001",
        headers={"Authorization": "Bearer wrong-token-value-that-is-long-enough"},
    ).status_code == 403
    assert client.get(
        "/internal/customer-service/orders/21002", headers=_auth()
    ).status_code == 403
    assert client.get(
        "/internal/customer-service/orders/99999", headers=_auth()
    ).status_code == 404
    assert client.get(
        "/internal/customer-service/orders/bad%20id", headers=_auth()
    ).status_code == 400
    assert client.post(
        "/internal/customer-service/orders/21001", headers=_auth()
    ).status_code == 405


def test_readonly_order_endpoint_rejects_woocommerce_token_reference(
    client, tmp_path, monkeypatch
):
    token_file = tmp_path / "bad-token"
    token_file.write_text("ck_synthetic_not_a_real_key_but_forbidden_shape", encoding="utf-8")
    if os.name != "nt":
        token_file.chmod(0o600)
    monkeypatch.setenv("MAIL_CENTER_ORDER_API_TOKEN_FILE", str(token_file))
    assert client.get(
        "/internal/customer-service/orders/21001",
        headers={"Authorization": "Bearer ck_synthetic_not_a_real_key_but_forbidden_shape"},
    ).status_code == 403


@pytest.mark.skipif(os.name == "nt", reason="POSIX permission and symlink semantics only")
def test_readonly_order_endpoint_rejects_unsafe_token_files(client, tmp_path, monkeypatch):
    broad = tmp_path / "group-readable-token"
    broad.write_text(TOKEN, encoding="utf-8")
    broad.chmod(0o640)
    monkeypatch.setenv("MAIL_CENTER_ORDER_API_TOKEN_FILE", str(broad))
    assert client.get(
        "/internal/customer-service/orders/21001", headers=_auth()
    ).status_code == 403

    private = tmp_path / "private-token"
    private.write_text(TOKEN, encoding="utf-8")
    private.chmod(0o600)
    link = tmp_path / "linked-token"
    link.symlink_to(private)
    monkeypatch.setenv("MAIL_CENTER_ORDER_API_TOKEN_FILE", str(link))
    assert client.get(
        "/internal/customer-service/orders/21001", headers=_auth()
    ).status_code == 403
