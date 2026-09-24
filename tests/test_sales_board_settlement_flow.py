"""Synthetic August settlement, commission and Excel regression."""

import importlib
import io
import json
import re
import shutil
import sqlite3
import subprocess

from openpyxl import load_workbook
import pytest


@pytest.fixture
def board(tmp_path, monkeypatch):
    database = tmp_path / "orders.sqlite3"
    conn = sqlite3.connect(database)
    conn.executescript("""
        CREATE TABLE orders (
            id TEXT PRIMARY KEY, source TEXT, date_created TEXT, status TEXT,
            payment_method TEXT, currency TEXT, total REAL, shipping_total REAL,
            line_items TEXT, warehouse_id INTEGER, is_undelivered INTEGER DEFAULT 0,
            is_problem_return INTEGER DEFAULT 0, shipping_loss_amount REAL DEFAULT 0,
            product_loss_amount REAL DEFAULT 0
        );
    """)
    conn.close()

    monkeypatch.setenv("WOO_DB_BACKEND", "sqlite")
    monkeypatch.setenv("WOO_SQLITE_PATH", str(database))
    app_module = importlib.import_module("app")
    monkeypatch.setattr(app_module, "DB_FILE", str(database))
    monkeypatch.setitem(app_module.app.config, "TESTING", True)

    conn = sqlite3.connect(database)
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY, username TEXT UNIQUE, name TEXT, role TEXT,
            can_view_sales_board INTEGER DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS sites (
            id INTEGER PRIMARY KEY, url TEXT, manager TEXT, country TEXT,
            cod_on_hold_is_shipped INTEGER DEFAULT 1,
            consumer_key TEXT NOT NULL, consumer_secret TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS partners (
            id INTEGER PRIMARY KEY, name TEXT, currency TEXT
        );
        CREATE TABLE IF NOT EXISTS partner_receipts (
            id INTEGER PRIMARY KEY, partner_id INTEGER, receipt_date TEXT,
            amount_pln REAL, exchange_rate_cny REAL, amount_cny REAL
        );
        CREATE TABLE IF NOT EXISTS exchange_rates (
            year_month TEXT, currency TEXT, rate_to_cny REAL,
            UNIQUE(year_month, currency)
        );
        CREATE TABLE IF NOT EXISTS sales_board_exchange_rates (
            year_month TEXT, currency TEXT, rate_to_cny REAL, updated_by TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(year_month, currency)
        );
        CREATE TABLE IF NOT EXISTS sales_board_settlement_rates (
            year_month TEXT, currency TEXT, rate_to_cny REAL, updated_by TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY(year_month, currency)
        );
        CREATE TABLE IF NOT EXISTS sales_targets (
            year_month TEXT, manager TEXT, monthly_target REAL,
            weekly_targets TEXT, base_salary REAL, commission_rate REAL,
            notes TEXT
        );
        CREATE TABLE IF NOT EXISTS no_commission_brands (brand_name TEXT);
        CREATE TABLE IF NOT EXISTS brands (
            id INTEGER PRIMARY KEY, name TEXT, aliases TEXT
        );
        CREATE TABLE IF NOT EXISTS series (
            id INTEGER PRIMARY KEY, brand_id INTEGER, name TEXT
        );
        CREATE TABLE IF NOT EXISTS sales_board_profit_settings (
            year_month TEXT, profit_mode TEXT, profit_percentage REAL,
            country_percentages TEXT
        );
        CREATE TABLE IF NOT EXISTS sales_groups (
            id INTEGER PRIMARY KEY, name TEXT, leader_manager TEXT,
            bonus_rate REAL
        );
        CREATE TABLE IF NOT EXISTS sales_group_members (
            group_id INTEGER, manager TEXT
        );
    """)
    conn.execute(
        "INSERT OR IGNORE INTO users (username, name, role) VALUES ('admin', '管理员', 'admin')"
    )
    conn.execute("UPDATE users SET can_view_sales_board = 1 WHERE username = 'admin'")
    conn.execute(
        "INSERT INTO sites (url, manager, country, cod_on_hold_is_shipped, "
        "consumer_key, consumer_secret) "
        "VALUES ('https://example.invalid', '员工甲', 'PL', 1, 'test-key', 'test-secret')"
    )
    conn.execute(
        "INSERT INTO partners (name, currency) VALUES ('测试合伙人', 'PLN')"
    )
    partner_id = conn.execute("SELECT id FROM partners WHERE name='测试合伙人'").fetchone()[0]
    conn.execute(
        "INSERT INTO partner_receipts (partner_id, receipt_date, amount_pln, amount_cny) "
        "VALUES (?, '2026-08-15', 100, 177.2)", (partner_id,)
    )
    conn.execute(
        "INSERT INTO exchange_rates (year_month, currency, rate_to_cny) "
        "VALUES ('2026-08', 'PLN', 1.82)"
    )
    conn.execute(
        "INSERT INTO sales_targets (year_month, manager, monthly_target, weekly_targets, "
        "base_salary, commission_rate, notes) "
        "VALUES ('2026-08', '员工甲', 0, '{}', 7000, 0.05, '')"
    )
    conn.execute(
        "INSERT INTO orders (id, source, date_created, status, payment_method, "
        "currency, total, shipping_total, line_items) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            "synthetic-august-order", "https://example.invalid", "2026-08-12",
            "completed", "cod", "PLN", 100, 0,
            json.dumps([{"name": "普通商品", "quantity": 2, "total": "100"}]),
        ),
    )
    conn.commit()
    admin_id = conn.execute("SELECT id FROM users WHERE username='admin'").fetchone()[0]
    conn.close()

    client = app_module.app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = str(admin_id)
        session["_fresh"] = True
    return app_module, client, database


def test_august_actual_settlement_changes_commission_and_export(board):
    app_module, client, database = board
    url = "/api/sales-board/exchange-rates"
    original = client.get(url + "?month=2026-08").get_json()["rates"]
    assert original[0]["source"] == "receipt"
    assert original[0]["in_use"] == pytest.approx(1.772)

    saved = client.post(url, json={
        "month": "2026-08",
        "rates": [{"currency": "PLN", "rate": 1.75, "settlement_rate": 1.69}],
    })
    assert saved.status_code == 200, saved.get_json()
    active = client.get(url + "?month=2026-08").get_json()["rates"]
    assert active[0]["source"] == "settlement"
    assert active[0]["in_use"] == pytest.approx(1.69)
    assert active[0]["receipt_rate"] == pytest.approx(1.772)
    assert active[0]["override_rate"] == pytest.approx(1.75)

    page = client.get("/sales-board?month=2026-08")
    assert page.status_code == 200
    page_html = page.get_data(as_text=True)
    assert "实际结算汇率" in page_html
    scripts = re.findall(r"<script(?:\s[^>]*)?>(.*?)</script>", page_html, re.S)
    board_script = next(script for script in scripts if "function openRateSettings" in script)
    if node := shutil.which("node"):
        checked = subprocess.run([node, "--check"], input=board_script, text=True,
                                 capture_output=True, check=False)
        assert checked.returncode == 0, checked.stderr

    with app_module.app.test_request_context("/sales-board"):
        data = app_module._compute_sales_board_data("2026-08")
    employee = data["board_data"][0]
    assert employee["commission_base_cny"] == 169
    assert employee["commission"] == 8.45
    assert data["rates_in_use"]["PLN"] == {"rate": 1.69, "source": "settlement"}

    workbook = load_workbook(io.BytesIO(app_module._generate_sales_board_excel(data).getvalue()))
    rule_lines = [cell.value for row in workbook["规则说明"] for cell in row if cell.value]
    assert any("PLN" in line and "实际结算" in line for line in rule_lines)
    assert any("实际结算汇率 → 当月回款加权汇率" in line for line in rule_lines)

    # An old client updating only the fallback must not clear the settlement rate.
    assert client.post(url, json={"month": "2026-08", "rates": [
        {"currency": "PLN", "rate": 1.74}
    ]}).status_code == 200
    assert client.get(url + "?month=2026-08").get_json()["rates"][0]["source"] == "settlement"

    invalid = client.post(url, json={"month": "2026-08", "rates": [
        {"currency": "AUD", "settlement_rate": 4.6},
        {"currency": "PLN", "settlement_rate": "NaN"},
    ]})
    assert invalid.status_code == 400
    conn = sqlite3.connect(database)
    assert conn.execute("SELECT COUNT(*) FROM sales_board_settlement_rates").fetchone()[0] == 1
    conn.close()

    assert client.post(url, json={"month": "2026-08", "rates": [
        {"currency": "PLN", "settlement_rate": None}
    ]}).status_code == 200
    restored = client.get(url + "?month=2026-08").get_json()["rates"]
    assert restored[0]["source"] == "receipt"
    assert restored[0]["override_rate"] == pytest.approx(1.74)
    with app_module.app.test_request_context("/sales-board"):
        assert app_module._compute_sales_board_data("2026-08")["board_data"][0]["commission"] == 8.86
