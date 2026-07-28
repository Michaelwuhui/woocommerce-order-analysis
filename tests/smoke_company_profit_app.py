"""Isolated full-app smoke test for company-profit routes and strict permissions."""

import os
import sqlite3
import sys
import tempfile


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)


def seed_bootstrap_database(path):
    conn = sqlite3.connect(path)
    conn.executescript(
        """
        CREATE TABLE users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            name TEXT,
            role TEXT DEFAULT 'user',
            can_ship INTEGER DEFAULT 0,
            can_view_report INTEGER DEFAULT 0,
            can_view_sales_board INTEGER DEFAULT 0,
            can_manage_users INTEGER DEFAULT 0,
            can_view_reconciliation INTEGER DEFAULT 0,
            can_edit_reconciliation INTEGER DEFAULT 0,
            can_manage_products INTEGER DEFAULT 0,
            can_manage_own_products INTEGER DEFAULT 0,
            can_view_costs INTEGER DEFAULT 0,
            can_edit_costs INTEGER DEFAULT 0,
            can_view_own_sales_board INTEGER DEFAULT 0,
            can_view_shipping INTEGER DEFAULT 0,
            can_manage_blocklist INTEGER DEFAULT 0,
            can_view_inventory INTEGER DEFAULT 0,
            can_manage_inventory INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        INSERT INTO users (username, password_hash, name, role)
        VALUES ('michael', 'test-only', '吴辉', 'viewer');

        CREATE TABLE orders (
            id TEXT PRIMARY KEY,
            source TEXT,
            status TEXT,
            date_created TEXT,
            total REAL DEFAULT 0,
            shipping_total REAL DEFAULT 0,
            currency TEXT DEFAULT 'CNY',
            line_items TEXT DEFAULT '[]',
            payment_method TEXT DEFAULT 'cod'
        );
        """
    )
    conn.commit()
    conn.close()


with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tempdir:
    original_cwd = os.getcwd()
    try:
        os.chdir(tempdir)
        seed_bootstrap_database("woocommerce_orders.db")

        from app import app  # noqa: E402

        conn = sqlite3.connect("woocommerce_orders.db")
        conn.row_factory = sqlite3.Row
        michael = conn.execute(
            """
            SELECT id, can_view_company_profit, can_edit_company_profit
            FROM users WHERE username = 'michael'
            """
        ).fetchone()
        admin = conn.execute(
            """
            SELECT id, can_view_company_profit, can_edit_company_profit
            FROM users WHERE username = 'admin'
            """
        ).fetchone()
        conn.close()

        assert michael and tuple(michael)[1:] == (0, 0)
        assert admin and tuple(admin)[1:] == (1, 1)

        client = app.test_client()
        with client.session_transaction() as session:
            session["_user_id"] = str(admin["id"])
            session["_fresh"] = True
        page = client.get("/company-profit?month=2026-07")
        assert page.status_code == 200, page.status_code
        assert "公司盈利".encode("utf-8") in page.data

        with client.session_transaction() as session:
            session["_user_id"] = str(michael["id"])
            session["_fresh"] = True
        blocked = client.get("/company-profit?month=2026-07")
        assert blocked.status_code == 403, blocked.status_code

        with client.session_transaction() as session:
            session["_user_id"] = str(admin["id"])
            session["_fresh"] = True
        users = client.get("/api/users")
        assert users.status_code == 200, users.status_code
        michael_json = next(
            user for user in users.get_json() if user["username"] == "michael"
        )
        assert michael_json["can_view_company_profit"] == 0
        assert michael_json["can_edit_company_profit"] == 0

        required_routes = {
            "/company-profit",
            "/api/company-profit/summary",
            "/api/company-profit/settings",
            "/api/company-profit/market-rules",
            "/api/company-profit/expenses",
        }
        rules = {rule.rule for rule in app.url_map.iter_rules()}
        assert not (required_routes - rules)
        app.jinja_env.get_template("company_profit.html")
        app.jinja_env.get_template("company_profit_denied.html")
        print(
            "company_profit_app_smoke=ok "
            f"routes={len(required_routes)} admin=200 michael=403"
        )
    finally:
        os.chdir(original_cwd)
