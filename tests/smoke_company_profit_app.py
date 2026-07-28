"""Full-app smoke test for retired Web routes and read-only offline export."""

import json
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import tempfile


ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


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
    temp_root = Path(tempdir)
    database = temp_root / "woocommerce_orders.db"
    seed_bootstrap_database(database)
    original_cwd = os.getcwd()
    try:
        os.chdir(temp_root)
        from app import app  # noqa: E402
        import company_profit  # noqa: E402

        client = app.test_client()
        for route in (
            "/company-profit",
            "/api/company-profit/summary",
            "/api/company-profit/settings",
            "/api/company-profit/market-rules",
            "/api/company-profit/forecast-scenario",
            "/api/company-profit/expenses",
        ):
            response = client.get(route)
            assert response.status_code == 404, (route, response.status_code)

        rules = {rule.rule for rule in app.url_map.iter_rules()}
        assert not any(
            route == "/company-profit" or route.startswith("/api/company-profit")
            for route in rules
        )
        assert "company_profit.html" not in app.jinja_env.list_templates()
        assert "company_profit_denied.html" not in app.jinja_env.list_templates()

        assert not hasattr(company_profit, "create_company_profit_blueprint")

        source = sqlite3.connect(database)
        source.row_factory = sqlite3.Row
        admin = source.execute(
            "SELECT id FROM users WHERE username = 'admin'"
        ).fetchone()
        michael = source.execute(
            "SELECT id FROM users WHERE username = 'michael'"
        ).fetchone()
        source.close()
        assert admin and michael
        with client.session_transaction() as session:
            session["_user_id"] = str(admin["id"])
            session["_fresh"] = True
        users_response = client.get("/api/users")
        assert users_response.status_code == 200
        michael_payload = next(
            user
            for user in users_response.get_json()
            if user["username"] == "michael"
        )
        assert "can_view_company_profit" not in michael_payload
        assert "can_edit_company_profit" not in michael_payload
        update_response = client.put(
            f"/api/users/{michael['id']}",
            json={
                "name": "吴辉",
                "role": "viewer",
                "can_view_report": 1,
            },
        )
        assert update_response.status_code == 200
        assert update_response.get_json()["success"] is True

        snapshot_output = temp_root / "offline-snapshot.json"
        command = [
            sys.executable,
            str(ROOT / "offline_company_profit_snapshot.py"),
            "--month",
            "2026-07",
            "--database",
            str(database),
            "--output",
            str(snapshot_output),
        ]
        completed = subprocess.run(
            command,
            cwd=ROOT,
            check=True,
            capture_output=True,
            text=True,
        )
        result = json.loads(completed.stdout.strip().splitlines()[-1])
        assert result["success"] is True
        snapshot = json.loads(snapshot_output.read_text(encoding="utf-8"))
        assert snapshot["schema_version"] == 1
        assert snapshot["month"] == "2026-07"
        assert snapshot["source"]["mode"].startswith("sqlite_readonly")
        assert snapshot["summary"]["year_month"] == "2026-07"

        source = sqlite3.connect(database)
        user_columns = {
            row[1] for row in source.execute("PRAGMA table_info(users)")
        }
        finance_tables = {
            row[0]
            for row in source.execute(
                """
                SELECT name FROM sqlite_master
                WHERE type = 'table' AND name LIKE 'company_profit_%'
                """
            )
        }
        source.close()
        assert "can_view_company_profit" not in user_columns
        assert "can_edit_company_profit" not in user_columns
        assert not finance_tables
        print(
            "company_profit_web_retired=ok "
            "routes=404 snapshot=readonly source_unchanged=true"
        )
    finally:
        os.chdir(original_cwd)
