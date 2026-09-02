import shutil
import sqlite3
import threading
from pathlib import Path

import pytest

import app as app_module
from sync_runtime_status import save_sync_runtime_status


ROOT = Path(__file__).resolve().parents[1]


class _PausedThread:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs

    def start(self):
        return None


@pytest.fixture()
def permission_app(tmp_path, monkeypatch):
    db_path = tmp_path / "site-sync-permissions.db"
    shutil.copyfile(ROOT / "woocommerce_orders.db", db_path)
    monkeypatch.setattr(app_module, "DB_FILE", str(db_path))
    monkeypatch.setattr(threading, "Thread", _PausedThread)
    app_module.app.config.update(TESTING=True)
    app_module.SYNC_STATUS.clear()

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    existing_user_columns = {
        row[1] for row in conn.execute("PRAGMA table_info(users)").fetchall()
    }
    for column in (
        "can_view_report",
        "can_view_inventory",
        "can_manage_inventory",
        "can_view_reconciliation",
        "can_edit_reconciliation",
        "can_view_costs",
        "can_edit_costs",
        "can_manage_blocklist",
        "can_manage_products",
        "can_manage_own_products",
        "can_view_own_sales_board",
        "can_view_shipping",
    ):
        if column not in existing_user_columns:
            conn.execute(f"ALTER TABLE users ADD COLUMN {column} INTEGER DEFAULT 0")
    existing_site_columns = {
        row[1] for row in conn.execute("PRAGMA table_info(sites)").fetchall()
    }
    for definition in (
        "mask_id TEXT",
        "api_read_status TEXT",
        "api_write_status TEXT",
    ):
        column = definition.split()[0]
        if column not in existing_site_columns:
            conn.execute(f"ALTER TABLE sites ADD COLUMN {definition}")
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS exchange_rates (
            id INTEGER PRIMARY KEY,
            year_month TEXT,
            currency TEXT,
            rate_to_cny REAL,
            updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS product_masters (
            id INTEGER PRIMARY KEY,
            label TEXT,
            url TEXT,
            consumer_key TEXT,
            api_status TEXT,
            last_api_error TEXT,
            last_tested_at TEXT,
            created_at TEXT,
            updated_at TEXT
        );
        """
    )
    conn.execute("DELETE FROM users")
    conn.execute("DELETE FROM sites")
    conn.execute("DELETE FROM sync_runtime_status")
    users = [
        (1, "admin", "Admin", "admin", 1, 0),
        (2, "alice", "金毅", "user", 0, 1),
        (3, "bob", "王渝淞", "user", 0, 1),
        (4, "plain", "无权限", "user", 0, 0),
        (5, "operator-admin", "运营管理员", "admin", 1, 0),
    ]
    conn.executemany(
        """
        INSERT INTO users (
            id, username, password_hash, name, role,
            can_manage_users, can_manage_own_site_sync
        ) VALUES (?, ?, 'test-only', ?, ?, ?, ?)
        """,
        users,
    )
    conn.executemany(
        """
        INSERT INTO sites (
            id, url, consumer_key, consumer_secret, manager, country, last_sync
        ) VALUES (?, ?, ?, ?, ?, 'PL', '2026-09-01 12:00:00')
        """,
        [
            (11, "https://alice.example", "ck_alice_secret", "cs_alice_secret", "金毅"),
            (22, "https://bob.example", "ck_bob_secret", "cs_bob_secret", "王渝淞"),
        ],
    )
    conn.commit()
    conn.close()
    yield app_module.app
    app_module.SYNC_STATUS.clear()


def _client_for(flask_app, user_id):
    client = flask_app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = str(user_id)
        session["_fresh"] = True
    return client


def test_connected_sites_page_contains_only_the_managers_sites(permission_app):
    response = _client_for(permission_app, 2).get("/settings")

    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert "alice.example" in html
    assert "bob.example" not in html
    assert "ck_alice_secret" not in html
    assert "Consumer Key" not in html
    assert "添加站点" not in html
    assert "数据备份与灾备" not in html
    assert "site_sync_settings.js" in html


def test_user_without_either_settings_permission_is_denied(permission_app):
    response = _client_for(permission_app, 4).get("/settings")

    assert response.status_code == 403


def test_full_settings_and_permission_ui_remain_available_to_super_admin(permission_app):
    client = _client_for(permission_app, 1)

    settings = client.get("/settings")
    users_page = client.get("/users")
    users_api = client.get("/api/users")

    assert settings.status_code == 200
    settings_html = settings.get_data(as_text=True)
    assert "alice.example" in settings_html
    assert "bob.example" in settings_html
    assert "Consumer Key" in settings_html
    assert "添加站点" in settings_html
    assert "site_sync_settings.js" not in settings_html
    assert users_page.status_code == 200
    assert "本人站点同步权限" in users_page.get_data(as_text=True)
    assert users_api.status_code == 200
    alice = next(user for user in users_api.get_json() if user["username"] == "alice")
    assert alice["can_manage_own_site_sync"] == 1


def test_own_site_sync_is_allowed_but_cross_site_and_global_sync_are_denied(permission_app):
    client = _client_for(permission_app, 2)

    own = client.post("/api/sync", json={"site_id": 11})
    other = client.post("/api/sync", json={"site_id": 22})
    global_sync = client.post("/api/sync/all", json={})

    assert own.status_code == 200
    assert other.status_code == 403
    assert global_sync.status_code == 403


def test_deep_and_clean_sync_create_site_bound_status(permission_app):
    client = _client_for(permission_app, 2)

    deep = client.post("/api/sync/deep/11")
    clean = client.post("/api/sync/clean/11")

    assert deep.status_code == 200
    assert clean.status_code == 200
    for response in (deep, clean):
        status_id = response.get_json()["sync_id"]
        status = client.get(f"/api/sync/status/{status_id}")
        assert status.status_code == 200
        assert status.get_json()["site_id"] == 11


def test_sync_status_cannot_be_read_through_another_owned_site(permission_app):
    conn = sqlite3.connect(app_module.DB_FILE)
    save_sync_runtime_status(
        conn,
        1234567,
        {
            "site_id": 22,
            "status": "running",
            "message": "Bob site",
            "logs": ["private status"],
        },
    )
    conn.close()

    response = _client_for(permission_app, 2).get("/api/sync/status/1234567")

    assert response.status_code == 403
    assert "private status" not in response.get_data(as_text=True)


def test_only_super_admin_can_grant_or_revoke_own_site_sync(permission_app):
    super_admin = _client_for(permission_app, 1)
    base_payload = {
        "name": "金毅",
        "role": "user",
        "can_manage_own_site_sync": 1,
        "reconciliation_scope": "all",
    }

    granted = super_admin.put("/api/users/2", json=base_payload)
    assert granted.status_code == 200

    operator_admin = _client_for(permission_app, 5)
    attempted_revoke = operator_admin.put(
        "/api/users/2",
        json={"name": "金毅", "role": "user", "can_manage_own_site_sync": 0},
    )
    assert attempted_revoke.status_code == 200

    conn = sqlite3.connect(app_module.DB_FILE)
    stored = conn.execute(
        "SELECT can_manage_own_site_sync FROM users WHERE id=2"
    ).fetchone()[0]
    conn.close()
    assert stored == 1


def test_permission_cannot_be_granted_without_an_owned_site(permission_app):
    response = _client_for(permission_app, 1).put(
        "/api/users/4",
        json={
            "name": "无权限",
            "role": "user",
            "can_manage_own_site_sync": 1,
            "reconciliation_scope": "all",
        },
    )

    assert response.status_code == 400
    assert "名下没有站点" in response.get_json()["error"]
