"""Settings permissions through real routes on an empty PostgreSQL schema copy."""
import os

import pytest
import requests

import db_backend as db
import sync_service


pytestmark = pytest.mark.skipif(
    not db.is_postgres_backend(), reason="isolated PostgreSQL required"
)


@pytest.fixture()
def permission_app(monkeypatch):
    test_database = os.environ.get("WOO_DB_NAME_OVERRIDE", "")
    assert test_database.startswith("woo_return_loss_test_"), "Use an empty isolated test database"
    monkeypatch.setattr(requests.sessions.Session, "request",
                        lambda *a, **k: pytest.fail("Real network calls forbidden"))
    monkeypatch.setattr(sync_service, "publish_pending_outbox", lambda **kw: 0)
    import app as app_module
    app_module.app.config.update(TESTING=True)
    conn = db.connect()
    assert conn.execute("SELECT current_database()").fetchone()[0] == test_database
    # These tests own only synthetic users/sites. No SQLite business snapshot is needed.
    for table in ("sync_task_outbox", "sync_runs"):
        conn.execute(f"DELETE FROM {table}")
    users = [
        (1, "admin", "Admin", "admin", True, False),
        (2, "alice", "Alice Test", "user", False, True),
        (3, "bob", "Bob Test", "user", False, True),
        (4, "plain", "No Permission", "user", False, False),
        (5, "operator-admin", "Operator Test", "admin", True, False),
    ]
    for row in users:
        conn.execute("""
            INSERT INTO users (id,username,password_hash,name,role,
                               can_manage_users,can_manage_own_site_sync)
            VALUES (?,?,'test-only',?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET name=excluded.name,role=excluded.role,
                can_manage_users=excluded.can_manage_users,
                can_manage_own_site_sync=excluded.can_manage_own_site_sync
        """, row)
    for row in [
        (11, "https://alice.example.invalid", "ck_alice_secret", "cs_alice_secret", "Alice Test"),
        (22, "https://bob.example.invalid", "ck_bob_secret", "cs_bob_secret", "Bob Test"),
    ]:
        conn.execute("""
            INSERT INTO sites (id,url,consumer_key,consumer_secret,manager,country,last_sync)
            VALUES (?,?,?,?,?,'PL','2026-09-01 12:00:00')
            ON CONFLICT(id) DO UPDATE SET manager=excluded.manager
        """, row)
    conn.commit()
    conn.close()
    yield app_module.app


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
    assert html.count('id="syncProgressModal"') == 1


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
    assert settings_html.count('id="syncProgressModal"') == 1
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

    assert own.status_code == 202
    assert other.status_code == 403
    assert global_sync.status_code == 403


def test_deep_sync_creates_site_bound_status(permission_app):
    client = _client_for(permission_app, 2)
    response = client.post("/api/sync/deep/11")
    assert response.status_code == 202
    status_id = response.get_json()["sync_id"]
    status = client.get(f"/api/sync/status/{status_id}")
    assert status.status_code == 200
    assert [site["site_id"] for site in status.get_json()["sites"]] == [11]


def test_clean_sync_requires_full_admin_and_queues_durable_job(permission_app):
    own_site_manager = _client_for(permission_app, 2)
    assert own_site_manager.post("/api/sync/clean/11").status_code == 403
    assert own_site_manager.post("/api/sync/clean/all").status_code == 403
    assert "clean-sync-btn" not in own_site_manager.get("/settings").get_data(as_text=True)

    conn = db.connect()
    assert conn.execute("SELECT COUNT(*) FROM sync_runs").fetchone()[0] == 0
    conn.close()

    admin = _client_for(permission_app, 1)
    response = admin.post("/api/sync/clean/11")
    assert response.status_code == 202
    run_id = response.get_json()["sync_id"]
    status = admin.get(f"/api/sync/status/{run_id}")
    assert status.status_code == 200
    assert status.get_json()["mode"] == "clean"
    assert [site["site_id"] for site in status.get_json()["sites"]] == [11]
    conn = db.connect()
    try:
        assert conn.execute(
            "SELECT task_name FROM sync_task_outbox WHERE dedupe_key=?",
            (f"clean:{run_id}:11",),
        ).fetchone()[0] == "woo_sync.clean_site"
    finally:
        conn.close()


def test_admin_can_queue_global_clean_and_manage_weekly_schedule(permission_app):
    admin = _client_for(permission_app, 1)
    initial = admin.get("/api/cron/clean/status")
    assert initial.status_code == 200
    assert initial.get_json()["enabled"] is False

    configured = admin.post(
        "/api/cron/clean/setup", json={"day": 0, "hour": 4, "minute": 30}
    )
    assert configured.status_code == 200
    schedule = admin.get("/api/cron/clean/status").get_json()
    assert (schedule["enabled"], schedule["day"], schedule["hour"], schedule["minute"]) == (
        True, 0, 4, 30
    )
    assert admin.delete("/api/cron/clean/remove").status_code == 200
    assert admin.get("/api/cron/clean/status").get_json()["enabled"] is False

    response = admin.post("/api/sync/clean/all")
    assert response.status_code == 202
    status = admin.get(f"/api/sync/status/{response.get_json()['sync_id']}")
    assert status.status_code == 200
    assert status.get_json()["mode"] == "clean"
    assert {site["site_id"] for site in status.get_json()["sites"]} == {11, 22}


def test_sync_status_cannot_be_read_through_another_owned_site(permission_app):
    status, created = sync_service.start_sync(
        mode="quick", created_by="pytest:permissions", site_ids=[22], publish=False
    )
    assert created
    response = _client_for(permission_app, 2).get(f"/api/sync/status/{status['run_id']}")

    assert response.status_code == 403
    assert "private status" not in response.get_data(as_text=True)


def test_only_super_admin_can_grant_or_revoke_own_site_sync(permission_app):
    super_admin = _client_for(permission_app, 1)
    base_payload = {
        "name": "Alice Test",
        "role": "user",
        "can_manage_own_site_sync": 1,
        "reconciliation_scope": "all",
    }

    granted = super_admin.put("/api/users/2", json=base_payload)
    assert granted.status_code == 200

    operator_admin = _client_for(permission_app, 5)
    attempted_revoke = operator_admin.put(
        "/api/users/2",
        json={"name": "Alice Test", "role": "user", "can_manage_own_site_sync": 0},
    )
    assert attempted_revoke.status_code == 200

    conn = db.connect()
    stored = conn.execute(
        "SELECT can_manage_own_site_sync FROM users WHERE id=2"
    ).fetchone()[0]
    conn.close()
    assert stored == 1


def test_permission_cannot_be_granted_without_an_owned_site(permission_app):
    response = _client_for(permission_app, 1).put(
        "/api/users/4",
        json={
            "name": "No Permission",
            "role": "user",
            "can_manage_own_site_sync": 1,
            "reconciliation_scope": "all",
        },
    )

    assert response.status_code == 400
    assert "名下没有站点" in response.get_json()["error"]
