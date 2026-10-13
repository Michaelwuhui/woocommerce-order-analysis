"""Notification activation with the deployed PostgreSQL schema and no real HTTP."""
import json
import os
from datetime import datetime, timedelta, timezone

import pytest
import requests

import db_backend as db
from order_notification_service import _automatic_rollout_reason, write_site_watermark
from sync_service import get_connection

pytestmark = pytest.mark.skipif(not db.is_postgres_backend(), reason="isolated PostgreSQL required")
SITE_ID = 991310013
OLD_URL = "https://watermark-old.example.invalid"
NEW_URL = "https://watermark-new.example.invalid"
SETTING = "order_notification_auto_watermarks_json"


@pytest.fixture
def isolated(monkeypatch):
    database = os.environ.get("WOO_DB_NAME_OVERRIDE", "")
    assert database.startswith("woo_return_loss_test_"), "Refusing a production database"

    def forbidden(*args, **kwargs):
        pytest.fail("Notification watermark verification must not call external HTTP")

    monkeypatch.setattr(requests.sessions.Session, "request", forbidden)
    conn = get_connection()
    assert conn.execute("SELECT current_database()").fetchone()[0] == database
    try:
        conn.execute("INSERT INTO sites(id,url,consumer_key,consumer_secret) VALUES (?,?,?,?)",
            (SITE_ID, OLD_URL, "synthetic-key", "synthetic-secret"))
        for woo_id in (101, 102):
            conn.execute("INSERT INTO orders(id,woo_id,source,status,date_created_gmt) VALUES (?,?,?,'processing',?)",
                (f"{SITE_ID}-{woo_id}", woo_id, OLD_URL, "2020-01-01T00:00:00"))
        conn.execute("INSERT INTO settings(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (SETTING, json.dumps({OLD_URL: {"max_woo_id": 50}})))
        yield conn
    finally:
        conn.rollback()
        conn.close()


def test_domain_change_retains_order_id_cutoff_and_legacy_watermark(isolated):
    entry = write_site_watermark(isolated, SITE_ID, NEW_URL, reason="site_url_changed")
    assert entry["max_woo_id"] == 102
    configured = json.loads(isolated.execute("SELECT value FROM settings WHERE key=?", (SETTING,)).fetchone()[0])
    assert configured[OLD_URL] == {"max_woo_id": 50}
    assert configured[f"site:{SITE_ID}"] == entry
    snapshot = {"site_id": SITE_ID, "store_id": NEW_URL,
        "order_id": f"{SITE_ID}-102", "woo_id": 102}
    assert _automatic_rollout_reason(isolated, snapshot) == "before_activation_watermark"
    snapshot.update(order_id=f"{SITE_ID}-103", woo_id=103)
    assert _automatic_rollout_reason(isolated, snapshot) is None


def test_old_imported_order_is_rejected_using_real_gmt_timestamp(isolated):
    entry = write_site_watermark(isolated, SITE_ID, NEW_URL, reason="site_restored", not_before=True)
    assert entry["not_before_gmt"]
    for woo_id, created in ((103, "2020-01-01T00:00:00"),
            (104, datetime.now(timezone.utc) + timedelta(seconds=2))):
        isolated.execute("INSERT INTO orders(id,woo_id,source,status,date_created_gmt) VALUES (?,?,?,'processing',?)",
            (f"{SITE_ID}-{woo_id}", woo_id, NEW_URL, created))
    snapshot = {"site_id": SITE_ID, "store_id": NEW_URL,
        "order_id": f"{SITE_ID}-103", "woo_id": 103}
    assert _automatic_rollout_reason(isolated, snapshot) == "before_activation_watermark"
    snapshot.update(order_id=f"{SITE_ID}-104", woo_id=104)
    assert _automatic_rollout_reason(isolated, snapshot) is None


def test_invalid_configuration_leaves_site_and_watermark_transaction_reversible(isolated):
    isolated.execute("UPDATE settings SET value='invalid-json' WHERE key=?", (SETTING,))
    isolated.execute("SAVEPOINT site_change")
    isolated.execute("UPDATE sites SET url=? WHERE id=?", (NEW_URL, SITE_ID))
    with pytest.raises(ValueError, match="Corrupt order_notification_auto_watermarks_json"):
        write_site_watermark(isolated, SITE_ID, NEW_URL, reason="site_url_changed")
    isolated.execute("ROLLBACK TO SAVEPOINT site_change")
    assert isolated.execute("SELECT url FROM sites WHERE id=?", (SITE_ID,)).fetchone()[0] == OLD_URL
    assert isolated.execute("SELECT value FROM settings WHERE key=?", (SETTING,)).fetchone()[0] == "invalid-json"
