"""Authenticated read-only smoke checks for deployed fulfillment pages."""

import os
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app


conn = sqlite3.connect("woocommerce_orders.db")
row = conn.execute("SELECT id FROM users WHERE username='admin' LIMIT 1").fetchone()
conn.close()
assert row, "admin user not found"

client = app.test_client()
with client.session_transaction() as session:
    session["_user_id"] = str(row[0])
    session["_fresh"] = True

page = client.get("/fulfillment")
assert page.status_code == 200, page.status_code
assert "多仓履约".encode("utf-8") in page.data

listing = client.get("/api/fulfillment/orders")
assert listing.status_code == 200, listing.status_code
assert listing.get_json()["items"] == []

config = client.get("/api/fulfillment/config/options")
assert config.status_code == 200, config.status_code
assert config.get_json()["settings"]["oms_fulfillment_enabled"] is False

legacy = client.get("/api/shipping/pending")
assert legacy.status_code == 200, legacy.status_code
assert isinstance(legacy.get_json(), list)

print(
    f"authenticated_smoke=ok fulfillment_page=200 fulfillment_api=200 "
    f"config_api=200 legacy_pending_count={len(legacy.get_json())}"
)
