"""Authenticated read-only smoke checks for deployed fulfillment pages."""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app, get_db_connection


anonymous = app.test_client()
assert anonymous.get("/").status_code in {302, 401}
assert anonymous.get("/api/sync/active").status_code in {302, 401}

conn = get_db_connection()
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
assert isinstance(listing.get_json()["items"], list)

config = client.get("/api/fulfillment/config/options")
assert config.status_code == 200, config.status_code
assert isinstance(
    config.get_json()["settings"]["oms_fulfillment_enabled"], bool
)

legacy = client.get("/api/shipping/pending")
assert legacy.status_code == 200, legacy.status_code
assert isinstance(legacy.get_json(), list)

settings = client.get("/settings")
assert settings.status_code == 200, settings.status_code
active_sync = client.get("/api/sync/active")
assert active_sync.status_code == 200, active_sync.status_code
assert isinstance(active_sync.get_json(), dict)

readonly_routes = (
    "/",
    "/orders?quick_date=this_month&per_page=20",
    "/monthly",
    "/customers",
    "/shipping",
    "/inventory/stock",
    "/fulfillment",
    "/settings",
    "/users",
    "/products",
    "/product-manager",
    "/product-costs",
    "/partner-reconciliation",
    "/sales-board",
    "/report",
    "/order-notifications",
    "/au-orders",
    "/api/shipping/pending",
    "/api/shipping/shipped",
    "/api/shipping/pending-outcome?days=0",
    "/api/inv/overview",
    "/api/sync/dashboard",
    "/api/settings/autosync",
    "/api/cron/status",
)
route_statuses = {path: client.get(path).status_code for path in readonly_routes}
failures = {path: status for path, status in route_statuses.items() if status != 200}
assert not failures, failures

print(
    f"authenticated_smoke=ok fulfillment_page=200 fulfillment_api=200 "
    f"config_api=200 settings=200 sync_active=200 routes={len(route_statuses)} "
    f"legacy_pending_count={len(legacy.get_json())}"
)
