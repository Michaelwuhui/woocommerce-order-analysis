"""Read-only post-deploy invariants for the live SQLite database."""

import os
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


conn = sqlite3.connect("woocommerce_orders.db")
conn.row_factory = sqlite3.Row
try:
    print("integrity=" + conn.execute("PRAGMA integrity_check").fetchone()[0])
    settings = {
        r["key"]: r["value"] for r in conn.execute(
            "SELECT key,value FROM settings WHERE key IN ('oms_fulfillment_enabled','oms_auto_plan_enabled')"
        )
    }
    print("fulfillment_enabled=" + settings.get("oms_fulfillment_enabled", "missing"))
    print("auto_plan_enabled=" + settings.get("oms_auto_plan_enabled", "missing"))
    integration = conn.execute(
        """SELECT external_code,channel_code,auto_submit,tracking_mode
           FROM oms_warehouse_integrations WHERE provider='hungary_wms'"""
    ).fetchone()
    print(
        "hu_integration="
        + (f"{integration['external_code']} channel={integration['channel_code']} "
           f"auto_submit={integration['auto_submit']} tracking={integration['tracking_mode']}"
           if integration else "missing")
    )
    for table in (
        "oms_order_fulfillment_state", "oms_fulfillments", "oms_shipments",
        "oms_integration_jobs", "oms_external_api_calls", "oms_webhook_inbox",
    ):
        count = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        print(f"{table}={count}")
finally:
    conn.close()
