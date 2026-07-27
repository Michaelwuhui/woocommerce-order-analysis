"""Validate migration 009 and Czech allocation against a disposable DB copy."""

from fulfillment_common import get_conn
from fulfillment_service import plan_order


conn = get_conn()
carrier = conn.execute(
    "SELECT slug, name, tracking_url, is_active FROM shipping_carriers WHERE slug='packeta'"
).fetchone()
assert carrier, "Packeta carrier is missing"
assert carrier["tracking_url"] == "https://tracking.packeta.com/en/{tracking}"

pl = conn.execute(
    """SELECT wi.inventory_authority
       FROM oms_warehouse_integrations wi
       JOIN warehouses w ON w.id=wi.warehouse_id
       WHERE w.country='PL' ORDER BY w.id LIMIT 1"""
).fetchone()
assert pl and pl["inventory_authority"] == "manual_partner"

orders = conn.execute(
    """SELECT o.id
       FROM orders o
       JOIN sites s ON s.url=o.source
       WHERE s.country='CZ'
         AND o.status IN ('processing','offline','on-hold','partial-shipped','shipped')
       ORDER BY o.date_created DESC LIMIT 10"""
).fetchall()

print(f"packeta={dict(carrier)}")
print(f"czech_orders_checked={len(orders)}")
for row in orders:
    result = plan_order(conn, row["id"])
    allocations = [
        dict(allocation)
        for allocation in conn.execute(
            """SELECT w.country, f.mode, f.provider, f.status,
                      SUM(fi.allocated_qty) AS quantity
               FROM oms_fulfillments f
               JOIN warehouses w ON w.id=f.warehouse_id
               JOIN oms_fulfillment_items fi ON fi.fulfillment_id=f.id
               WHERE f.order_id=? AND f.status!='superseded'
               GROUP BY w.country, f.mode, f.provider, f.status
               ORDER BY w.country""",
            (row["id"],),
        ).fetchall()
    ]
    print(
        {
            "order_id": row["id"],
            "action": result["action"],
            "allocations": allocations,
            "shortage_count": len(result["shortages"]),
            "reason": result["reason"],
        }
    )

# Planning a disposable copy must never enqueue or submit a real WMS request.
assert (
    conn.execute(
        """SELECT COUNT(*) FROM oms_integration_jobs
           WHERE job_type='SUBMIT_HU_FULFILLMENT'"""
    ).fetchone()[0]
    == 0
)
conn.close()
