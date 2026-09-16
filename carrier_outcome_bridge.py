"""Link carrier evidence to the exact current parcel, never the whole order."""
from datetime import datetime, timezone

from fulfillment_service import SHIPMENT_TRANSITIONS, _table_exists, add_tracking_event


def _timestamp(value):
    if not value:
        return None
    try:
        result = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return result.replace(tzinfo=timezone.utc) if result.tzinfo is None else result
    except ValueError:
        return None


def bridge_outcome(conn, order_id, outcome, tracking_number, observed_at):
    if outcome not in {"delivered", "returned"} or not str(tracking_number or "").strip():
        return 0
    shared = ""
    if _table_exists(conn, "oms_shipment_fulfillments"):
        shared = " OR EXISTS (SELECT 1 FROM oms_shipment_fulfillments sf WHERE sf.shipment_id=s.id AND sf.fulfillment_id=f.id)"
    rows = conn.execute("""
        SELECT DISTINCT s.id,s.status,s.shipped_at
        FROM oms_order_fulfillment_state ofs
        JOIN oms_fulfillments f ON f.order_id=ofs.order_id AND f.revision=ofs.revision
        JOIN oms_shipments s ON (s.fulfillment_id=f.id""" + shared + """)
        WHERE ofs.order_id=? AND f.status NOT IN ('superseded','cancelled','failed_terminal')
          AND trim(s.tracking_number)=?
    """, (order_id, str(tracking_number).strip())).fetchall()
    changed = 0
    for row in rows:
        if outcome not in SHIPMENT_TRANSITIONS.get(row["status"], set()):
            continue
        observed, shipped = _timestamp(observed_at), _timestamp(row["shipped_at"])
        if not observed or (shipped and observed < shipped):
            continue
        add_tracking_event(conn, row["id"], "carrier-outcome", outcome,
                           raw_status=outcome, event_at=str(observed_at),
                           external_event_id=f"carrier-outcome:{row['id']}:{outcome}:{observed_at}",
                           description="订单物流检测结果按同一运单号关联包裹", commit=False)
        changed += 1
    return changed


def reconcile_cached_outcomes(conn, limit=100):
    """Repair old cached terminal evidence only when its tracking is unambiguous."""
    # Imported lazily; the detector imports this bridge only when writing.
    from resolve_outcomes import extract_tracking

    rows = conn.execute("""
        SELECT o.id,o.meta_data,o.line_items,o.shipping_lines,o.carrier_status,o.carrier_status_at
        FROM orders o
        WHERE o.status IN ('on-hold','shipped','partial-shipped')
          AND COALESCE(o.delivery_confirmed,0)=0
          AND COALESCE(o.is_undelivered,0)=0 AND COALESCE(o.is_problem_return,0)=0
          AND o.carrier_status IN ('delivered','returned') AND o.carrier_status_at IS NOT NULL
          AND EXISTS (SELECT 1 FROM oms_order_fulfillment_state ofs WHERE ofs.order_id=o.id)
        ORDER BY o.id
    """).fetchall()
    changed = 0
    for row in rows:
        number, _ = extract_tracking(row)
        numbers = [str(r[0]).strip() for r in conn.execute(
            "SELECT DISTINCT trim(tracking_number) FROM shipping_logs WHERE order_id=? AND NULLIF(trim(tracking_number),'') IS NOT NULL",
            (row["id"],)).fetchall()]
        if len(numbers) != 1 or not number or str(number).strip() != numbers[0]:
            continue
        changed += bridge_outcome(conn, row["id"], row["carrier_status"], number, row["carrier_status_at"])
        if changed >= limit:
            break
    conn.commit()
    return changed
