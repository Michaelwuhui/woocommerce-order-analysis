"""Recover evidenced carrier outcomes from two obsolete fulfillment states.

This repairs derived state only. It never allocates inventory, invents a parcel,
ships a line, clears a human hold, or calls a store. The confirmation workers
still own their normal eligibility checks and external-operation ledger.
"""
from collections import defaultdict
from datetime import datetime, timezone

import db_backend
from fulfillment_service import (
    _source_items_match_synced, _table_exists, record_event, recompute_order_status,
)
from order_shipments import extract_tracking_candidates


STALE_PLAN_REASON = "异步任务失败: PLAN_ORDER - 履约已开始，不能自动重新分仓；请转人工处理"
LEGACY_EVENT = "legacy_carrier_outcome_reconciled"


def _time(value):
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return parsed.replace(tzinfo=timezone.utc) if parsed.tzinfo is None else parsed
    except (ValueError, TypeError):
        return None


def _single_tracking(conn, order):
    logs = conn.execute(
        "SELECT tracking_number,carrier_slug,shipped_at FROM shipping_logs WHERE order_id=?",
        (order["id"],),
    ).fetchall()
    candidates = extract_tracking_candidates(
        order["meta_data"], order["line_items"], order["shipping_lines"], logs,
    )
    if len(candidates) != 1:
        return None
    observed = _time(order["carrier_status_at"])
    shipped_times = [_time(row["shipped_at"]) for row in logs if row["tracking_number"]]
    # A changed upstream order must first receive fresh carrier evidence.
    shipped_times.append(_time(order["date_modified"]))
    if not observed or any(stamp and stamp > observed for stamp in shipped_times):
        return None
    return candidates[0]["tracking_number"]


def _quantities_consistent(conn, order_id, revision):
    items = conn.execute("SELECT * FROM oms_order_items WHERE order_id=?", (order_id,)).fetchall()
    if not items or any(row["shortage_qty"] or row["cancelled_qty"] for row in items):
        return False
    assigned = defaultdict(int)
    for row in conn.execute("""
        SELECT fi.*, f.status AS fulfillment_status,
               COALESCE((SELECT SUM(si.quantity) FROM oms_shipment_items si
                         JOIN oms_shipments s ON s.id=si.shipment_id
                         WHERE si.fulfillment_item_id=fi.id AND s.status!='cancelled'),0) AS shipped_qty
        FROM oms_fulfillment_items fi JOIN oms_fulfillments f ON f.id=fi.fulfillment_id
        WHERE f.order_id=? AND f.revision=? AND f.status NOT IN ('superseded','cancelled')
    """, (order_id, revision)).fetchall():
        qty = int(row["allocated_qty"]) - int(row["cancelled_qty"])
        fulfilled = int(row["fulfilled_qty"])
        if qty < 0 or not 0 <= fulfilled <= qty or fulfilled != int(row["shipped_qty"]):
            return False
        if row["fulfillment_status"] == "delivered" and fulfilled != qty:
            return False
        assigned[row["order_item_id"]] += qty
    return assigned == {row["id"]: int(row["ordered_qty"]) for row in items}


def recover_terminal_state(conn, order_id):
    """Repair one derived state inside the caller's transaction; return details."""
    lock = " FOR UPDATE" if db_backend.is_postgres_backend() else ""
    order = conn.execute("SELECT * FROM orders WHERE id=?" + lock, (order_id,)).fetchone()
    state = conn.execute("SELECT * FROM oms_order_fulfillment_state WHERE order_id=?" + lock,
                         (order_id,)).fetchone()
    if (not order or not state or order["status"] not in {"on-hold", "shipped", "partial-shipped"}
            or order["payment_method"] != "cod" or order["delivery_confirmed"]
            or order["is_undelivered"] or order["is_problem_return"] or state["has_shortage"]
            or order["carrier_status"] not in {"delivered", "returned"}
            or state["completion_sync_status"] in {"pending", "running", "synced"}):
        return None
    if not _source_items_match_synced(conn, order):
        return None
    number = _single_tracking(conn, order)
    if not number or conn.execute("""
        SELECT 1 FROM oms_integration_jobs WHERE aggregate_type='order' AND aggregate_id=?
          AND job_type IN ('PLAN_ORDER','COMPLETE_WOOCOMMERCE_ORDER')
          AND status IN ('pending','retry','running') LIMIT 1
    """, (order_id,)).fetchone():
        return None
    fulfillments = conn.execute("SELECT * FROM oms_fulfillments WHERE order_id=?" + lock,
                                (order_id,)).fetchall()
    if state["manual_review"] and state["manual_reason"] == STALE_PLAN_REASON:
        review = conn.execute("""
            SELECT * FROM oms_domain_events WHERE aggregate_type='order' AND aggregate_id=?
              AND event_type='manual_review_required' ORDER BY id DESC LIMIT 1
        """, (order_id,)).fetchone()
        if (not review or review["actor_type"] != "system" or review["actor_id"]
                or review["reason"] != STALE_PLAN_REASON):
            return None
        active = [f for f in fulfillments if f["revision"] == state["revision"]
                  and f["status"] not in {"superseded", "cancelled"}]
        if (not active or any(f["status"] in {"manual_hold", "manual_review", "failed_terminal", "rejected"}
                              or f["last_error_code"] for f in active)
                or not _quantities_consistent(conn, order_id, state["revision"])):
            return None
        matched = conn.execute("""
            SELECT s.id FROM oms_shipments s JOIN oms_fulfillments f ON f.id=s.fulfillment_id
            WHERE f.order_id=? AND f.revision=? AND trim(s.tracking_number)=? AND s.status=?
        """, (order_id, state["revision"], number, order["carrier_status"])).fetchone()
        if not matched:
            return None
        conn.execute("""UPDATE oms_order_fulfillment_state SET manual_review=0,manual_reason=NULL
                        WHERE order_id=?""", (order_id,))
        # Leave completion to the bounded Celery worker and its shared ledger.
        result = recompute_order_status(conn, order_id, commit=False, enqueue_completion=False)
        record_event(conn, "order", order_id, "stale_plan_review_recovered",
                     from_status=state["aggregate_status"], to_status=result["aggregate_status"],
                     reason="商品未变化且包裹有明确物流结局，清理旧分仓任务误留的复核状态",
                     payload={"review_event_id": review["id"], "tracking_number": number,
                              "carrier_status": order["carrier_status"]})
        return {"order_id": order_id, "kind": "stale_plan_review", "status": result["aggregate_status"]}

    # Old manual shipping already cleared this plan's shortage but left its
    # state row behind. Its absence of OMS parcels does not mean unshipped.
    if (fulfillments or state["manual_review"] or state["manual_reason"]
            or state["aggregate_status"] != "shipped" or order["status"] != "shipped"):
        return None
    legacy = conn.execute("""
        SELECT id FROM oms_domain_events WHERE aggregate_type='order' AND aggregate_id=?
          AND event_type='legacy_terminal_shortage_reconciled'
          AND actor_type='system' AND to_status='shipped' ORDER BY id DESC LIMIT 1
    """, (order_id,)).fetchone()
    if not legacy or conn.execute("""SELECT 1 FROM oms_order_items WHERE order_id=?
        AND (allocated_qty!=0 OR shortage_qty!=0 OR cancelled_qty!=0) LIMIT 1""", (order_id,)).fetchone():
        return None
    outcome = order["carrier_status"]
    conn.execute("""UPDATE oms_order_fulfillment_state SET aggregate_status=?,updated_at=CURRENT_TIMESTAMP
                    WHERE order_id=?""", (outcome, order_id))
    record_event(conn, "order", order_id, LEGACY_EVENT, from_status=state["aggregate_status"],
                 to_status=outcome, reason="旧手工发货订单按唯一运单的物流结局更新履约状态",
                 payload={"legacy_event_id": legacy["id"], "tracking_number": number,
                          "carrier_status_at": str(order["carrier_status_at"])})
    return {"order_id": order_id, "kind": "legacy_manual_shipment", "status": outcome}


def recover_terminal_states(conn, *, outcome, limit=100):
    rows = conn.execute("""
        SELECT o.id FROM orders o JOIN oms_order_fulfillment_state st ON st.order_id=o.id
        WHERE o.carrier_status=? AND o.carrier_status_at IS NOT NULL
          AND o.status IN ('on-hold','shipped','partial-shipped') AND o.payment_method='cod'
          AND COALESCE(o.delivery_confirmed,0)=0 AND COALESCE(o.is_undelivered,0)=0
          AND COALESCE(o.is_problem_return,0)=0 AND COALESCE(st.has_shortage,0)=0
          AND (st.manual_reason=? OR (st.aggregate_status='shipped' AND NOT EXISTS
               (SELECT 1 FROM oms_fulfillments f WHERE f.order_id=o.id)))
        ORDER BY o.date_created,o.id
    """, (outcome, STALE_PLAN_REASON)).fetchall()
    recovered = []
    for row in rows:
        result = recover_terminal_state(conn, row["id"])
        if result:
            recovered.append(result)
        if len(recovered) >= limit:
            break
    conn.commit()
    return recovered


def pending_outcome_blockers(conn, order_ids):
    """Describe why a visible carrier delivery cannot complete the whole order."""
    if not order_ids or not _table_exists(conn, "oms_order_fulfillment_state"):
        return {}
    result = {}
    for offset in range(0, len(order_ids), 200):
        ids = order_ids[offset:offset + 200]
        for row in conn.execute("SELECT order_id,aggregate_status,manual_review FROM oms_order_fulfillment_state "
                                "WHERE order_id IN (" + ",".join("?" for _ in ids) + ")", ids).fetchall():
            status = row["aggregate_status"]
            if status == "delivered" and not row["manual_review"]:
                continue
            reason = ("需先处理人工复核" if row["manual_review"] else
                      "部分包裹已签收，仍有商品未完成发货或签收" if status == "partially_delivered" else
                      "履约尚未全部完成，请核对各商品的运单和物流状态")
            result[row["order_id"]] = {"status": status, "reason": reason}
    return result
