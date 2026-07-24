"""Domain service for durable multi-warehouse fulfillment.

The service owns allocation, legal state transitions, shortage/manual-review
flags, shipments and aggregate completion.  External HTTP calls live in the
worker/adapters so database transactions never remain open across the network.
"""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import uuid
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from typing import Any

from fulfillment_common import json_dump, json_load, utcnow


ACTIVE_ORDER_STATUSES = {"processing", "offline", "on-hold", "partial-shipped", "shipped"}
PLANNABLE_FULFILLMENT_STATUSES = {
    "planned",
    "ready_to_pick",
    "ready_to_submit",
    "stock_shortage",
    "manual_hold",
}

FULFILLMENT_TRANSITIONS = {
    "planned": {"ready_to_pick", "ready_to_submit", "stock_shortage", "manual_hold", "cancelled", "superseded"},
    "stock_shortage": {"ready_to_pick", "ready_to_submit", "manual_hold", "cancel_pending", "cancelled", "superseded"},
    "manual_hold": {"ready_to_pick", "ready_to_submit", "stock_shortage", "cancel_pending", "cancelled", "superseded"},
    "ready_to_pick": {"picking", "packed", "shipped", "manual_hold", "cancelled", "superseded"},
    "picking": {"packed", "shipped", "stock_shortage", "manual_hold", "cancel_pending"},
    "packed": {"shipped", "manual_hold", "cancel_pending"},
    "ready_to_submit": {"submitting", "stock_shortage", "manual_hold", "cancel_pending", "cancelled", "superseded"},
    "submitting": {"accepted", "submission_unknown", "rejected", "failed_retryable", "failed_terminal"},
    "submission_unknown": {"accepted", "ready_to_submit", "manual_review", "failed_retryable"},
    "failed_retryable": {"ready_to_submit", "submitting", "manual_review", "failed_terminal"},
    "accepted": {"picking", "packed", "shipped", "stock_shortage", "manual_hold", "cancel_pending"},
    "rejected": {"ready_to_submit", "manual_review", "cancelled"},
    "cancel_pending": {"cancelled", "cancel_rejected", "manual_review"},
    "cancel_rejected": {"accepted", "picking", "packed", "shipped", "manual_review"},
    "shipped": {"delivered", "exception", "returning", "manual_review"},
    "exception": {"shipped", "delivered", "returning", "returned", "manual_review"},
    "returning": {"returned", "delivered", "manual_review"},
    "manual_review": {"ready_to_pick", "ready_to_submit", "accepted", "picking", "packed", "shipped", "cancelled"},
    "delivered": set(),
    "returned": set(),
    "cancelled": set(),
    "failed_terminal": {"manual_review"},
    "superseded": set(),
}

SHIPMENT_TRANSITIONS = {
    "label_pending": {"label_ready", "shipped", "cancelled", "exception"},
    "label_ready": {"shipped", "cancelled", "exception"},
    "shipped": {"in_transit", "pickup_ready", "delivered", "undelivered", "exception", "expired", "cancelled"},
    "not_found": {"shipped", "in_transit", "exception", "expired"},
    "in_transit": {"pickup_ready", "delivered", "undelivered", "exception", "expired", "returning"},
    "pickup_ready": {"in_transit", "delivered", "undelivered", "exception", "expired", "returning"},
    "undelivered": {"in_transit", "delivered", "returning", "returned", "exception"},
    "exception": {"in_transit", "pickup_ready", "delivered", "returning", "returned"},
    "expired": {"in_transit", "delivered", "returning", "returned"},
    "returning": {"returned", "delivered", "exception"},
    "delivered": set(),
    "returned": set(),
    "cancelled": set(),
}

SHIPMENT_PROGRESS = {
    "label_pending": 0,
    "label_ready": 1,
    "not_found": 1,
    "shipped": 2,
    "in_transit": 3,
    "pickup_ready": 4,
    "undelivered": 4,
    "exception": 4,
    "expired": 4,
    "returning": 5,
    "returned": 6,
    "delivered": 7,
    "cancelled": 7,
}

COUNTRY_ZH = {
    "DE": "德国",
    "HU": "匈牙利",
    "SK": "斯洛伐克",
    "CZ": "捷克",
    "PL": "波兰",
    "BG": "保加利亚",
    "AT": "奥地利",
    "HR": "克罗地亚",
    "SI": "斯洛文尼亚",
    "RO": "罗马尼亚",
    "GR": "希腊",
    "EL": "希腊",
    "IT": "意大利",
}


class DomainError(RuntimeError):
    def __init__(self, message: str, code: str = "domain_error"):
        super().__init__(message)
        self.code = code


def _hash(value: Any) -> str:
    return hashlib.sha256(json_dump(value).encode("utf-8")).hexdigest()


def _uuid() -> str:
    return str(uuid.uuid4())


def _actor(actor: dict | None) -> tuple[str, str | None, str | None]:
    actor = actor or {}
    return (
        actor.get("type") or "system",
        str(actor.get("id")) if actor.get("id") is not None else None,
        actor.get("name"),
    )


def record_event(
    conn: sqlite3.Connection,
    aggregate_type: str,
    aggregate_id: str,
    event_type: str,
    *,
    from_status: str | None = None,
    to_status: str | None = None,
    actor: dict | None = None,
    correlation_id: str | None = None,
    reason: str | None = None,
    payload: Any = None,
) -> int:
    actor_type, actor_id, actor_name = _actor(actor)
    cur = conn.execute(
        '''INSERT INTO oms_domain_events
           (aggregate_type, aggregate_id, event_type, from_status, to_status,
            actor_type, actor_id, actor_name, correlation_id, reason, payload_json)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
        (
            aggregate_type,
            aggregate_id,
            event_type,
            from_status,
            to_status,
            actor_type,
            actor_id,
            actor_name,
            correlation_id,
            reason,
            json_dump(payload) if payload is not None else None,
        ),
    )
    return cur.lastrowid


def enqueue_job(
    conn: sqlite3.Connection,
    job_type: str,
    aggregate_type: str,
    aggregate_id: str,
    idempotency_key: str,
    payload: Any = None,
    *,
    available_at: str | None = None,
    max_attempts: int = 10,
) -> int:
    payload_json = json_dump(payload or {})
    conn.execute(
        '''INSERT INTO oms_integration_jobs
           (job_type, aggregate_type, aggregate_id, idempotency_key,
            payload_json, payload_hash, status, max_attempts, available_at)
           VALUES (?,?,?,?,?,?,'pending',?,COALESCE(?,CURRENT_TIMESTAMP))
           ON CONFLICT(idempotency_key) DO NOTHING''',
        (
            job_type,
            aggregate_type,
            aggregate_id,
            idempotency_key,
            payload_json,
            _hash(payload or {}),
            max_attempts,
            available_at,
        ),
    )
    row = conn.execute(
        "SELECT id FROM oms_integration_jobs WHERE idempotency_key=?", (idempotency_key,)
    ).fetchone()
    return row["id"]


def transition_fulfillment(
    conn: sqlite3.Connection,
    fulfillment_id: str,
    to_status: str,
    *,
    actor: dict | None = None,
    reason: str | None = None,
    correlation_id: str | None = None,
    extra_updates: dict | None = None,
) -> dict:
    row = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fulfillment_id,)).fetchone()
    if not row:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    current = row["status"]
    if current == to_status:
        return dict(row)
    if to_status not in FULFILLMENT_TRANSITIONS.get(current, set()):
        raise DomainError(f"非法履约状态转换: {current} → {to_status}", "illegal_transition")

    updates = {"status": to_status, "updated_at": utcnow(), "row_version": row["row_version"] + 1}
    if to_status == "accepted":
        updates["accepted_at"] = utcnow()
    elif to_status == "shipped":
        updates["shipped_at"] = utcnow()
    elif to_status == "delivered":
        updates["delivered_at"] = utcnow()
    elif to_status == "cancelled":
        updates["cancelled_at"] = utcnow()
    updates.update(extra_updates or {})
    allowed_columns = {
        "status", "updated_at", "row_version", "accepted_at", "shipped_at",
        "delivered_at", "cancelled_at", "submitted_at", "external_pick_code",
        "external_label_url", "payload_hash", "last_error_code",
        "last_error_message", "retry_count", "next_retry_at",
    }
    updates = {k: v for k, v in updates.items() if k in allowed_columns}
    assignments = ", ".join(f"{key}=?" for key in updates)
    conn.execute(
        f"UPDATE oms_fulfillments SET {assignments} WHERE id=? AND row_version=?",
        (*updates.values(), fulfillment_id, row["row_version"]),
    )
    if conn.execute("SELECT changes()").fetchone()[0] != 1:
        raise DomainError("履约单被其他操作更新，请刷新后重试", "optimistic_lock")
    record_event(
        conn,
        "fulfillment",
        fulfillment_id,
        "status_changed",
        from_status=current,
        to_status=to_status,
        actor=actor,
        correlation_id=correlation_id,
        reason=reason,
    )
    return dict(conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fulfillment_id,)).fetchone())


def transition_shipment(
    conn: sqlite3.Connection,
    shipment_id: str,
    to_status: str,
    *,
    actor: dict | None = None,
    reason: str | None = None,
    correlation_id: str | None = None,
    allow_correction: bool = False,
) -> dict:
    row = conn.execute("SELECT * FROM oms_shipments WHERE id=?", (shipment_id,)).fetchone()
    if not row:
        raise DomainError("包裹不存在", "shipment_not_found")
    current = row["status"]
    if current == to_status:
        return dict(row)
    legal = to_status in SHIPMENT_TRANSITIONS.get(current, set())
    if not legal and not allow_correction:
        # Tracking feeds are often out of order.  Older events are retained by
        # add_tracking_event but are not allowed to regress the current state.
        if SHIPMENT_PROGRESS.get(to_status, -1) <= SHIPMENT_PROGRESS.get(current, -1):
            return dict(row)
        raise DomainError(f"非法包裹状态转换: {current} → {to_status}", "illegal_transition")

    delivered_at = utcnow() if to_status == "delivered" else row["delivered_at"]
    shipped_at = row["shipped_at"] or (utcnow() if SHIPMENT_PROGRESS.get(to_status, 0) >= 2 else None)
    conn.execute(
        "UPDATE oms_shipments SET status=?, shipped_at=?, delivered_at=?, updated_at=? WHERE id=?",
        (to_status, shipped_at, delivered_at, utcnow(), shipment_id),
    )
    record_event(
        conn,
        "shipment",
        shipment_id,
        "status_changed",
        from_status=current,
        to_status=to_status,
        actor=actor,
        correlation_id=correlation_id,
        reason=reason,
    )
    return dict(conn.execute("SELECT * FROM oms_shipments WHERE id=?", (shipment_id,)).fetchone())


def _resolve_sku(conn: sqlite3.Connection, site_id: int, item: dict) -> int | None:
    product_id = item.get("product_id")
    variation_id = item.get("variation_id") or 0
    wc_sku = (item.get("sku") or "").strip()
    row = conn.execute(
        '''SELECT sku_id FROM inv_site_sku_map
           WHERE site_id=? AND wc_product_id=? AND wc_variation_id=? AND is_active=1''',
        (site_id, product_id, variation_id),
    ).fetchone()
    if not row and variation_id:
        row = conn.execute(
            '''SELECT sku_id FROM inv_site_sku_map
               WHERE site_id=? AND wc_product_id=? AND wc_variation_id=0 AND is_active=1''',
            (site_id, product_id),
        ).fetchone()
    if not row and wc_sku:
        row = conn.execute(
            '''SELECT sku_id FROM inv_site_sku_map
               WHERE site_id=? AND UPPER(wc_sku)=UPPER(?) AND is_active=1 LIMIT 1''',
            (site_id, wc_sku),
        ).fetchone()
    if not row and wc_sku:
        row = conn.execute(
            "SELECT id AS sku_id FROM inv_skus WHERE is_active=1 AND (UPPER(sku_code)=UPPER(?) OR UPPER(barcode)=UPPER(?)) LIMIT 1",
            (wc_sku, wc_sku),
        ).fetchone()
    return row["sku_id"] if row else None


def sync_order_items(conn: sqlite3.Connection, order_id: str) -> list[dict]:
    order = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
    if not order:
        raise DomainError("订单不存在", "order_not_found")
    site = conn.execute("SELECT id, country FROM sites WHERE url=?", (order["source"],)).fetchone()
    if not site:
        raise DomainError("订单站点不存在", "site_not_found")
    line_items = json_load(order["line_items"], []) or []
    active_keys = []
    for index, item in enumerate(line_items):
        line_id = str(item.get("id") if item.get("id") is not None else f"idx-{index}")
        active_keys.append(line_id)
        qty = int(item.get("quantity") or 0)
        sku_id = _resolve_sku(conn, site["id"], item)
        conn.execute(
            '''INSERT INTO oms_order_items
               (order_id, woo_line_item_id, line_index, wc_product_id,
                wc_variation_id, wc_sku, sku_id, name, ordered_qty, raw_json)
               VALUES (?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(order_id, woo_line_item_id) DO UPDATE SET
                 line_index=excluded.line_index, wc_product_id=excluded.wc_product_id,
                 wc_variation_id=excluded.wc_variation_id, wc_sku=excluded.wc_sku,
                 sku_id=excluded.sku_id, name=excluded.name,
                 ordered_qty=excluded.ordered_qty, raw_json=excluded.raw_json,
                 updated_at=CURRENT_TIMESTAMP''',
            (
                order_id,
                line_id,
                index,
                item.get("product_id"),
                item.get("variation_id") or 0,
                item.get("sku") or "",
                sku_id,
                item.get("name") or item.get("parent_name") or "未命名商品",
                qty,
                json_dump(item),
            ),
        )
    if active_keys:
        placeholders = ",".join("?" for _ in active_keys)
        conn.execute(
            f"DELETE FROM oms_order_items WHERE order_id=? AND woo_line_item_id NOT IN ({placeholders}) AND allocated_qty=0",
            (order_id, *active_keys),
        )
    return [dict(r) for r in conn.execute(
        "SELECT * FROM oms_order_items WHERE order_id=? ORDER BY line_index, id", (order_id,)
    ).fetchall()]


def _candidate_warehouses(conn: sqlite3.Connection, market: str, sku_id: int) -> list[dict]:
    explicit = conn.execute(
        "SELECT warehouse_id FROM oms_sku_warehouses WHERE sku_id=? AND is_enabled=1",
        (sku_id,),
    ).fetchall()
    explicit_ids = {r["warehouse_id"] for r in explicit}

    rows = conn.execute(
        '''SELECT mw.warehouse_id, mw.priority, w.name, w.code, w.country,
                  COALESCE(wi.provider, 'internal') AS provider,
                  COALESCE(wi.inventory_authority, 'local') AS inventory_authority,
                  sw.is_primary, sw.wms_product_name_zh, sw.wms_product_name_en,
                  sw.wms_product_image, sw.product_type,
                  (SELECT MIN(sc.amount) FROM oms_shipping_costs sc
                   WHERE sc.market_code=? AND sc.warehouse_id=mw.warehouse_id
                     AND sc.is_active=1
                     AND (sc.effective_from IS NULL OR sc.effective_from<=date('now'))
                     AND (sc.effective_to IS NULL OR sc.effective_to>=date('now'))) AS shipping_cost
           FROM inv_market_warehouses mw
           JOIN warehouses w ON w.id=mw.warehouse_id AND w.is_active=1
           LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=w.id
           LEFT JOIN oms_sku_warehouses sw ON sw.sku_id=? AND sw.warehouse_id=w.id
           WHERE mw.market_code=? AND mw.is_active=1
             AND COALESCE(wi.is_enabled,1)=1
           ORDER BY mw.priority, mw.id''',
        (market, sku_id, market),
    ).fetchall()
    result = []
    for row in rows:
        d = dict(row)
        if explicit_ids and d["warehouse_id"] not in explicit_ids:
            continue
        if not explicit_ids:
            # A stock record is also a valid implicit product/warehouse map.
            stock_exists = conn.execute(
                "SELECT 1 FROM inv_stock WHERE warehouse_id=? AND sku_id=?",
                (d["warehouse_id"], sku_id),
            ).fetchone()
            ext_exists = conn.execute(
                "SELECT 1 FROM oms_external_stock WHERE warehouse_id=? AND sku_id=?",
                (d["warehouse_id"], sku_id),
            ).fetchone()
            if not stock_exists and not ext_exists:
                continue
        if d["inventory_authority"] == "external_wms":
            stock = conn.execute(
                '''SELECT COALESCE(SUM(available_quantity),0) AS available
                   FROM oms_external_stock WHERE warehouse_id=? AND sku_id=?''',
                (d["warehouse_id"], sku_id),
            ).fetchone()
        else:
            stock = conn.execute(
                '''SELECT COALESCE(on_hand-reserved,0) AS available
                   FROM inv_stock WHERE warehouse_id=? AND sku_id=?''',
                (d["warehouse_id"], sku_id),
            ).fetchone()
        d["available"] = max(0, int(stock["available"] if stock else 0))
        result.append(d)

    def rank(c: dict):
        country = (c.get("country") or "").upper()
        if market == "PL":
            site_preference = 0 if country == "PL" else 1
            return (site_preference, c["priority"], c["warehouse_id"])
        if market == "HU":
            site_preference = 0 if country == "HU" else 1
            return (site_preference, c["priority"], c["warehouse_id"])
        if market == "CZ":
            # Known costs win. Unknown cost falls back to route priority and is
            # explicitly surfaced in plan_reason.
            known = 0 if c.get("shipping_cost") is not None else 1
            cost = float(c["shipping_cost"]) if c.get("shipping_cost") is not None else 0.0
            return (known, cost, c["priority"], c["warehouse_id"])
        return (c["priority"], c["warehouse_id"])

    return sorted(result, key=rank)


def _upsert_order_state(
    conn: sqlite3.Connection,
    order_id: str,
    *,
    revision: int,
    plan_hash: str,
    aggregate_status: str,
    has_shortage: bool,
    manual_review: bool = False,
    manual_reason: str | None = None,
    completion_sync_status: str = "not_ready",
) -> None:
    conn.execute(
        '''INSERT INTO oms_order_fulfillment_state
           (order_id, revision, plan_hash, aggregate_status, has_shortage,
            manual_review, manual_reason, completion_sync_status, last_planned_at)
           VALUES (?,?,?,?,?,?,?,?,?)
           ON CONFLICT(order_id) DO UPDATE SET
             revision=excluded.revision, plan_hash=excluded.plan_hash,
             aggregate_status=excluded.aggregate_status,
             has_shortage=excluded.has_shortage,
             manual_review=excluded.manual_review,
             manual_reason=excluded.manual_reason,
             completion_sync_status=excluded.completion_sync_status,
             last_planned_at=excluded.last_planned_at,
             updated_at=CURRENT_TIMESTAMP''',
        (
            order_id,
            revision,
            plan_hash,
            aggregate_status,
            1 if has_shortage else 0,
            1 if manual_review else 0,
            manual_reason,
            completion_sync_status,
            utcnow(),
        ),
    )


def _money_text(value: Any) -> str:
    try:
        return format(Decimal(str(value or "0")).quantize(Decimal("0.01")), "f")
    except (InvalidOperation, TypeError, ValueError):
        return "0.00"


def _financial_terms_by_warehouse(
    conn: sqlite3.Connection,
    order: sqlite3.Row | dict,
    warehouse_ids,
) -> dict[int, dict]:
    """Assign customer COD ownership without mixing in warehouse fees."""

    ids = sorted({int(warehouse_id) for warehouse_id in warehouse_ids})
    if not ids:
        return {}
    placeholders = ",".join("?" for _ in ids)
    countries = {
        int(row["id"]): (row["country"] or "").upper()
        for row in conn.execute(
            f"SELECT id, country FROM warehouses WHERE id IN ({placeholders})",
            ids,
        ).fetchall()
    }
    assigned_countries = set(countries.values())
    is_cod = str(order["payment_method"] or "").lower() == "cod"
    split_pl_hu = {"PL", "HU"}.issubset(assigned_countries)
    full_cod_amount = _money_text(order["total"])
    currency = order["currency"] or "EUR"
    result = {}
    for warehouse_id in ids:
        country = countries.get(warehouse_id, "")
        if not is_cod:
            role, cod_amount = "not_applicable", "0.00"
        elif split_pl_hu:
            # The Poland parcel collects the complete customer COD.  Hungary
            # receives a shipping instruction only and must never collect COD.
            role = "collector" if country == "PL" else "instruction_only"
            cod_amount = full_cod_amount if country == "PL" else "0.00"
        else:
            role, cod_amount = "collector", full_cod_amount
        result[warehouse_id] = {
            "payment_method": order["payment_method"] or "",
            "cod_collection_role": role,
            "cod_amount": cod_amount,
            "cod_currency": currency,
            "settlement_mode": "monthly_statement" if country == "HU" else "internal",
            "fee_currency": currency,
            "reconciliation_status": "unbilled",
            "notes": (
                "匈牙利仓仓储费与运输费按供应商月结账单核对"
                if country == "HU"
                else "波兰仓负责客户 COD 代收" if role == "collector" else ""
            ),
        }
    return result


def plan_order(
    conn: sqlite3.Connection,
    order_id: str,
    *,
    actor: dict | None = None,
    commit: bool = True,
) -> dict:
    order = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
    if not order:
        raise DomainError("订单不存在", "order_not_found")
    site = conn.execute("SELECT id, country, url FROM sites WHERE url=?", (order["source"],)).fetchone()
    if not site:
        raise DomainError("订单站点不存在", "site_not_found")
    market = (site["country"] or "").upper()
    if market not in {"PL", "CZ", "HU"}:
        raise DomainError(f"站点市场 {market or '未知'} 尚未启用双仓履约", "market_not_enabled")
    if order["status"] not in ACTIVE_ORDER_STATUSES:
        raise DomainError(f"订单状态 {order['status']} 不允许创建履约单", "order_not_plannable")

    items = sync_order_items(conn, order_id)
    old_state = conn.execute(
        "SELECT * FROM oms_order_fulfillment_state WHERE order_id=?", (order_id,)
    ).fetchone()
    old_revision = int(old_state["revision"] if old_state else 0)
    current_fulfillments = conn.execute(
        "SELECT id, status FROM oms_fulfillments WHERE order_id=? AND revision=?",
        (order_id, old_revision),
    ).fetchall() if old_revision else []

    if any(f["status"] not in PLANNABLE_FULFILLMENT_STATUSES for f in current_fulfillments):
        raise DomainError("履约已开始，不能自动重新分仓；请转人工处理", "fulfillment_already_started")

    availability = {}
    assignments: dict[int, list[dict]] = defaultdict(list)
    shortages = []
    plan_reasons = set()
    for item in items:
        qty = max(0, int(item["ordered_qty"] or 0) - int(item["cancelled_qty"] or 0))
        if qty <= 0:
            continue
        if not item["sku_id"]:
            shortages.append({
                "order_item_id": item["id"], "sku_id": None, "name": item["name"],
                "qty": qty, "reason": "sku_unmapped",
            })
            continue
        candidates = _candidate_warehouses(conn, market, item["sku_id"])
        if not candidates:
            shortages.append({
                "order_item_id": item["id"], "sku_id": item["sku_id"], "name": item["name"],
                "qty": qty, "reason": "warehouse_mapping_missing",
            })
            continue
        if market == "CZ":
            plan_reasons.add("shipping_cost" if candidates[0].get("shipping_cost") is not None else "route_priority_fallback")
        else:
            plan_reasons.add("site_preference")
        remaining = qty
        for candidate in candidates:
            key = (candidate["warehouse_id"], item["sku_id"])
            if key not in availability:
                availability[key] = candidate["available"]
            take = min(remaining, availability[key])
            if take <= 0:
                continue
            assignments[candidate["warehouse_id"]].append({
                "order_item_id": item["id"],
                "sku_id": item["sku_id"],
                "qty": take,
                "candidate": candidate,
            })
            availability[key] -= take
            remaining -= take
            if remaining <= 0:
                break
        if remaining:
            shortages.append({
                "order_item_id": item["id"], "sku_id": item["sku_id"], "name": item["name"],
                "qty": remaining, "reason": "out_of_stock",
            })

    financial_terms = _financial_terms_by_warehouse(conn, order, assignments.keys())
    canonical_plan = {
        "order_id": order_id,
        "market": market,
        "assignments": [
            {
                "warehouse_id": warehouse_id,
                "lines": sorted(
                    ({"order_item_id": line["order_item_id"], "sku_id": line["sku_id"], "qty": line["qty"]}
                     for line in lines),
                    key=lambda x: (x["order_item_id"], x["sku_id"]),
                ),
            }
            for warehouse_id, lines in sorted(assignments.items())
        ],
        "shortages": sorted(shortages, key=lambda x: (x["order_item_id"], x["reason"])),
        "reason": sorted(plan_reasons),
        "financial_terms": [
            {"warehouse_id": warehouse_id, **financial_terms[warehouse_id]}
            for warehouse_id in sorted(financial_terms)
        ],
    }
    plan_hash = _hash(canonical_plan)
    if old_state and old_state["plan_hash"] == plan_hash:
        if commit:
            conn.commit()
        return {
            **canonical_plan,
            "revision": old_revision,
            "plan_hash": plan_hash,
            "action": "noop",
        }

    revision = old_revision + 1
    if current_fulfillments:
        for fulfillment in current_fulfillments:
            transition_fulfillment(
                conn,
                fulfillment["id"],
                "superseded",
                actor=actor,
                reason=f"分仓计划更新为 revision {revision}",
            )

    conn.execute(
        "UPDATE oms_order_items SET allocated_qty=0, shortage_qty=0, updated_at=CURRENT_TIMESTAMP WHERE order_id=?",
        (order_id,),
    )
    shortage_by_item = defaultdict(int)
    for shortage in shortages:
        shortage_by_item[shortage["order_item_id"]] += shortage["qty"]
    for item_id, qty in shortage_by_item.items():
        conn.execute("UPDATE oms_order_items SET shortage_qty=? WHERE id=?", (qty, item_id))

    created = []
    has_shortage = bool(shortages)
    for warehouse_id, lines in sorted(assignments.items()):
        integration = conn.execute(
            "SELECT * FROM oms_warehouse_integrations WHERE warehouse_id=?", (warehouse_id,)
        ).fetchone()
        provider = integration["provider"] if integration else "internal"
        mode = "external_wms" if provider == "hungary_wms" else "internal"
        status = "stock_shortage" if has_shortage else ("ready_to_submit" if mode == "external_wms" else "ready_to_pick")
        fid = _uuid()
        invoice_code = build_invoice_code(order, integration, revision) if mode == "external_wms" else None
        idem = f"fulfillment:{order_id}:{warehouse_id}:r{revision}"
        actor_type, actor_id, actor_name = _actor(actor)
        conn.execute(
            '''INSERT INTO oms_fulfillments
               (id, order_id, warehouse_id, revision, status, mode, provider,
                external_warehouse_code, external_invoice_code, channel_code,
                idempotency_key, plan_hash, operator_id, operator_name)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (
                fid,
                order_id,
                warehouse_id,
                revision,
                status,
                mode,
                provider,
                integration["external_code"] if integration else None,
                invoice_code,
                integration["channel_code"] if integration else None,
                idem,
                plan_hash,
                int(actor_id) if actor_id and str(actor_id).isdigit() else None,
                actor_name,
            ),
        )
        finance = financial_terms[warehouse_id]
        conn.execute(
            '''INSERT INTO oms_fulfillment_financials
               (fulfillment_id, payment_method, cod_collection_role, cod_amount,
                cod_currency, settlement_mode, fee_currency,
                reconciliation_status, notes)
               VALUES (?,?,?,?,?,?,?,?,?)
               ON CONFLICT(fulfillment_id) DO UPDATE SET
                 payment_method=excluded.payment_method,
                 cod_collection_role=excluded.cod_collection_role,
                 cod_amount=excluded.cod_amount,
                 cod_currency=excluded.cod_currency,
                 settlement_mode=excluded.settlement_mode,
                 fee_currency=excluded.fee_currency,
                 reconciliation_status=excluded.reconciliation_status,
                 notes=excluded.notes,
                 updated_at=CURRENT_TIMESTAMP''',
            (
                fid,
                finance["payment_method"],
                finance["cod_collection_role"],
                finance["cod_amount"],
                finance["cod_currency"],
                finance["settlement_mode"],
                finance["fee_currency"],
                finance["reconciliation_status"],
                finance["notes"],
            ),
        )
        allocated_by_item = defaultdict(int)
        for line in lines:
            sku = conn.execute(
                "SELECT sku_code, barcode, name FROM inv_skus WHERE id=?", (line["sku_id"],)
            ).fetchone()
            sw = line["candidate"]
            raw = conn.execute("SELECT raw_json FROM oms_order_items WHERE id=?", (line["order_item_id"],)).fetchone()
            conn.execute(
                '''INSERT INTO oms_fulfillment_items
                   (fulfillment_id, order_item_id, sku_id, allocated_qty,
                    sku_code_snapshot, barcode_snapshot, name_snapshot, raw_json)
                   VALUES (?,?,?,?,?,?,?,?)''',
                (
                    fid,
                    line["order_item_id"],
                    line["sku_id"],
                    line["qty"],
                    sku["sku_code"] if sku else None,
                    sku["barcode"] if sku else None,
                    sku["name"] if sku else None,
                    raw["raw_json"] if raw else None,
                ),
            )
            allocated_by_item[line["order_item_id"]] += line["qty"]
        for item_id, qty in allocated_by_item.items():
            conn.execute(
                "UPDATE oms_order_items SET allocated_qty=allocated_qty+?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (qty, item_id),
            )
        record_event(
            conn,
            "fulfillment",
            fid,
            "created",
            to_status=status,
            actor=actor,
            payload={
                "revision": revision,
                "warehouse_id": warehouse_id,
                "financial_terms": finance,
            },
        )
        created.append(fid)
        if mode == "external_wms" and not has_shortage and integration and integration["auto_submit"]:
            enqueue_job(
                conn,
                "SUBMIT_HU_FULFILLMENT",
                "fulfillment",
                fid,
                f"submit:{idem}",
                {"fulfillment_id": fid},
            )

    aggregate = "stock_shortage" if has_shortage else ("allocated" if created else "allocation_blocked")
    manual_reason = "订单存在缺货或未映射商品，需联系客户处理" if has_shortage else None
    _upsert_order_state(
        conn,
        order_id,
        revision=revision,
        plan_hash=plan_hash,
        aggregate_status=aggregate,
        has_shortage=has_shortage,
        manual_review=has_shortage,
        manual_reason=manual_reason,
    )
    record_event(
        conn,
        "order",
        order_id,
        "fulfillment_planned",
        from_status=old_state["aggregate_status"] if old_state else "unallocated",
        to_status=aggregate,
        actor=actor,
        payload=canonical_plan,
    )
    if commit:
        conn.commit()
    return {
        **canonical_plan,
        "revision": revision,
        "plan_hash": plan_hash,
        "fulfillment_ids": created,
        "aggregate_status": aggregate,
        "action": "planned",
    }


def build_invoice_code(order: sqlite3.Row | dict, integration: sqlite3.Row | dict | None, revision: int) -> str:
    raw = str(order["id"])
    safe = re.sub(r"[^A-Za-z0-9]", "", raw)
    external = re.sub(r"[^A-Za-z0-9]", "", (integration["external_code"] if integration else "HU01") or "HU01")
    # Keep it compact for undocumented WMS length limits while remaining
    # deterministic across retries.
    return f"OMS{safe}{external}R{revision}"[:48]


def build_wms_payload(conn: sqlite3.Connection, fulfillment_id: str) -> dict:
    fulfillment = conn.execute(
        '''SELECT f.*, wi.config_json, o.total, o.currency, o.payment_method,
                  o.billing, o.shipping, o.line_items, o.source,
                  w.country AS warehouse_country,
                  ff.cod_collection_role, ff.cod_amount,
                  ff.cod_currency, ff.settlement_mode
           FROM oms_fulfillments f
           JOIN orders o ON o.id=f.order_id
           LEFT JOIN warehouses w ON w.id=f.warehouse_id
           LEFT JOIN oms_warehouse_integrations wi ON wi.warehouse_id=f.warehouse_id
           LEFT JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
           WHERE f.id=?''',
        (fulfillment_id,),
    ).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    if fulfillment["mode"] != "external_wms":
        raise DomainError("该履约单不是外部 WMS 模式", "not_external_wms")
    config = json_load(fulfillment["config_json"], {}) or {}
    shipping = json_load(fulfillment["shipping"], {}) or {}
    billing = json_load(fulfillment["billing"], {}) or {}
    address = shipping if shipping.get("address_1") else billing
    country_code = (address.get("country") or "").upper()
    country_zh = COUNTRY_ZH.get(country_code)
    if not country_zh:
        raise DomainError(f"WMS 不支持或无法转换国家代码: {country_code or '空'}", "country_not_supported")
    email = address.get("email") or billing.get("email")
    phone = address.get("phone") or billing.get("phone")
    if not email:
        raise DomainError("欧洲 WMS 发货要求邮箱，当前订单缺少邮箱", "email_missing")
    if not phone:
        raise DomainError("WMS 发货要求收件人电话", "phone_missing")
    consignee = " ".join(filter(None, [address.get("first_name"), address.get("last_name")])).strip()
    if not consignee:
        raise DomainError("WMS 发货要求收件人姓名", "consignee_missing")

    rows = conn.execute(
        '''SELECT fi.*, sw.wms_product_name_zh, sw.wms_product_name_en,
                  sw.wms_product_image, sw.product_type, oi.raw_json
           FROM oms_fulfillment_items fi
           JOIN oms_order_items oi ON oi.id=fi.order_item_id
           LEFT JOIN oms_sku_warehouses sw
             ON sw.sku_id=fi.sku_id AND sw.warehouse_id=?
           WHERE fi.fulfillment_id=? ORDER BY fi.id''',
        (fulfillment["warehouse_id"], fulfillment_id),
    ).fetchall()
    details = []
    declared_value = Decimal("0")
    for row in rows:
        raw = json_load(row["raw_json"], {}) or {}
        name_zh = (row["wms_product_name_zh"] or "").strip()
        name_en = (row["wms_product_name_en"] or row["name_snapshot"] or "").strip()
        barcode = (row["barcode_snapshot"] or row["sku_code_snapshot"] or "").strip()
        if not name_zh:
            raise DomainError(
                f"SKU {row['sku_code_snapshot'] or row['sku_id']} 缺少 WMS 中文品名",
                "wms_product_name_zh_missing",
            )
        if not name_en:
            raise DomainError("WMS 商品英文规格不能为空", "wms_product_name_en_missing")
        if not barcode:
            raise DomainError("WMS 商品编码不能为空", "wms_barcode_missing")
        details.append({
            "productName": name_zh,
            "productSkuName": name_en,
            "productSkuBarcode": barcode,
            "productSkuImage": row["wms_product_image"] or raw.get("image", {}).get("src") or "",
            "quantity": int(row["allocated_qty"]),
        })
        try:
            line_total = Decimal(str(raw.get("total") or "0"))
            line_qty = Decimal(str(raw.get("quantity") or 1))
            declared_value += (line_total / max(line_qty, Decimal("1"))) * int(row["allocated_qty"])
        except (InvalidOperation, TypeError, ValueError, ZeroDivisionError):
            pass

    is_cod = (fulfillment["payment_method"] or "").lower() == "cod"
    role = fulfillment["cod_collection_role"]
    if not is_cod:
        total = "0.00"
    elif role == "collector":
        total = _money_text(fulfillment["cod_amount"])
    elif role == "instruction_only":
        total = "0.00"
    else:
        # Compatibility fallback for a pre-007 row: a Hungary fulfillment in
        # a Poland+Hungary split must still fail safe to zero COD.
        countries = {
            (row["country"] or "").upper()
            for row in conn.execute(
                '''SELECT w.country
                   FROM oms_fulfillments f
                   LEFT JOIN warehouses w ON w.id=f.warehouse_id
                   WHERE f.order_id=? AND f.revision=? AND f.status!='superseded' ''',
                (fulfillment["order_id"], fulfillment["revision"]),
            ).fetchall()
        }
        if (fulfillment["warehouse_country"] or "").upper() == "HU" and {"PL", "HU"}.issubset(countries):
            total = "0.00"
        else:
            total = _money_text(fulfillment["total"])
    address_1 = (address.get("address_1") or "").strip()
    address_2 = (address.get("address_2") or "").strip()
    if not address_1:
        raise DomainError("WMS 发货要求详细地址", "address_missing")
    product_types = {row["product_type"] for row in rows if row["product_type"]}
    product_type = next(iter(product_types)) if len(product_types) == 1 else (config.get("product_type") or "P")

    return {
        "storehouseCode": fulfillment["external_warehouse_code"] or "HU01",
        "invoiceCode": fulfillment["external_invoice_code"],
        "packType": int(config.get("pack_type", 0)),
        "comments": "",  # Italy forbids comments; the safe cross-country default is blank.
        "invoicePrice": total,
        "currency": fulfillment["currency"] or "EUR",
        "consignee": consignee,
        "tel": phone,
        "contry": country_zh,
        "province": address.get("state") or "",
        "city": address.get("city") or "",
        "district": "",
        "detail": address_1,
        "zipCode": address.get("postcode") or "",
        "email": email,
        "channelCode": fulfillment["channel_code"] or "欧洲直发-25",
        "productType": product_type,
        "addressExt1": address_2,
        "addressExt2": "",
        "ordersDeclaredCurrency": fulfillment["currency"] or "EUR",
        "ordersDeclaredValue": format(declared_value.quantize(Decimal("0.01")), "f"),
        "invoiceDetailsCreateRequests": details,
    }


def create_shipment(
    conn: sqlite3.Connection,
    fulfillment_id: str,
    tracking_number: str,
    *,
    carrier_slug: str = "custom",
    carrier_name: str | None = None,
    label_url: str | None = None,
    external_shipment_id: str | None = None,
    tracking_source: str = "manual",
    actor: dict | None = None,
    initial_status: str = "shipped",
    commit: bool = True,
) -> dict:
    tracking_number = (tracking_number or "").strip()
    if not tracking_number:
        raise DomainError("运单号不能为空", "tracking_missing")
    fulfillment = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fulfillment_id,)).fetchone()
    if not fulfillment:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    existing = conn.execute(
        "SELECT * FROM oms_shipments WHERE carrier_slug=? AND tracking_number=?",
        (carrier_slug, tracking_number),
    ).fetchone()
    if existing:
        if existing["fulfillment_id"] != fulfillment_id:
            raise DomainError("该运单号已属于其他履约单", "tracking_conflict")
        return dict(existing)

    shipment_id = _uuid()
    conn.execute(
        '''INSERT INTO oms_shipments
           (id, fulfillment_id, external_shipment_id, carrier_slug, carrier_name,
            tracking_number, label_url, status, tracking_source, shipped_at)
           VALUES (?,?,?,?,?,?,?,?,?,?)''',
        (
            shipment_id,
            fulfillment_id,
            external_shipment_id,
            carrier_slug or "custom",
            carrier_name or carrier_slug or "WMS动态物流",
            tracking_number,
            label_url,
            initial_status,
            tracking_source,
            utcnow() if SHIPMENT_PROGRESS.get(initial_status, 0) >= 2 else None,
        ),
    )
    items = conn.execute(
        "SELECT * FROM oms_fulfillment_items WHERE fulfillment_id=?", (fulfillment_id,)
    ).fetchall()
    for item in items:
        remaining = max(0, int(item["allocated_qty"]) - int(item["fulfilled_qty"]) - int(item["cancelled_qty"]))
        if remaining and SHIPMENT_PROGRESS.get(initial_status, 0) >= 2:
            conn.execute(
                "INSERT INTO oms_shipment_items (shipment_id, fulfillment_item_id, quantity) VALUES (?,?,?)",
                (shipment_id, item["id"], remaining),
            )
            conn.execute(
                "UPDATE oms_fulfillment_items SET fulfilled_qty=fulfilled_qty+?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (remaining, item["id"]),
            )
    if SHIPMENT_PROGRESS.get(initial_status, 0) >= 2 and fulfillment["status"] != "shipped":
        transition_fulfillment(conn, fulfillment_id, "shipped", actor=actor, reason="已生成运单")
    record_event(
        conn,
        "shipment",
        shipment_id,
        "created",
        to_status=initial_status,
        actor=actor,
        payload={"tracking_number": tracking_number, "carrier_slug": carrier_slug},
    )
    conn.execute(
        '''INSERT INTO oms_shipment_notifications
           (shipment_id, channel, template_version, status)
           VALUES (?, 'email', 'ast-v1', 'pending')
           ON CONFLICT(shipment_id, channel, template_version) DO NOTHING''',
        (shipment_id,),
    )
    if SHIPMENT_PROGRESS.get(initial_status, 0) >= 2:
        _enqueue_shipment_side_effects(conn, fulfillment, shipment_id, "created")
    if commit:
        conn.commit()
    return dict(conn.execute("SELECT * FROM oms_shipments WHERE id=?", (shipment_id,)).fetchone())


def _enqueue_shipment_side_effects(conn, fulfillment, shipment_id: str, suffix: str) -> None:
    enqueue_job(
        conn,
        "SYNC_SHIPMENT_TO_WOOCOMMERCE",
        "shipment",
        shipment_id,
        f"woo-shipment:{shipment_id}:v1",
        {"shipment_id": shipment_id},
    )
    enqueue_job(
        conn,
        "POLL_SHIPMENT_TRACKING",
        "shipment",
        shipment_id,
        f"track:{shipment_id}:initial",
        {"shipment_id": shipment_id},
    )
    enqueue_job(
        conn,
        "RECOMPUTE_ORDER_STATUS",
        "order",
        fulfillment["order_id"],
        f"recompute:{fulfillment['order_id']}:{shipment_id}:{suffix}",
        {"order_id": fulfillment["order_id"]},
    )


def mark_shipment_shipped(
    conn: sqlite3.Connection,
    shipment_id: str,
    *,
    actor: dict | None = None,
    reason: str = "仓库确认出库",
    commit: bool = True,
) -> dict:
    shipment = conn.execute("SELECT * FROM oms_shipments WHERE id=?", (shipment_id,)).fetchone()
    if not shipment:
        raise DomainError("包裹不存在", "shipment_not_found")
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (shipment["fulfillment_id"],)
    ).fetchone()
    if SHIPMENT_PROGRESS.get(shipment["status"], 0) < 2:
        transition_shipment(conn, shipment_id, "shipped", actor=actor, reason=reason)
    existing_items = conn.execute(
        "SELECT COUNT(*) AS n FROM oms_shipment_items WHERE shipment_id=?", (shipment_id,)
    ).fetchone()["n"]
    if not existing_items:
        items = conn.execute(
            "SELECT * FROM oms_fulfillment_items WHERE fulfillment_id=?", (fulfillment["id"],)
        ).fetchall()
        for item in items:
            remaining = max(0, int(item["allocated_qty"]) - int(item["fulfilled_qty"]) - int(item["cancelled_qty"]))
            if remaining:
                conn.execute(
                    "INSERT INTO oms_shipment_items (shipment_id, fulfillment_item_id, quantity) VALUES (?,?,?)",
                    (shipment_id, item["id"], remaining),
                )
                conn.execute(
                    "UPDATE oms_fulfillment_items SET fulfilled_qty=fulfilled_qty+?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                    (remaining, item["id"]),
                )
    if fulfillment["status"] != "shipped":
        transition_fulfillment(conn, fulfillment["id"], "shipped", actor=actor, reason=reason)
    _enqueue_shipment_side_effects(conn, fulfillment, shipment_id, "shipped")
    if commit:
        conn.commit()
    return dict(conn.execute("SELECT * FROM oms_shipments WHERE id=?", (shipment_id,)).fetchone())


def add_tracking_event(
    conn: sqlite3.Connection,
    shipment_id: str,
    provider: str,
    normalized_status: str,
    *,
    raw_status: str | None = None,
    event_at: str | None = None,
    location: str | None = None,
    description: str | None = None,
    external_event_id: str | None = None,
    raw_payload: Any = None,
    correlation_id: str | None = None,
    commit: bool = True,
) -> dict:
    shipment = conn.execute("SELECT * FROM oms_shipments WHERE id=?", (shipment_id,)).fetchone()
    if not shipment:
        raise DomainError("包裹不存在", "shipment_not_found")
    canonical = {
        "provider": provider,
        "raw_status": raw_status,
        "normalized_status": normalized_status,
        "event_at": event_at,
        "location": location,
        "description": description,
    }
    fingerprint = external_event_id or _hash(canonical)
    payload_hash = _hash(raw_payload) if raw_payload is not None else None
    conn.execute(
        '''INSERT INTO oms_tracking_events
           (shipment_id, provider, external_event_id, event_fingerprint,
            raw_status, normalized_status, event_at, location, description,
            payload_hash, raw_payload)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)
           ON CONFLICT(shipment_id, event_fingerprint) DO NOTHING''',
        (
            shipment_id,
            provider,
            external_event_id,
            fingerprint,
            raw_status,
            normalized_status,
            event_at,
            location,
            description,
            payload_hash,
            json_dump(raw_payload) if raw_payload is not None else None,
        ),
    )
    before = shipment["status"]
    updated = transition_shipment(
        conn,
        shipment_id,
        normalized_status,
        reason=f"{provider}:{raw_status or normalized_status}",
        correlation_id=correlation_id,
    )
    fulfillment = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE id=?", (shipment["fulfillment_id"],)
    ).fetchone()
    if updated["status"] == "delivered" and fulfillment["status"] == "shipped":
        open_shipments = conn.execute(
            "SELECT COUNT(*) AS n FROM oms_shipments WHERE fulfillment_id=? AND status NOT IN ('delivered','cancelled')",
            (fulfillment["id"],),
        ).fetchone()["n"]
        if open_shipments == 0:
            transition_fulfillment(
                conn,
                fulfillment["id"],
                "delivered",
                reason="该履约单全部包裹妥投",
                correlation_id=correlation_id,
            )
    if before != updated["status"]:
        enqueue_job(
            conn,
            "RECOMPUTE_ORDER_STATUS",
            "order",
            fulfillment["order_id"],
            f"recompute:{fulfillment['order_id']}:{shipment_id}:{updated['status']}:{fingerprint[:12]}",
            {"order_id": fulfillment["order_id"]},
        )
    if commit:
        conn.commit()
    return updated


def recompute_order_status(
    conn: sqlite3.Connection,
    order_id: str,
    *,
    actor: dict | None = None,
    commit: bool = True,
) -> dict:
    state = conn.execute(
        "SELECT * FROM oms_order_fulfillment_state WHERE order_id=?", (order_id,)
    ).fetchone()
    if not state:
        raise DomainError("订单尚未建立履约计划", "order_not_planned")
    revision = state["revision"]
    fulfillments = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE order_id=? AND revision=? AND status!='superseded'",
        (order_id, revision),
    ).fetchall()
    if not fulfillments:
        aggregate = "allocation_blocked"
    elif state["manual_review"] or state["has_shortage"]:
        aggregate = "manual_review" if state["manual_review"] else "stock_shortage"
    else:
        shipments = conn.execute(
            '''SELECT s.* FROM oms_shipments s JOIN oms_fulfillments f ON f.id=s.fulfillment_id
               WHERE f.order_id=? AND f.revision=? AND s.status!='cancelled' ''',
            (order_id, revision),
        ).fetchall()
        shipped_count = sum(1 for s in shipments if SHIPMENT_PROGRESS.get(s["status"], 0) >= 2)
        delivered_count = sum(1 for s in shipments if s["status"] == "delivered")
        active_fulfillments = [f for f in fulfillments if f["status"] != "cancelled"]
        if shipments and delivered_count == len(shipments) and all(f["status"] == "delivered" for f in active_fulfillments):
            aggregate = "delivered"
        elif delivered_count:
            aggregate = "partially_delivered"
        elif shipments and shipped_count == len(shipments) and all(
            f["status"] in {"shipped", "delivered"} for f in active_fulfillments
        ):
            aggregate = "shipped"
        elif shipped_count:
            aggregate = "partially_shipped"
        elif any(f["status"] in {"picking", "packed", "accepted", "submitting"} for f in active_fulfillments):
            aggregate = "fulfillment_in_progress"
        else:
            aggregate = "allocated"

    old = state["aggregate_status"]
    completion_status = state["completion_sync_status"]
    if aggregate == "delivered" and completion_status not in {"pending", "synced"}:
        completion_status = "pending"
        enqueue_job(
            conn,
            "COMPLETE_WOOCOMMERCE_ORDER",
            "order",
            order_id,
            f"woo-complete:{order_id}:r{revision}",
            {"order_id": order_id, "revision": revision},
        )
    conn.execute(
        '''UPDATE oms_order_fulfillment_state
           SET aggregate_status=?, completion_sync_status=?, updated_at=CURRENT_TIMESTAMP
           WHERE order_id=?''',
        (aggregate, completion_status, order_id),
    )
    if old != aggregate:
        record_event(
            conn,
            "order",
            order_id,
            "aggregate_status_changed",
            from_status=old,
            to_status=aggregate,
            actor=actor,
        )
    if commit:
        conn.commit()
    return {
        "order_id": order_id,
        "revision": revision,
        "from_status": old,
        "aggregate_status": aggregate,
        "completion_sync_status": completion_status,
    }


def mark_manual_review(
    conn: sqlite3.Connection,
    order_id: str,
    reason: str,
    *,
    actor: dict | None = None,
    commit: bool = True,
) -> None:
    state = conn.execute(
        "SELECT * FROM oms_order_fulfillment_state WHERE order_id=?", (order_id,)
    ).fetchone()
    if not state:
        _upsert_order_state(
            conn,
            order_id,
            revision=0,
            plan_hash="",
            aggregate_status="manual_review",
            has_shortage=False,
            manual_review=True,
            manual_reason=reason,
        )
    else:
        conn.execute(
            '''UPDATE oms_order_fulfillment_state
               SET aggregate_status='manual_review', manual_review=1,
                   manual_reason=?, updated_at=CURRENT_TIMESTAMP WHERE order_id=?''',
            (reason, order_id),
        )
    record_event(
        conn,
        "order",
        order_id,
        "manual_review_required",
        from_status=state["aggregate_status"] if state else None,
        to_status="manual_review",
        actor=actor,
        reason=reason,
    )
    if commit:
        conn.commit()


def completion_guard(conn: sqlite3.Connection, order_id: str) -> tuple[bool, str | None]:
    """Return whether a new-domain order may be marked completed.

    Legacy orders without an OMS state remain governed by the old workflow.
    """

    state = conn.execute(
        "SELECT aggregate_status, completion_sync_status FROM oms_order_fulfillment_state WHERE order_id=?",
        (order_id,),
    ).fetchone()
    if not state:
        return True, None
    if state["aggregate_status"] != "delivered":
        return False, f"所有包裹尚未妥投（当前履约状态: {state['aggregate_status']}）"
    return True, None
