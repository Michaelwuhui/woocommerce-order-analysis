"""Flask UI/API for multi-warehouse fulfillment operations."""

from __future__ import annotations

import hashlib
import hmac
import os
from functools import wraps

from flask import Blueprint, jsonify, render_template, request
from flask_login import current_user, login_required

from fulfillment_common import get_conn, json_dump, json_load, utcnow
from fulfillment_service import (
    DomainError,
    create_shipment,
    enqueue_job,
    mark_manual_review,
    plan_order,
    recompute_order_status,
    transition_fulfillment,
)


fulfillment_bp = Blueprint("fulfillment", __name__)


def _callback_value(value, keys):
    if isinstance(value, dict):
        for key, item in value.items():
            if str(key).lower() in keys and item not in (None, ""):
                return str(item)
        for item in value.values():
            found = _callback_value(item, keys)
            if found:
                return found
    elif isinstance(value, list):
        for item in value:
            found = _callback_value(item, keys)
            if found:
                return found
    return None


@fulfillment_bp.route("/api/fulfillment/webhook/wms", methods=["POST"])
def wms_webhook():
    """Authenticated, deduplicated callback inbox.

    The supplier has not documented a callback payload contract.  Therefore a
    callback never mutates state directly; it only schedules an authoritative
    status/tracking query using any invoice/tracking identifier it contains.
    """
    expected = os.environ.get("WMS_WEBHOOK_TOKEN", "")
    supplied = request.headers.get("X-WMS-Webhook-Token", "")
    if not supplied and request.headers.get("Authorization", "").lower().startswith("bearer "):
        supplied = request.headers.get("Authorization", "")[7:]
    if not expected:
        return jsonify({"error": "WMS 回调未启用"}), 503
    if not hmac.compare_digest(expected, supplied):
        return jsonify({"error": "回调认证失败"}), 401
    body = request.get_json(silent=True)
    if body is None:
        return jsonify({"error": "回调必须是 JSON"}), 400
    raw = json_dump(body)
    payload_hash = hashlib.sha256(raw.encode("utf-8")).hexdigest()
    event_id = request.headers.get("X-Event-Id") or _callback_value(
        body, {"eventid", "event_id", "callbackid", "callback_id"}
    )
    conn = get_conn()
    try:
        cur = conn.execute(
            '''INSERT INTO oms_webhook_inbox
               (provider, external_event_id, payload_hash, headers_json, payload_json, status)
               VALUES ('hungary_wms',?,?,?,?, 'received')
               ON CONFLICT DO NOTHING''',
            (
                event_id,
                payload_hash,
                json_dump({"user_agent": request.headers.get("User-Agent"), "event_id": event_id}),
                raw,
            ),
        )
        if not cur.rowcount:
            return jsonify({"accepted": True, "duplicate": True})
        inbox_id = cur.lastrowid
        invoice = _callback_value(body, {"invoicecode", "invoice_code"})
        tracking = _callback_value(
            body, {"expresscode", "express_code", "trackingnumber", "tracking_number"}
        )
        queued = []
        if invoice:
            ff = conn.execute(
                "SELECT id FROM oms_fulfillments WHERE provider='hungary_wms' AND external_invoice_code=?",
                (invoice,),
            ).fetchone()
            if ff:
                queued.append(enqueue_job(
                    conn, "POLL_WMS_STATUS", "fulfillment", ff["id"],
                    f"wms-callback:{inbox_id}:status", {"fulfillment_id": ff["id"]},
                ))
        if tracking:
            shipment = conn.execute(
                "SELECT id FROM oms_shipments WHERE tracking_number=?", (tracking,)
            ).fetchone()
            if shipment:
                queued.append(enqueue_job(
                    conn, "POLL_SHIPMENT_TRACKING", "shipment", shipment["id"],
                    f"wms-callback:{inbox_id}:tracking", {"shipment_id": shipment["id"]},
                ))
        conn.execute(
            "UPDATE oms_webhook_inbox SET status=?, processed_at=CURRENT_TIMESTAMP WHERE id=?",
            ("queued" if queued else "unmatched", inbox_id),
        )
        conn.commit()
        return jsonify({"accepted": True, "duplicate": False, "queued": len(queued)})
    finally:
        conn.close()


def _is_super_admin() -> bool:
    return getattr(current_user, "username", None) == "admin"


def _can_manage_inventory() -> bool:
    if _is_super_admin():
        return True
    try:
        from inv_common import can_manage_inventory
        return bool(can_manage_inventory())
    except Exception:
        return False


def _allowed_warehouse_ids(capability: str = "can_view") -> list[int] | None:
    if _is_super_admin() or _can_manage_inventory():
        return None
    allowed = set()
    conn = get_conn()
    try:
        valid = {
            "can_view", "can_pick", "can_pack", "can_ship", "can_cancel",
            "can_retry", "can_reconcile",
        }
        column = capability if capability in valid else "can_view"
        rows = conn.execute(
            f"SELECT warehouse_id FROM oms_warehouse_user_permissions WHERE user_id=? AND {column}=1",
            (current_user.id,),
        ).fetchall()
        allowed.update(r["warehouse_id"] for r in rows)
    finally:
        conn.close()
    # Existing partner assignments remain a safe compatibility source for view
    # scope until every warehouse operator has an explicit V2 permission row.
    if capability == "can_view":
        try:
            from inv_common import partner_warehouse_ids
            allowed.update(partner_warehouse_ids())
        except Exception:
            pass
    return sorted(allowed)


def _has_any_access() -> bool:
    ids = _allowed_warehouse_ids("can_view")
    return ids is None or bool(ids)


def fulfillment_view_required(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not getattr(current_user, "is_authenticated", False):
            return jsonify({"error": "请先登录"}), 401
        if not _has_any_access():
            return jsonify({"error": "没有仓库履约查看权限"}), 403
        return func(*args, **kwargs)
    return wrapper


def _check_fulfillment_permission(conn, fulfillment_id: str, capability: str):
    row = conn.execute(
        "SELECT warehouse_id FROM oms_fulfillments WHERE id=?", (fulfillment_id,)
    ).fetchone()
    if not row:
        raise DomainError("履约单不存在", "fulfillment_not_found")
    allowed = _allowed_warehouse_ids(capability)
    if allowed is not None and row["warehouse_id"] not in allowed:
        raise DomainError("没有该仓库的操作权限", "warehouse_forbidden")
    return row["warehouse_id"]


def _actor():
    return {
        "type": "user",
        "id": getattr(current_user, "id", None),
        "name": getattr(current_user, "name", None) or getattr(current_user, "username", None),
    }


def _rollout_readiness(conn) -> dict:
    counts = {
        "sku_count": conn.execute(
            "SELECT COUNT(*) AS n FROM inv_skus WHERE is_active=1"
        ).fetchone()["n"],
        "site_sku_map_count": conn.execute(
            "SELECT COUNT(*) AS n FROM inv_site_sku_map WHERE is_active=1"
        ).fetchone()["n"],
        "warehouse_sku_map_count": conn.execute(
            "SELECT COUNT(*) AS n FROM oms_sku_warehouses WHERE is_enabled=1"
        ).fetchone()["n"],
        "wms_ready_sku_count": conn.execute(
            '''SELECT COUNT(*) AS n
               FROM oms_sku_warehouses sw
               JOIN inv_skus s ON s.id=sw.sku_id
               JOIN oms_warehouse_integrations wi ON wi.warehouse_id=sw.warehouse_id
               WHERE sw.is_enabled=1 AND s.is_active=1
                 AND wi.provider='hungary_wms' AND wi.is_enabled=1
                 AND trim(COALESCE(s.barcode,''))!=''
                 AND trim(COALESCE(sw.wms_product_name_zh,''))!=''
                 AND trim(COALESCE(sw.wms_product_name_en,''))!='' '''
        ).fetchone()["n"],
        "wms_stocked_sku_count": conn.execute(
            '''SELECT COUNT(DISTINCT es.sku_id) AS n
               FROM oms_external_stock es
               JOIN oms_warehouse_integrations wi ON wi.warehouse_id=es.warehouse_id
               WHERE wi.provider='hungary_wms' AND wi.is_enabled=1
                 AND es.sku_id IS NOT NULL AND es.available_quantity>0'''
        ).fetchone()["n"],
    }
    can_auto_plan = all(
        counts[key] > 0
        for key in ("sku_count", "site_sku_map_count", "warehouse_sku_map_count")
    )
    can_auto_submit = (
        can_auto_plan
        and counts["wms_ready_sku_count"] > 0
        and counts["wms_stocked_sku_count"] > 0
    )
    blockers = []
    if counts["sku_count"] <= 0:
        blockers.append("SKU 主档为空")
    if counts["site_sku_map_count"] <= 0:
        blockers.append("站点商品与 SKU 映射为空")
    if counts["warehouse_sku_map_count"] <= 0:
        blockers.append("SKU 仓库归属为空")
    if counts["wms_ready_sku_count"] <= 0:
        blockers.append("没有同时具备条码、中英文品名的匈牙利 WMS SKU")
    if counts["wms_stocked_sku_count"] <= 0:
        blockers.append("HU01 当前没有可识别的可用库存")
    return {
        **counts,
        "can_auto_plan": can_auto_plan,
        "can_auto_submit": can_auto_submit,
        "blockers": blockers,
    }


def _domain_error(exc: DomainError):
    status = 403 if exc.code == "warehouse_forbidden" else 404 if exc.code.endswith("not_found") else 409
    return jsonify({"error": str(exc), "code": exc.code}), status


@fulfillment_bp.route("/fulfillment")
@login_required
@fulfillment_view_required
def fulfillment_page():
    return render_template(
        "fulfillment.html",
        can_manage_fulfillment=_can_manage_inventory(),
        allowed_warehouses=_allowed_warehouse_ids("can_view"),
    )


@fulfillment_bp.route("/api/fulfillment/orders")
@login_required
@fulfillment_view_required
def list_fulfillment_orders():
    conn = get_conn()
    try:
        allowed = _allowed_warehouse_ids("can_view")
        params = []
        where = ["f.status!='superseded'"]
        if allowed is not None:
            if not allowed:
                return jsonify({"summary": {}, "items": []})
            where.append(f"f.warehouse_id IN ({','.join('?' for _ in allowed)})")
            params.extend(allowed)
        status = (request.args.get("status") or "").strip()
        if status:
            where.append("f.status=?")
            params.append(status)
        warehouse_id = request.args.get("warehouse_id", type=int)
        if warehouse_id:
            where.append("f.warehouse_id=?")
            params.append(warehouse_id)
        search = (request.args.get("search") or "").strip()
        if search:
            where.append("(o.number LIKE ? OR f.external_invoice_code LIKE ? OR f.external_pick_code LIKE ?)")
            params.extend([f"%{search}%"] * 3)
        try:
            limit = min(500, max(1, int(request.args.get("limit") or 200)))
        except ValueError:
            limit = 200
        rows = conn.execute(
            f'''SELECT f.*, w.name AS warehouse_name, w.country AS warehouse_country,
                       ff.cod_collection_role, ff.cod_amount, ff.cod_currency,
                       ff.merchandise_amount, ff.customer_shipping_amount,
                       ff.order_adjustment_amount, ff.source_order_total,
                       ff.source_shipping_total, ff.allocation_method,
                       ff.settlement_mode, ff.reconciliation_status,
                       o.number AS order_number, o.date_created AS order_date,
                       o.source, s.country AS market_code,
                       ofs.aggregate_status, ofs.has_shortage, ofs.manual_review,
                       ofs.manual_reason,
                       (SELECT COUNT(*) FROM oms_shipments sh WHERE sh.fulfillment_id=f.id) AS parcel_count
                FROM oms_fulfillments f
                JOIN orders o ON o.id=f.order_id
                LEFT JOIN sites s ON s.url=o.source
                LEFT JOIN warehouses w ON w.id=f.warehouse_id
                LEFT JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
                LEFT JOIN oms_order_fulfillment_state ofs ON ofs.order_id=f.order_id
                WHERE {' AND '.join(where)}
                ORDER BY ofs.manual_review DESC, ofs.has_shortage DESC,
                         o.date_created DESC LIMIT ?''',
            (*params, limit),
        ).fetchall()
        result = []
        for row in rows:
            item = dict(row)
            item["items"] = [dict(x) for x in conn.execute(
                '''SELECT fi.id, fi.allocated_qty, fi.fulfilled_qty, fi.cancelled_qty,
                          fi.sku_code_snapshot, fi.name_snapshot, oi.name AS order_item_name
                   FROM oms_fulfillment_items fi
                   JOIN oms_order_items oi ON oi.id=fi.order_item_id
                   WHERE fi.fulfillment_id=? ORDER BY fi.id''',
                (row["id"],),
            ).fetchall()]
            result.append(item)
        summary_where = list(where)
        summary_params = list(params)
        # Search/status filters intentionally affect the visible summary.
        summary = {
            r["status"]: r["n"]
            for r in conn.execute(
                f"SELECT f.status, COUNT(*) AS n FROM oms_fulfillments f JOIN orders o ON o.id=f.order_id WHERE {' AND '.join(summary_where)} GROUP BY f.status",
                summary_params,
            ).fetchall()
        }
        warehouses = [dict(r) for r in conn.execute(
            "SELECT id, name, country FROM warehouses WHERE is_active=1 ORDER BY country, id"
        ).fetchall()]
        if allowed is not None:
            warehouses = [w for w in warehouses if w["id"] in allowed]
        return jsonify({"summary": summary, "items": result, "warehouses": warehouses})
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/order/<path:order_id>")
@login_required
@fulfillment_view_required
def fulfillment_order_detail(order_id):
    conn = get_conn()
    try:
        allowed = _allowed_warehouse_ids("can_view")
        state = conn.execute(
            "SELECT * FROM oms_order_fulfillment_state WHERE order_id=?", (order_id,)
        ).fetchone()
        if not state:
            return jsonify({"error": "订单尚未建立履约计划"}), 404
        query = '''SELECT f.*, w.name AS warehouse_name, w.country AS warehouse_country,
                          ff.cod_collection_role, ff.cod_amount, ff.cod_currency,
                          ff.merchandise_amount, ff.customer_shipping_amount,
                          ff.order_adjustment_amount, ff.source_order_total,
                          ff.source_shipping_total, ff.allocation_method,
                          ff.settlement_mode, ff.statement_month,
                          ff.warehouse_storage_fee, ff.warehouse_shipping_fee,
                          ff.fee_currency, ff.reconciliation_status
                   FROM oms_fulfillments f LEFT JOIN warehouses w ON w.id=f.warehouse_id
                   LEFT JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
                   WHERE f.order_id=? AND f.revision=? AND f.status!='superseded' '''
        params = [order_id, state["revision"]]
        if allowed is not None:
            if not allowed:
                return jsonify({"error": "没有该订单的仓库权限"}), 403
            query += f" AND f.warehouse_id IN ({','.join('?' for _ in allowed)})"
            params.extend(allowed)
        fulfillments = conn.execute(query + " ORDER BY f.warehouse_id", params).fetchall()
        if not fulfillments:
            return jsonify({"error": "没有该订单的仓库权限"}), 403
        output = []
        for fulfillment in fulfillments:
            data = dict(fulfillment)
            data["items"] = [dict(r) for r in conn.execute(
                '''SELECT fi.*, oi.name AS order_item_name, oi.ordered_qty,
                          oi.shortage_qty, oi.cancelled_qty AS order_cancelled_qty
                   FROM oms_fulfillment_items fi JOIN oms_order_items oi ON oi.id=fi.order_item_id
                   WHERE fi.fulfillment_id=? ORDER BY fi.id''',
                (fulfillment["id"],),
            ).fetchall()]
            shipments = []
            for shipment in conn.execute(
                "SELECT * FROM oms_shipments WHERE fulfillment_id=? ORDER BY created_at",
                (fulfillment["id"],),
            ).fetchall():
                parcel = dict(shipment)
                parcel["events"] = [dict(r) for r in conn.execute(
                    '''SELECT provider, raw_status, normalized_status, event_at,
                              received_at, location, description
                       FROM oms_tracking_events WHERE shipment_id=?
                       ORDER BY COALESCE(event_at, received_at) DESC LIMIT 50''',
                    (shipment["id"],),
                ).fetchall()]
                shipments.append(parcel)
            data["shipments"] = shipments
            output.append(data)
        order = conn.execute(
            "SELECT id, number, status, source, date_created FROM orders WHERE id=?", (order_id,)
        ).fetchone()
        return jsonify({"order": dict(order), "state": dict(state), "fulfillments": output})
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/plan/<path:order_id>", methods=["POST"])
@login_required
def api_plan_order(order_id):
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        return jsonify(plan_order(conn, order_id, actor=_actor()))
    except DomainError as exc:
        conn.rollback()
        return _domain_error(exc)
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/<fulfillment_id>/transition", methods=["POST"])
@login_required
def api_transition_fulfillment(fulfillment_id):
    body = request.get_json(silent=True) or {}
    to_status = body.get("status")
    capability = {
        "picking": "can_pick",
        "packed": "can_pack",
        "cancel_pending": "can_cancel",
        "cancelled": "can_cancel",
    }.get(to_status, "can_ship")
    conn = get_conn()
    try:
        _check_fulfillment_permission(conn, fulfillment_id, capability)
        result = transition_fulfillment(
            conn, fulfillment_id, to_status, actor=_actor(), reason=body.get("reason")
        )
        fulfillment = conn.execute(
            "SELECT order_id FROM oms_fulfillments WHERE id=?", (fulfillment_id,)
        ).fetchone()
        recompute_order_status(conn, fulfillment["order_id"], actor=_actor(), commit=False)
        conn.commit()
        return jsonify(result)
    except DomainError as exc:
        conn.rollback()
        return _domain_error(exc)
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/<fulfillment_id>/shipment", methods=["POST"])
@login_required
def api_create_shipment(fulfillment_id):
    body = request.get_json(silent=True) or {}
    conn = get_conn()
    try:
        _check_fulfillment_permission(conn, fulfillment_id, "can_ship")
        shipment = create_shipment(
            conn,
            fulfillment_id,
            body.get("tracking_number"),
            carrier_slug=(body.get("carrier_slug") or "custom").strip(),
            carrier_name=body.get("carrier_name"),
            label_url=body.get("label_url"),
            tracking_source="internal_warehouse",
            actor=_actor(),
        )
        return jsonify(shipment)
    except DomainError as exc:
        conn.rollback()
        return _domain_error(exc)
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/<fulfillment_id>/submit-wms", methods=["POST"])
@login_required
def api_submit_wms(fulfillment_id):
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        fulfillment = conn.execute("SELECT * FROM oms_fulfillments WHERE id=?", (fulfillment_id,)).fetchone()
        if not fulfillment:
            return jsonify({"error": "履约单不存在"}), 404
        if fulfillment["mode"] != "external_wms":
            return jsonify({"error": "该履约单不是外部 WMS 模式"}), 409
        job_id = enqueue_job(
            conn,
            "SUBMIT_HU_FULFILLMENT",
            "fulfillment",
            fulfillment_id,
            f"manual-submit:{fulfillment['idempotency_key']}",
            {"fulfillment_id": fulfillment_id},
        )
        conn.commit()
        return jsonify({"queued": True, "job_id": job_id})
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/<fulfillment_id>/cancel", methods=["POST"])
@login_required
def api_cancel_fulfillment(fulfillment_id):
    body = request.get_json(silent=True) or {}
    reason = (body.get("reason") or "人工取消履约").strip()
    conn = get_conn()
    try:
        _check_fulfillment_permission(conn, fulfillment_id, "can_cancel")
        fulfillment = conn.execute(
            "SELECT * FROM oms_fulfillments WHERE id=?", (fulfillment_id,)
        ).fetchone()
        shipped_elsewhere = conn.execute(
            '''SELECT COUNT(*) AS n FROM oms_fulfillments
               WHERE order_id=? AND id!=? AND status IN ('shipped','delivered')''',
            (fulfillment["order_id"], fulfillment_id),
        ).fetchone()["n"]
        if shipped_elsewhere or fulfillment["status"] in {"shipped", "delivered"}:
            mark_manual_review(
                conn, fulfillment["order_id"],
                "一个仓库已经发货，另一仓取消必须人工联系客户并处理", actor=_actor(),
            )
            return jsonify({"error": "已有仓库发货，订单已标记人工处理"}), 409
        if fulfillment["status"] in {"submitting", "submission_unknown"}:
            mark_manual_review(
                conn, fulfillment["order_id"],
                "WMS 提交结果尚未确认，取消前必须先核验外部单据", actor=_actor(),
            )
            return jsonify({"error": "WMS 提交结果尚未确认；已转人工处理"}), 409
        if fulfillment["mode"] == "external_wms" and fulfillment["submitted_at"]:
            manual_reason = (
                "匈牙利 WMS 不支持 API 取消；请在物流群联系对方运营人工拦截。"
                f"申请原因：{reason}"
            )
            mark_manual_review(
                conn, fulfillment["order_id"], manual_reason, actor=_actor()
            )
            return jsonify({
                "manual_required": True,
                "message": "已标记人工处理；请在物流群联系对方运营拦截该出库单",
            }), 202
        # Not yet sent to an external warehouse: local cancellation is final.
        transition_fulfillment(
            conn, fulfillment_id, "cancelled", actor=_actor(), reason=reason
        )
        recompute_order_status(conn, fulfillment["order_id"], actor=_actor(), commit=False)
        conn.commit()
        return jsonify({"cancelled": True})
    except DomainError as exc:
        conn.rollback()
        return _domain_error(exc)
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/order/<path:order_id>/manual-review", methods=["POST"])
@login_required
def api_manual_review(order_id):
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    body = request.get_json(silent=True) or {}
    reason = (body.get("reason") or "人工处理").strip()
    conn = get_conn()
    try:
        mark_manual_review(conn, order_id, reason, actor=_actor())
        return jsonify({"success": True, "order_id": order_id, "reason": reason})
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/order/<path:order_id>/resolve-shortage", methods=["POST"])
@login_required
def api_resolve_shortage(order_id):
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    body = request.get_json(silent=True) or {}
    cancellations = body.get("cancellations") or []
    conn = get_conn()
    try:
        started = conn.execute(
            '''SELECT COUNT(*) AS n FROM oms_fulfillments
               WHERE order_id=? AND status NOT IN
                 ('planned','ready_to_pick','ready_to_submit','stock_shortage','manual_hold','superseded','cancelled')''',
            (order_id,),
        ).fetchone()["n"]
        if started:
            mark_manual_review(conn, order_id, "已有仓库开始履约，缺货/部分取消必须人工处理", actor=_actor())
            return jsonify({"error": "已有仓库开始履约，不能自动重分仓；已保持人工处理标记"}), 409
        for item in cancellations:
            item_id = int(item.get("order_item_id"))
            qty = max(0, int(item.get("quantity") or 0))
            row = conn.execute(
                "SELECT ordered_qty FROM oms_order_items WHERE id=? AND order_id=?", (item_id, order_id)
            ).fetchone()
            if not row:
                raise DomainError(f"订单商品 {item_id} 不存在", "order_item_not_found")
            if qty > int(row["ordered_qty"]):
                raise DomainError("取消数量不能超过下单数量", "invalid_cancel_qty")
            conn.execute(
                "UPDATE oms_order_items SET cancelled_qty=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                (qty, item_id),
            )
        conn.execute(
            '''UPDATE oms_order_fulfillment_state
               SET manual_review=0, manual_reason=NULL, has_shortage=0,
                   updated_at=CURRENT_TIMESTAMP WHERE order_id=?''',
            (order_id,),
        )
        result = plan_order(conn, order_id, actor=_actor(), commit=False)
        conn.commit()
        return jsonify(result)
    except DomainError as exc:
        conn.rollback()
        return _domain_error(exc)
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/shipment/<shipment_id>/track", methods=["POST"])
@login_required
@fulfillment_view_required
def api_refresh_tracking(shipment_id):
    conn = get_conn()
    try:
        row = conn.execute(
            '''SELECT f.id AS fulfillment_id FROM oms_shipments s
               JOIN oms_fulfillments f ON f.id=s.fulfillment_id WHERE s.id=?''',
            (shipment_id,),
        ).fetchone()
        if not row:
            return jsonify({"error": "包裹不存在"}), 404
        _check_fulfillment_permission(conn, row["fulfillment_id"], "can_view")
        key = f"track-on-demand:{shipment_id}:{utcnow()[:16]}"
        job_id = enqueue_job(
            conn, "POLL_SHIPMENT_TRACKING", "shipment", shipment_id, key,
            {"shipment_id": shipment_id},
        )
        conn.commit()
        return jsonify({"queued": True, "job_id": job_id})
    except DomainError as exc:
        return _domain_error(exc)
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/jobs")
@login_required
def api_jobs():
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        rows = conn.execute(
            '''SELECT * FROM oms_integration_jobs
               ORDER BY CASE status WHEN 'dead_letter' THEN 0 WHEN 'retry' THEN 1 ELSE 2 END,
                        updated_at DESC LIMIT 300'''
        ).fetchall()
        summary = {r["status"]: r["n"] for r in conn.execute(
            "SELECT status, COUNT(*) AS n FROM oms_integration_jobs GROUP BY status"
        ).fetchall()}
        return jsonify({"summary": summary, "items": [dict(r) for r in rows]})
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/jobs/<int:job_id>/retry", methods=["POST"])
@login_required
def api_retry_job(job_id):
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        conn.execute(
            '''UPDATE oms_integration_jobs
               SET status='retry', attempts=0, available_at=CURRENT_TIMESTAMP,
                   locked_at=NULL, locked_by=NULL, lease_expires_at=NULL,
                   last_error=NULL, last_error_code=NULL, updated_at=CURRENT_TIMESTAMP
               WHERE id=? AND status IN ('dead_letter','retry')''',
            (job_id,),
        )
        changed = conn.execute("SELECT changes()").fetchone()[0]
        conn.commit()
        return jsonify({"success": bool(changed)})
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/config/sku-warehouses", methods=["GET", "POST"])
@login_required
def api_sku_warehouses():
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        if request.method == "POST":
            body = request.get_json(silent=True) or {}
            conn.execute(
                '''INSERT INTO oms_sku_warehouses
                   (sku_id, warehouse_id, is_primary, is_enabled,
                    wms_product_name_zh, wms_product_name_en, wms_product_image,
                    product_type, notes)
                   VALUES (?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(sku_id, warehouse_id) DO UPDATE SET
                     is_primary=excluded.is_primary, is_enabled=excluded.is_enabled,
                     wms_product_name_zh=excluded.wms_product_name_zh,
                     wms_product_name_en=excluded.wms_product_name_en,
                     wms_product_image=excluded.wms_product_image,
                     product_type=excluded.product_type, notes=excluded.notes,
                     updated_at=CURRENT_TIMESTAMP''',
                (
                    int(body["sku_id"]), int(body["warehouse_id"]),
                    1 if body.get("is_primary") else 0,
                    0 if body.get("is_enabled") is False else 1,
                    body.get("wms_product_name_zh"), body.get("wms_product_name_en"),
                    body.get("wms_product_image"), body.get("product_type") or "P",
                    body.get("notes"),
                ),
            )
            conn.commit()
        rows = conn.execute(
            '''SELECT sw.*, s.sku_code, s.name AS sku_name, s.barcode,
                      w.name AS warehouse_name, w.country
               FROM oms_sku_warehouses sw JOIN inv_skus s ON s.id=sw.sku_id
               JOIN warehouses w ON w.id=sw.warehouse_id
               ORDER BY s.sku_code, w.country'''
        ).fetchall()
        return jsonify([dict(r) for r in rows])
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/config/options", methods=["GET", "POST"])
@login_required
def api_fulfillment_options():
    """Configuration bootstrap and guarded rollout switches.

    Secrets are environment-only and are deliberately never returned here.
    """
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        if request.method == "POST":
            body = request.get_json(silent=True) or {}
            readiness = _rollout_readiness(conn)
            if body.get("oms_auto_plan_enabled") and not readiness["can_auto_plan"]:
                return jsonify({
                    "error": "自动分仓尚未达到上线条件",
                    "code": "auto_plan_not_ready",
                    "readiness": readiness,
                }), 409
            if body.get("auto_submit") and not readiness["can_auto_submit"]:
                return jsonify({
                    "error": "匈牙利 WMS 自动提交尚未达到上线条件",
                    "code": "wms_auto_submit_not_ready",
                    "readiness": readiness,
                }), 409
            for key in ("oms_fulfillment_enabled", "oms_auto_plan_enabled"):
                if key in body:
                    value = "1" if body.get(key) else "0"
                    conn.execute(
                        "INSERT OR REPLACE INTO settings (key, value) VALUES (?,?)",
                        (key, value),
                    )
            if body.get("warehouse_id") is not None:
                fields = []
                values = []
                if "auto_submit" in body:
                    fields.append("auto_submit=?")
                    values.append(1 if body.get("auto_submit") else 0)
                if body.get("channel_code"):
                    fields.append("channel_code=?")
                    values.append(str(body["channel_code"]).strip())
                if fields:
                    values.append(int(body["warehouse_id"]))
                    conn.execute(
                        f"UPDATE oms_warehouse_integrations SET {','.join(fields)}, updated_at=CURRENT_TIMESTAMP WHERE warehouse_id=?",
                        values,
                    )
            conn.commit()
        setting_rows = conn.execute(
            "SELECT key, value FROM settings WHERE key IN ('oms_fulfillment_enabled','oms_auto_plan_enabled')"
        ).fetchall()
        settings = {r["key"]: str(r["value"]).lower() in {"1", "true", "yes", "on"} for r in setting_rows}
        readiness = _rollout_readiness(conn)
        integrations = [dict(r) for r in conn.execute(
            '''SELECT wi.warehouse_id, wi.provider, wi.external_code, wi.channel_code,
                      wi.is_enabled, wi.auto_submit, wi.inventory_authority,
                      wi.tracking_mode, w.name AS warehouse_name, w.country
               FROM oms_warehouse_integrations wi JOIN warehouses w ON w.id=wi.warehouse_id
               ORDER BY w.country, w.id'''
        ).fetchall()]
        return jsonify({
            "settings": settings,
            "readiness": readiness,
            "integrations": integrations,
            "warehouses": [dict(r) for r in conn.execute(
                "SELECT id, name, country FROM warehouses WHERE is_active=1 ORDER BY country, name"
            ).fetchall()],
            "skus": [dict(r) for r in conn.execute(
                "SELECT id, sku_code, name, barcode FROM inv_skus WHERE is_active=1 ORDER BY sku_code"
            ).fetchall()],
            "users": [dict(r) for r in conn.execute(
                "SELECT id, username, name FROM users ORDER BY COALESCE(name, username), id"
            ).fetchall()],
        })
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/config/shipping-costs", methods=["GET", "POST"])
@login_required
def api_shipping_costs():
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        if request.method == "POST":
            body = request.get_json(silent=True) or {}
            conn.execute(
                '''INSERT INTO oms_shipping_costs
                   (market_code, warehouse_id, service_code, amount, currency,
                    is_active, effective_from, effective_to, notes)
                   VALUES (?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(market_code, warehouse_id, service_code) DO UPDATE SET
                     amount=excluded.amount, currency=excluded.currency,
                     is_active=excluded.is_active,
                     effective_from=excluded.effective_from,
                     effective_to=excluded.effective_to, notes=excluded.notes,
                     updated_at=CURRENT_TIMESTAMP''',
                (
                    (body.get("market_code") or "CZ").upper(), int(body["warehouse_id"]),
                    body.get("service_code") or "default", body.get("amount"),
                    body.get("currency") or "EUR", 0 if body.get("is_active") is False else 1,
                    body.get("effective_from"), body.get("effective_to"), body.get("notes"),
                ),
            )
            conn.commit()
        rows = conn.execute(
            '''SELECT sc.*, w.name AS warehouse_name FROM oms_shipping_costs sc
               JOIN warehouses w ON w.id=sc.warehouse_id
               ORDER BY sc.market_code, sc.amount, sc.warehouse_id'''
        ).fetchall()
        return jsonify([dict(r) for r in rows])
    finally:
        conn.close()


@fulfillment_bp.route("/api/fulfillment/config/permissions", methods=["GET", "POST"])
@login_required
def api_warehouse_permissions():
    if not _can_manage_inventory():
        return jsonify({"error": "需要库存管理权限"}), 403
    conn = get_conn()
    try:
        if request.method == "POST":
            body = request.get_json(silent=True) or {}
            flags = ["can_view", "can_pick", "can_pack", "can_ship", "can_cancel", "can_retry", "can_reconcile"]
            values = [1 if body.get(flag) else 0 for flag in flags]
            conn.execute(
                f'''INSERT INTO oms_warehouse_user_permissions
                    (user_id, warehouse_id, {','.join(flags)})
                    VALUES (?, ?, {','.join('?' for _ in flags)})
                    ON CONFLICT(user_id, warehouse_id) DO UPDATE SET
                      {','.join(f'{flag}=excluded.{flag}' for flag in flags)},
                      updated_at=CURRENT_TIMESTAMP''',
                (int(body["user_id"]), int(body["warehouse_id"]), *values),
            )
            conn.commit()
        rows = conn.execute(
            '''SELECT p.*, u.username, u.name AS user_name, w.name AS warehouse_name
               FROM oms_warehouse_user_permissions p
               JOIN users u ON u.id=p.user_id JOIN warehouses w ON w.id=p.warehouse_id
               ORDER BY u.name, w.name'''
        ).fetchall()
        return jsonify([dict(r) for r in rows])
    finally:
        conn.close()
