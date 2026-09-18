"""Cancel an unshipped order and its warehouse work from one entry point.

Stop local fulfillment before the WooCommerce mutation, without holding a
database transaction across HTTP. Durable operation records prevent duplicate
writes after a timeout. An unconfirmed website cancellation stays visible for
review; stopped warehouse work is never silently restarted.
"""
from __future__ import annotations

import requests

from external_operations import (
    SHIPMENT_LOCK_NAMESPACE, begin_operation, transition_operation,
)
from fulfillment_service import (
    DomainError, record_event, transition_fulfillment, _table_exists,
)
from oid_utils import woo_post_id


LOCAL_CANCELLABLE = {
    "planned", "ready_to_pick", "ready_to_submit", "stock_shortage",
    "manual_hold", "manual_review", "rejected",
}
REMOTE_CANCELLABLE = {"pending", "processing", "offline", "failed", "checkout-draft", "cancelled"}
OPERATION_TYPE = "cancel_fulfillment_order"
PENDING_REASON = "仓库履约已停止，网站订单取消结果待确认；请在订单详情再次取消以核对结果"
HEADERS = {"User-Agent": "WooCommerce API Client-Python/3.0.0", "Accept": "application/json"}


def cancellation_guard(conn, order_id, *, allowed_warehouse_ids=None, lock=False):
    """Validate every warehouse before changing any of them."""
    pg = hasattr(conn, "_raw")
    if lock and pg:
        # Serialize with legacy shipment submissions and with replanning.
        conn.execute("SELECT pg_advisory_xact_lock(?,hashtext(?))",
                     (SHIPMENT_LOCK_NAMESPACE, str(order_id)))
    order = conn.execute(
        "SELECT * FROM orders WHERE id=?" + (" FOR UPDATE" if lock and pg else ""),
        (order_id,),
    ).fetchone()
    if not order:
        raise DomainError("订单不存在", "order_not_found")
    if order["status"] not in REMOTE_CANCELLABLE:
        raise DomainError("订单已有发货、完成或退款状态，请先核对包裹并办理拦截/售后", "order_not_cancellable")
    rows = conn.execute(
        "SELECT * FROM oms_fulfillments WHERE order_id=? ORDER BY id"
        + (" FOR UPDATE" if lock and pg else ""), (order_id,),
    ).fetchall()
    for row in rows:
        if row["shipped_at"] or row["delivered_at"] or row["status"] in {
            "shipped", "delivered", "returning", "returned", "exception",
        }:
            raise DomainError("已有仓库发货，请先联系仓库/承运商拦截，不能直接取消整单", "order_already_shipped")
        if row["status"] in {"cancelled", "superseded"}:
            continue
        if allowed_warehouse_ids is not None and row["warehouse_id"] not in allowed_warehouse_ids:
            raise DomainError("没有全部相关仓库的取消权限，请联系有权限的负责人取消整单", "warehouse_forbidden")
        if row["status"] in {"submitting", "submission_unknown", "cancel_pending", "cancel_rejected"}:
            raise DomainError("外部仓提交或取消结果尚未确认，请先在多仓履约核验/拦截外部单据", "external_cancellation_required")
        if row["mode"] == "external_wms" and (
            row["submitted_at"] or row["accepted_at"] or row["external_pick_code"]
        ):
            raise DomainError("订单已提交外部仓，请先在多仓履约确认仓库取消或联系仓库拦截，再取消整单", "external_cancellation_required")
        if row["status"] not in LOCAL_CANCELLABLE and not (
            row["mode"] == "internal" and row["status"] in {"picking", "packed"}
        ):
            raise DomainError(f"履约状态 {row['status']} 需要先核对仓库处理结果", "fulfillment_not_cancellable")
    shipped_items = conn.execute(
        """SELECT 1 FROM oms_fulfillment_items i JOIN oms_fulfillments f
           ON f.id=i.fulfillment_id WHERE f.order_id=? AND i.fulfilled_qty>0 LIMIT 1""", (order_id,),
    ).fetchone()
    shipments = conn.execute(
        """SELECT 1 FROM oms_shipments s JOIN oms_fulfillments f ON f.id=s.fulfillment_id
           WHERE f.order_id=? AND (s.status!='cancelled' OR s.shipped_at IS NOT NULL) LIMIT 1""", (order_id,),
    ).fetchone()
    if shipped_items or shipments:
        raise DomainError("订单已有包裹、运单或出库记录，请先核对并办理拦截", "order_has_shipment")
    if _table_exists(conn, "shipping_logs") and conn.execute(
        "SELECT 1 FROM shipping_logs WHERE order_id=? LIMIT 1", (order_id,),
    ).fetchone():
        raise DomainError("订单已有发货记录，请先核对并办理拦截", "order_has_shipment")
    if _table_exists(conn, "external_operations") and conn.execute(
        """SELECT 1 FROM external_operations WHERE order_id=? AND operation_type='ship_order'
           AND status NOT IN ('failed','cancelled') LIMIT 1""", (order_id,),
    ).fetchone():
        raise DomainError("订单存在已提交或待核对的发货操作，请先核对发货结果", "shipment_pending")
    return rows


def stop_local_fulfillments(conn, order_id, *, actor, allowed_warehouse_ids=None):
    rows = cancellation_guard(conn, order_id, allowed_warehouse_ids=allowed_warehouse_ids, lock=True)
    cancelled = []
    for row in rows:
        if row["status"] == "superseded":
            continue
        if row["status"] != "cancelled":
            if row["status"] in {"picking", "packed"}:
                transition_fulfillment(conn, row["id"], "manual_hold", actor=actor, reason="整单取消停止拣货/打包")
            transition_fulfillment(conn, row["id"], "cancelled", actor=actor, reason="订单详情整单取消")
            cancelled.append(row["id"])
        # Also repair counts on fulfillments cancelled by the old endpoint.
        conn.execute(
            """UPDATE oms_fulfillment_items SET cancelled_qty=allocated_qty-fulfilled_qty,
               updated_at=CURRENT_TIMESTAMP WHERE fulfillment_id=?""", (row["id"],),
        )
    conn.execute(
        """UPDATE oms_order_fulfillment_state SET aggregate_status='cancelled',
           has_shortage=0,manual_review=1,manual_reason=?,updated_at=CURRENT_TIMESTAMP
           WHERE order_id=?""", (PENDING_REASON, order_id),
    )
    return cancelled


def _remote_order(response, order_id):
    if response.status_code != 200:
        raise ValueError(f"网站订单读取失败（HTTP {response.status_code}）")
    payload = response.json()
    if not isinstance(payload, dict) or str(payload.get("id")) != str(woo_post_id(order_id)):
        raise ValueError("网站返回的订单身份不一致")
    return payload


def _check_remote_unshipped(payload):
    if payload.get("status") not in REMOTE_CANCELLABLE:
        raise DomainError("网站订单已有发货、完成或退款状态，请先核对并办理拦截/售后", "remote_order_not_cancellable")
    if payload.get("tracking_number") or payload.get("shipment_tracking"):
        raise DomainError("网站订单已有运单，请先核对并办理拦截", "remote_order_has_shipment")
    for item in payload.get("meta_data") or []:
        if item.get("key") in {"_wc_shipment_tracking_items", "_tracking_number", "tracking_number"}:
            if item.get("value") not in (None, "", [], {}, "[]", "{}"):
                raise DomainError("网站订单已有运单，请先核对并办理拦截", "remote_order_has_shipment")


def _finish(conn, order_id, operation_id, actor):
    # Keep the same order -> operation lock order used by preparation/retries.
    conn.execute("SELECT id FROM orders WHERE id=? FOR UPDATE", (order_id,)).fetchone()
    row = conn.execute("SELECT status FROM external_operations WHERE operation_id=? FOR UPDATE", (operation_id,)).fetchone()
    if row["status"] in {"pending", "reconciliation_required"}:
        transition_operation(conn, operation_id, "external_success", evidence={"status": "cancelled"})
    if row["status"] != "local_committed":
        transition_operation(conn, operation_id, "local_committed")
    conn.execute("UPDATE orders SET status='cancelled' WHERE id=?", (order_id,))
    conn.execute(
        """UPDATE oms_order_items SET cancelled_qty=ordered_qty,shortage_qty=0,
           updated_at=CURRENT_TIMESTAMP WHERE order_id=?""", (order_id,),
    )
    conn.execute(
        """UPDATE oms_order_fulfillment_state SET aggregate_status='cancelled',has_shortage=0,
           manual_review=0,manual_reason=NULL,completion_sync_status='not_ready',
           updated_at=CURRENT_TIMESTAMP WHERE order_id=?""", (order_id,),
    )
    if row["status"] != "local_committed":
        record_event(conn, "order", order_id, "order_cancelled", to_status="cancelled", actor=actor,
                     reason="订单详情整单取消，网站状态已确认", correlation_id=operation_id)
    conn.commit()


def cancel_fulfillment_order(conn, order_id, *, actor, allowed_warehouse_ids=None):
    """Return a JSON payload and HTTP status; caller closes the connection."""
    stopped = False
    try:
        cancellation_guard(conn, order_id, allowed_warehouse_ids=allowed_warehouse_ids)
        order = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        site_row = conn.execute("SELECT * FROM sites WHERE url=?", (order["source"],)).fetchone()
        if not site_row:
            return {"success": False, "error": "站点配置不存在"}, 404
        site = dict(site_row)
        if site.get("api_write_status") == "error":
            return {"success": False, "error": "该站点没有 API 写入权限，无法取消订单"}, 403
        old_status = order["status"]
        url = f"{site['url'].rstrip('/')}/wp-json/wc/v3/orders/{woo_post_id(order_id)}"
        auth = (site["consumer_key"], site["consumer_secret"])
        conn.rollback()
        try:
            remote = _remote_order(requests.get(url, auth=auth, headers=HEADERS, timeout=(5, 10)), order_id)
        except (requests.RequestException, ValueError):
            return {"success": False, "error": "暂时无法核对网站订单状态，请稍后再试；尚未提交取消"}, 422
        _check_remote_unshipped(remote)

        stop_local_fulfillments(conn, order_id, actor=actor, allowed_warehouse_ids=allowed_warehouse_ids)
        pending = conn.execute(
            """SELECT operation_id,status FROM external_operations WHERE order_id=? AND operation_type=?
               AND status IN ('pending','external_success','reconciliation_required')
               ORDER BY created_at DESC LIMIT 1 FOR UPDATE""", (order_id, OPERATION_TYPE),
        ).fetchone()
        if not pending and remote.get("status") == "cancelled":
            pending = conn.execute(
                """SELECT operation_id,status FROM external_operations WHERE order_id=? AND operation_type=?
                   AND status='local_committed' ORDER BY created_at DESC LIMIT 1 FOR UPDATE""",
                (order_id, OPERATION_TYPE),
            ).fetchone()
        operation = dict(pending) if pending else begin_operation(
            conn, operation_type=OPERATION_TYPE, order_id=order_id, site_id=site["id"],
            request_payload={"status": "cancelled", "source_modified": remote.get("date_modified_gmt") or remote.get("date_modified")},
            created_by=actor.get("id"),
        )
        operation_id = str(operation["operation_id"])
        conn.commit()
        stopped = True
        confirmed = remote.get("status") == "cancelled" or operation["status"] == "local_committed"
        uncertain = True
        if not confirmed and operation.get("should_execute"):
            try:
                response = requests.put(url, json={"status": "cancelled"}, auth=auth,
                                        headers=HEADERS, timeout=(5, 15))
                if response.status_code == 200:
                    confirmed = _remote_order(response, order_id).get("status") == "cancelled"
                elif 400 <= response.status_code < 500:
                    # A valid WooCommerce JSON rejection is safe for a later explicit retry.
                    error_payload = response.json()
                    uncertain = not (isinstance(error_payload, dict) and bool(error_payload.get("code")))
            except (requests.RequestException, ValueError):
                pass
            if not confirmed:
                try:
                    verified = _remote_order(requests.get(url, auth=auth, headers=HEADERS, timeout=(5, 10)), order_id)
                    confirmed = verified.get("status") == "cancelled"
                except (requests.RequestException, ValueError):
                    pass
            if not confirmed:
                current = conn.execute("SELECT status FROM external_operations WHERE operation_id=? FOR UPDATE", (operation_id,)).fetchone()
                if current["status"] == "pending":
                    transition_operation(conn, operation_id, "reconciliation_required" if uncertain else "failed",
                                         error="网站取消结果待核验" if uncertain else "网站明确拒绝取消")
                conn.commit()
        if not confirmed:
            return {"success": False, "uncertain": uncertain, "fulfillments_stopped": True,
                    "retry_safe": not uncertain,
                    "error": "仓库履约已停止，但网站订单取消尚未确认。" + (
                        "请稍后再次点击取消核对结果，系统不会重复发送未确认的取消请求；仍未确认时请检查网站后台。"
                        if uncertain else "网站明确拒绝了请求，请核对站点权限后重试。")}, 409 if uncertain else 422
        _finish(conn, order_id, operation_id, actor)
        return {"success": True, "old_status": old_status, "new_status": "cancelled",
                "message": "订单已取消，相关仓库履约已停止，预留库存已释放，网站状态已确认"}, 200
    except DomainError as exc:
        conn.rollback()
        return {"success": False, "error": str(exc), "code": exc.code}, 403 if exc.code == "warehouse_forbidden" else 409
    except Exception:
        conn.rollback()
        if not stopped:
            raise
        # Preserve the durable intent even if final local persistence failed.
        return {"success": False, "uncertain": True, "retry_safe": False,
                "error": "仓库履约已停止，网站取消结果需要核对。请再次点击取消核验，系统不会重复发送未确认的请求。"}, 409
