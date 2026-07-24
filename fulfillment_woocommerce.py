"""WooCommerce/AST adapter for fulfillment shipments and final completion."""

from __future__ import annotations

import hashlib
import html
import json
import time
from typing import Any

import requests

from fulfillment_common import json_load, utcnow
from fulfillment_service import (
    DomainError,
    SHIPMENT_NOTIFICATION_TEMPLATE_VERSION,
    completion_guard,
    record_event,
)
from oid_utils import woo_post_id


class WooError(RuntimeError):
    def __init__(self, message: str, *, code: str = "woo_error", retryable: bool = False, unknown_outcome: bool = False):
        super().__init__(message)
        self.code = code
        self.retryable = retryable
        self.unknown_outcome = unknown_outcome


API_HEADERS = {
    "User-Agent": "woo-analysis-fulfillment/2.0",
    "Content-Type": "application/json",
    "Accept": "application/json",
}
NOTIFICATION_TEMPLATE_VERSION = SHIPMENT_NOTIFICATION_TEMPLATE_VERSION


def detect_tracking_format(conn, site_url: str) -> str:
    rows = conn.execute(
        '''SELECT meta_data, line_items FROM orders
           WHERE source=? AND status IN ('on-hold','shipped','completed','partial-shipped')
           ORDER BY date_modified DESC LIMIT 20''',
        (site_url,),
    ).fetchall()
    ast = villa = custom = 0
    for row in rows:
        meta = row["meta_data"] or ""
        lines = row["line_items"] or ""
        ast += int("_wc_shipment_tracking_items" in meta)
        villa += int("_vi_wot_order_item_tracking_data" in lines)
        custom += int('"key":"tracking_number"' in lines or '"key": "tracking_number"' in lines)
    if ast >= max(villa, custom) and ast:
        return "ast"
    if villa >= custom and villa:
        return "villatheme"
    if custom:
        return "custom_lineitem"
    # The three European stores are expected to use AST; the metadata is also
    # harmless when the plugin is temporarily unavailable.
    return "ast"


def _ast_provider(slug: str | None) -> str:
    value = (slug or "custom").lower().strip()
    return {
        "inpost": "inpost-paczkomaty",
        "dpd": "dpd-pl",
        "auspost": "australia-post",
        "wms-auto": "custom",
    }.get(value, value)


def _shipment_products(conn, shipment_id: str) -> list[dict]:
    rows = conn.execute(
        '''SELECT oi.raw_json, si.quantity
           FROM oms_shipment_items si
           JOIN oms_fulfillment_items fi ON fi.id=si.fulfillment_item_id
           JOIN oms_order_items oi ON oi.id=fi.order_item_id
           WHERE si.shipment_id=?''',
        (shipment_id,),
    ).fetchall()
    products = []
    for row in rows:
        raw = json_load(row["raw_json"], {}) or {}
        if raw.get("id") is None:
            continue
        products.append({
            "product": str(raw.get("product_id") or ""),
            "item_id": str(raw.get("id")),
            "qty": str(row["quantity"]),
        })
    return products


def _all_order_shipments(conn, order_id: str, revision: int) -> list[dict]:
    rows = conn.execute(
        '''SELECT s.*, f.warehouse_id
           FROM oms_shipments s
           JOIN oms_fulfillments f ON f.id=s.fulfillment_id
           WHERE f.order_id=? AND f.revision=? AND f.status!='superseded'
             AND s.status NOT IN ('cancelled','label_pending','label_ready')
           ORDER BY s.created_at, s.id''',
        (order_id, revision),
    ).fetchall()
    result = []
    for row in rows:
        item = dict(row)
        item["products"] = _shipment_products(conn, row["id"])
        result.append(item)
    return result


def _ast_items(shipments: list[dict]) -> list[dict]:
    result = []
    for shipment in shipments:
        tracking = (shipment.get("tracking_number") or "").strip()
        if not tracking:
            continue
        shipped_at = shipment.get("shipped_at") or utcnow()
        try:
            date_shipped = str(int(time.mktime(time.strptime(shipped_at[:19].replace("T", " "), "%Y-%m-%d %H:%M:%S"))))
        except Exception:
            date_shipped = str(int(time.time()))
        result.append({
            "tracking_number": tracking,
            "shipping_note": "",
            "tracking_provider": _ast_provider(shipment.get("carrier_slug")),
            "custom_tracking_link": "",
            "tracking_product_code": "",
            "date_shipped": date_shipped,
            "products_list": shipment.get("products") or [],
            "status_shipped": "1",
            "tracking_id": hashlib.md5(f"{tracking}{date_shipped}".encode()).hexdigest(),
        })
    return result


def _villatheme_lines(conn, shipments: list[dict], order_lines: list[dict]) -> list[dict]:
    records_by_line: dict[str, list[dict]] = {}
    carriers = {r["slug"]: dict(r) for r in conn.execute("SELECT * FROM shipping_carriers").fetchall()}
    for shipment in shipments:
        carrier = carriers.get(shipment.get("carrier_slug")) or {}
        template = (carrier.get("tracking_url") or "").replace("{tracking}", "{tracking_number}")
        record = {
            "tracking_number": shipment.get("tracking_number"),
            "carrier_slug": shipment.get("carrier_slug") or "custom",
            "carrier_name": shipment.get("carrier_name") or carrier.get("name") or "Custom",
            "carrier_url": template,
            "carrier_type": "custom-carrier",
            "time": int(time.time()),
        }
        for product in shipment.get("products") or []:
            records_by_line.setdefault(str(product["item_id"]), []).append(record)
    payload = []
    for line in order_lines:
        line_id = str(line.get("id") or "")
        records = records_by_line.get(line_id)
        if records:
            payload.append({
                "id": line.get("id"),
                "meta_data": [{
                    "key": "_vi_wot_order_item_tracking_data",
                    "value": json.dumps(records, ensure_ascii=False),
                }],
            })
    return payload


def _custom_lines(shipments: list[dict], order_lines: list[dict]) -> list[dict]:
    by_line = {}
    for shipment in shipments:
        for product in shipment.get("products") or []:
            by_line[str(product["item_id"])] = shipment
    payload = []
    for line in order_lines:
        shipment = by_line.get(str(line.get("id") or ""))
        if shipment:
            payload.append({
                "id": line["id"],
                "meta_data": [
                    {"key": "tracking_number", "value": shipment.get("tracking_number")},
                    {"key": "carrier_slug", "value": shipment.get("carrier_slug") or "custom"},
                ],
            })
    return payload


def _is_fully_shipped(conn, order_id: str, revision: int) -> bool:
    state = conn.execute(
        "SELECT has_shortage, manual_review FROM oms_order_fulfillment_state WHERE order_id=?",
        (order_id,),
    ).fetchone()
    if not state or state["has_shortage"] or state["manual_review"]:
        return False
    open_qty = conn.execute(
        '''SELECT COALESCE(SUM(MAX(fi.allocated_qty-fi.fulfilled_qty-fi.cancelled_qty,0)),0) AS qty
           FROM oms_fulfillment_items fi
           JOIN oms_fulfillments f ON f.id=fi.fulfillment_id
           WHERE f.order_id=? AND f.revision=? AND f.status NOT IN ('cancelled','superseded')''',
        (order_id, revision),
    ).fetchone()["qty"]
    unshipped = conn.execute(
        '''SELECT COUNT(*) AS n FROM oms_fulfillments f
           WHERE f.order_id=? AND f.revision=?
             AND f.status NOT IN ('shipped','delivered','cancelled','superseded')''',
        (order_id, revision),
    ).fetchone()["n"]
    return int(open_qty or 0) == 0 and int(unshipped or 0) == 0


def _remote_has_tracking(payload: Any, tracking_number: str) -> bool:
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except ValueError:
            return payload.strip() == tracking_number
    if isinstance(payload, dict):
        return any(_remote_has_tracking(value, tracking_number) for value in payload.values())
    if isinstance(payload, list):
        return any(_remote_has_tracking(value, tracking_number) for value in payload)
    return str(payload or "").strip() == tracking_number


def _request(method: str, url: str, site, *, payload=None, timeout=60):
    try:
        response = requests.request(
            method,
            url,
            json=payload,
            auth=(site["consumer_key"], site["consumer_secret"]),
            headers=API_HEADERS,
            timeout=timeout,
        )
    except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
        raise WooError("WooCommerce 请求超时或连接失败", code="network", retryable=True, unknown_outcome=method != "GET") from exc
    except requests.exceptions.RequestException as exc:
        raise WooError(f"WooCommerce 请求异常: {exc}", code="request", retryable=True) from exc
    if response.status_code == 429 or response.status_code >= 500:
        raise WooError(f"WooCommerce HTTP {response.status_code}", code=f"http_{response.status_code}", retryable=True)
    if response.status_code not in (200, 201):
        raise WooError(
            f"WooCommerce HTTP {response.status_code}: {(response.text or '')[:300]}",
            code=f"http_{response.status_code}",
            retryable=False,
        )
    try:
        return response.json()
    except ValueError:
        return {}


def _shipment_financial_context(conn, fulfillment_id: str) -> dict:
    row = conn.execute(
        '''SELECT ff.*, w.name AS warehouse_name, w.country AS warehouse_country,
                  (SELECT COUNT(*)
                   FROM oms_fulfillments active
                   WHERE active.order_id=f.order_id
                     AND active.revision=f.revision
                     AND active.status NOT IN ('cancelled','superseded')) AS fulfillment_count
           FROM oms_fulfillments f
           LEFT JOIN warehouses w ON w.id=f.warehouse_id
           LEFT JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
           WHERE f.id=?''',
        (fulfillment_id,),
    ).fetchone()
    return dict(row) if row else {}


def _customer_note_body(conn, order, shipment: dict, tracking_url: str = "") -> str:
    tracking = html.escape(str(shipment["tracking_number"]))
    carrier = html.escape(
        str(shipment.get("carrier_name") or shipment.get("carrier_slug") or "物流商")
    )
    finance = _shipment_financial_context(conn, shipment["fulfillment_id"])
    warehouse = html.escape(
        str(finance.get("warehouse_name") or finance.get("warehouse_country") or "仓库")
    )
    if tracking_url:
        safe_url = html.escape(str(tracking_url), quote=True)
        tracking_part = f"<a href='{safe_url}'>{tracking}</a>"
    else:
        tracking_part = tracking
    body = (
        f"您的订单已有一个包裹从 {warehouse} 通过 {carrier} 发出。"
        f"运单号：{tracking_part}。"
    )
    if str(order["payment_method"] or "").lower() == "cod":
        if not finance or finance.get("cod_collection_role") not in {"collector", "instruction_only"}:
            raise WooError("包裹缺少 COD 金额快照，已阻止发送含糊的客户通知", code="cod_notice_missing")
        amount = html.escape(str(finance.get("cod_amount") or "0.00"))
        currency = html.escape(str(finance.get("cod_currency") or order["currency"] or ""))
        body += f"本包裹货到付款金额：{amount} {currency}。"
        if int(finance.get("fulfillment_count") or 0) > 1:
            shipping = str(finance.get("customer_shipping_amount") or "0.00")
            if float(shipping) > 0:
                body += (
                    f"其中包含本订单全部客户运费 {html.escape(shipping)} {currency}；"
                    "其他仓库包裹不会重复收取订单运费。"
                )
            else:
                body += "本包裹只收取所含商品金额，不重复收取订单运费。"
            body += "同一订单的其他仓库包裹将分别收取对应商品金额。"
    else:
        body += "本包裹无需货到付款。"
    if int(finance.get("fulfillment_count") or 0) > 1:
        body += "订单内其他商品将由另一仓库分批发出。"
    return body


def _post_customer_note(conn, site, order, shipment, tracking_url: str = "") -> None:
    body = _customer_note_body(conn, order, shipment, tracking_url)
    url = f"{site['url']}/wp-json/wc/v3/orders/{woo_post_id(order['id'])}/notes"
    _request("POST", url, site, payload={"note": body, "customer_note": True}, timeout=45)


def _notify_shipment(conn, site, order, shipment: dict, fmt: str, final: bool) -> str:
    notification = conn.execute(
        '''SELECT * FROM oms_shipment_notifications
           WHERE shipment_id=? AND channel='email' AND template_version=? ''',
        (shipment["id"], NOTIFICATION_TEMPLATE_VERSION),
    ).fetchone()
    if notification and notification["status"] == "sent":
        return "already_sent"
    try:
        finance = _shipment_financial_context(conn, shipment["fulfillment_id"])
        split_cod = (
            str(order["payment_method"] or "").lower() == "cod"
            and int(finance.get("fulfillment_count") or 0) > 1
        )
        if split_cod:
            # A single generic AST template cannot explain two different COD
            # amounts safely.  Use one idempotent customer-note email per
            # parcel; tracking metadata is still written to AST.
            _post_customer_note(conn, site, order, shipment)
            result = "sent_split_cod_customer_note"
        else:
            trigger_url = f"{site['url']}/wp-json/woo-tracking/v1/orders/{woo_post_id(order['id'])}/trigger-shipment-email"
            response = requests.post(
                trigger_url,
                json={"tracking_number": shipment["tracking_number"], "carrier_slug": shipment["carrier_slug"]},
                params={"consumer_key": site["consumer_key"], "consumer_secret": site["consumer_secret"]},
                headers=API_HEADERS,
                timeout=30,
            )
            data = response.json() if response.status_code == 200 and response.text else {}
            plugin = data.get("plugin")
            sent = data.get("email_sent")
            # AST only sends on the final status->shipped transition.  A partial
            # shipment has no status transition, so guarantee the per-parcel
            # notice with an idempotent Woo customer note.
            if fmt == "ast" and not final:
                _post_customer_note(conn, site, order, shipment)
                result = "sent_partial_customer_note"
            elif sent is True or (fmt == "ast" and final):
                result = f"sent_{plugin or fmt}"
            elif response.status_code == 404 or sent is False:
                _post_customer_note(conn, site, order, shipment)
                result = "sent_customer_note_fallback"
            else:
                # AST final email is queued asynchronously and returns null.
                result = f"scheduled_{plugin or fmt}"
    except Exception as exc:
        conn.execute(
            '''UPDATE oms_shipment_notifications
               SET status='failed', attempts=attempts+1, last_error=?, updated_at=CURRENT_TIMESTAMP
               WHERE shipment_id=? AND channel='email' AND template_version=? ''',
            (str(exc)[:500], shipment["id"], NOTIFICATION_TEMPLATE_VERSION),
        )
        raise WooError(f"发货邮件触发失败: {exc}", code="email_failed", retryable=True) from exc
    conn.execute(
        '''UPDATE oms_shipment_notifications
           SET status='sent', attempts=attempts+1, provider_message_id=?,
               sent_at=?, last_error=NULL, updated_at=CURRENT_TIMESTAMP
           WHERE shipment_id=? AND channel='email' AND template_version=? ''',
        (result, utcnow(), shipment["id"], NOTIFICATION_TEMPLATE_VERSION),
    )
    return result


def sync_shipment(conn, shipment_id: str) -> dict:
    shipment_row = conn.execute(
        '''SELECT s.*, f.order_id, f.revision
           FROM oms_shipments s JOIN oms_fulfillments f ON f.id=s.fulfillment_id
           WHERE s.id=?''',
        (shipment_id,),
    ).fetchone()
    if not shipment_row:
        raise DomainError("包裹不存在", "shipment_not_found")
    shipment = dict(shipment_row)
    order = conn.execute("SELECT * FROM orders WHERE id=?", (shipment["order_id"],)).fetchone()
    site = conn.execute("SELECT * FROM sites WHERE url=?", (order["source"],)).fetchone()
    if not site or not site["consumer_key"] or not site["consumer_secret"]:
        raise WooError("WooCommerce 站点写入凭据缺失", code="site_credentials_missing")

    shipments = _all_order_shipments(conn, order["id"], shipment["revision"])
    fmt = detect_tracking_format(conn, order["source"])
    final = _is_fully_shipped(conn, order["id"], shipment["revision"])
    order_lines = json_load(order["line_items"], []) or []
    latest = shipments[-1]
    payload = {
        "meta_data": [
            {"key": "_tracking_number", "value": latest["tracking_number"]},
            {"key": "_tracking_provider", "value": latest["carrier_slug"] or "custom"},
            {"key": "_date_shipped", "value": str(int(time.time()))},
        ]
    }
    if final:
        payload["status"] = "shipped" if fmt == "ast" else "on-hold"
    if fmt == "ast":
        payload["meta_data"].append({"key": "_wc_shipment_tracking_items", "value": _ast_items(shipments)})
    elif fmt == "villatheme":
        payload["line_items"] = _villatheme_lines(conn, shipments, order_lines)
    else:
        payload["line_items"] = _custom_lines(shipments, order_lines)

    url = f"{site['url']}/wp-json/wc/v3/orders/{woo_post_id(order['id'])}"
    try:
        remote = _request("PUT", url, site, payload=payload)
    except WooError as exc:
        if not exc.unknown_outcome:
            raise
        remote = _request("GET", url, site)
        if not _remote_has_tracking(remote, shipment["tracking_number"]):
            raise
    if not _remote_has_tracking(remote, shipment["tracking_number"]):
        verify = _request("GET", url, site)
        if not _remote_has_tracking(verify, shipment["tracking_number"]):
            raise WooError("WooCommerce 回查未找到新运单号", code="tracking_not_persisted", retryable=True)

    exists = conn.execute(
        "SELECT id FROM shipping_logs WHERE order_id=? AND tracking_number=?",
        (order["id"], shipment["tracking_number"]),
    ).fetchone()
    if not exists:
        conn.execute(
            '''INSERT INTO shipping_logs
               (order_id, woo_order_id, source, tracking_number, carrier_slug, shipped_at)
               VALUES (?,?,?,?,?,?)''',
            (
                order["id"],
                order["number"],
                order["source"],
                shipment["tracking_number"],
                shipment["carrier_slug"] or "custom",
                shipment["shipped_at"] or utcnow(),
            ),
        )
    conn.execute(
        "UPDATE oms_shipments SET woo_sync_status='synced', updated_at=CURRENT_TIMESTAMP WHERE id=?",
        (shipment_id,),
    )
    if final:
        conn.execute("UPDATE orders SET status=? WHERE id=?", (payload["status"], order["id"]))
    notification = _notify_shipment(conn, site, order, shipment, fmt, final)
    record_event(
        conn,
        "shipment",
        shipment_id,
        "synced_to_woocommerce",
        payload={"format": fmt, "final": final, "notification": notification},
    )
    conn.commit()
    return {"shipment_id": shipment_id, "format": fmt, "final": final, "notification": notification}


def complete_order(conn, order_id: str) -> dict:
    allowed, reason = completion_guard(conn, order_id)
    if not allowed:
        raise DomainError(reason or "订单尚未满足完成条件", "completion_blocked")
    order = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
    if not order:
        raise DomainError("订单不存在", "order_not_found")
    site = conn.execute("SELECT * FROM sites WHERE url=?", (order["source"],)).fetchone()
    if not site or not site["consumer_key"] or not site["consumer_secret"]:
        raise WooError("WooCommerce 站点写入凭据缺失", code="site_credentials_missing")
    url = f"{site['url']}/wp-json/wc/v3/orders/{woo_post_id(order['id'])}"
    try:
        remote = _request("PUT", url, site, payload={"status": "completed"})
    except WooError as exc:
        if not exc.unknown_outcome:
            raise
        remote = _request("GET", url, site)
    if remote.get("status") != "completed":
        remote = _request("GET", url, site)
        if remote.get("status") != "completed":
            raise WooError("WooCommerce 回查状态不是 completed", code="completion_not_persisted", retryable=True)
    conn.execute("UPDATE orders SET status='completed' WHERE id=?", (order_id,))
    conn.execute(
        "UPDATE shipping_logs SET status='completed', completed_at=COALESCE(completed_at,datetime('now')) WHERE order_id=?",
        (order_id,),
    )
    conn.execute(
        '''UPDATE oms_order_fulfillment_state
           SET completion_sync_status='synced', updated_at=CURRENT_TIMESTAMP WHERE order_id=?''',
        (order_id,),
    )
    record_event(conn, "order", order_id, "completed_in_woocommerce", to_status="completed")
    conn.commit()
    return {"order_id": order_id, "status": "completed"}
