"""Verified WooCommerce product writes for the product-management UI.

This module intentionally has no Flask or database dependency.  It implements
the remote write boundary: defensive JSON parsing, two-phase hard-sold-out
updates, authoritative read-back verification, and master-to-child checks.
"""

import html
import time
from decimal import Decimal, InvalidOperation


PRODUCT_EDIT_FIELDS = {
    "manage_stock",
    "stock_quantity",
    "stock_status",
    "regular_price",
    "sale_price",
}

PRODUCT_STATE_FIELDS = (
    "id",
    "parent_id",
    "name",
    "slug",
    "sku",
    "type",
    "manage_stock",
    "stock_quantity",
    "stock_status",
    "regular_price",
    "sale_price",
    "price",
)


class ProductOperationExpired(Exception):
    pass


def new_product_operation_deadline():
    # Includes master writes, child lookup, child repair and all read-backs.
    # Leave headroom below the production Gunicorn worker's 120-second limit.
    return time.monotonic() + 90


def _request_timeout(deadline, read_limit=20):
    remaining = deadline - time.monotonic()
    if remaining < 1:
        raise ProductOperationExpired("本次核验时间已用完，请稍后重试核验已有结果")
    return min(5, remaining / 2), min(read_limit, remaining / 2)


def _payload_matches(item, payload):
    if product_payload_mismatches(item, payload):
        return False
    metadata = {m.get("key"): m.get("value") for m in (item.get("meta_data") or [])}
    return all(str(metadata.get(m["key"])) == str(m.get("value"))
               for m in payload.get("meta_data", []))


def _read_product(req, resource_url, auth, headers, deadline):
    try:
        resp = req.get(resource_url, auth=auth, timeout=_request_timeout(deadline),
                       headers={**headers, "Cache-Control": "no-cache"})
    except (req.RequestException, ProductOperationExpired) as exc:
        return None, str(exc)
    item, error = parse_wc_response(resp)
    if error:
        return None, error
    if not isinstance(item, dict) or not item.get("id"):
        return None, "商品核验响应缺少有效商品 ID"
    return item, None


def parse_wc_response(resp):
    """Return ``(json_data, error_message)`` for a WC REST response."""
    body_preview = ((resp.text or "")[:300]).replace("\n", " ")
    looks_like_html = body_preview.lstrip().lower().startswith(
        ("<!doctype", "<html", "<?xml")
    )

    if resp.status_code not in (200, 201):
        try:
            data = resp.json()
            message = data.get("message") or body_preview or "(无响应内容)"
        except Exception:
            message = body_preview or "(无响应内容)"
        return None, f"WC API 错误 HTTP {resp.status_code}: {message}"

    try:
        return resp.json(), None
    except Exception:
        if looks_like_html:
            message = (
                f"WC API 返回了 HTML 而不是 JSON（HTTP {resp.status_code}）。"
                f"常见原因：CloudFlare 拦截、缓存插件命中、PHP 致命错误页。"
                f"响应开头: {body_preview}"
            )
        else:
            message = (
                f"WC API 响应解析失败（HTTP {resp.status_code}）："
                f"{body_preview}"
            )
        return None, message


def product_state_snapshot(item):
    """Return the small, non-sensitive part of a WC product response we audit."""
    item = item or {}
    return {field: item.get(field) for field in PRODUCT_STATE_FIELDS}


def product_payload_mismatches(item, payload):
    """Compare a freshly-read WC resource with requested mutable fields."""
    item = item or {}
    mismatches = []
    for field, expected in payload.items():
        if field not in PRODUCT_EDIT_FIELDS:
            continue
        actual = item.get(field)
        matched = False
        if field == "manage_stock":
            matched = bool(actual) is bool(expected)
        elif field == "stock_quantity":
            try:
                matched = int(actual) == int(expected)
            except (TypeError, ValueError):
                matched = False
        elif field in ("regular_price", "sale_price"):
            if expected in ("", None):
                matched = actual in ("", None)
            else:
                try:
                    matched = Decimal(str(actual)) == Decimal(str(expected))
                except (InvalidOperation, TypeError, ValueError):
                    matched = False
        else:
            matched = str(actual or "") == str(expected or "")
        if not matched:
            mismatches.append(
                f"{field}: 期望 {expected!r}，实际 {actual!r}"
            )
    return mismatches


def wc_product_update_verified(req, resource_url, auth, payload, *, deadline=None):
    from stock_sync_guard import legacy_write
    from stock_sync_common import SyncError
    try:
        with legacy_write(resource_url, payload):
            return _wc_product_update_verified(
                req, resource_url, auth, payload,
                deadline=deadline if deadline is not None else new_product_operation_deadline(),
            )
    except SyncError as exc:
        return None, f'{exc.code}: {exc}', {"phases": [], "final_state": None}


def _wc_product_update_verified(req, resource_url, auth, payload, *, deadline):
    """Write a product/variation and GET it back before reporting success."""
    headers = {
        "User-Agent": "WooCommerce API Client-Python/3.0.0",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    hard_sold_out = (
        payload.get("manage_stock") is False
        and payload.get("stock_status") == "outofstock"
    )
    phases = []
    if hard_sold_out:
        first = dict(payload)
        first.pop("stock_status", None)
        if first:
            phases.append(first)
        phases.append({"stock_status": "outofstock"})
    else:
        phases.append(dict(payload))

    trace = {"phases": [], "final_state": None}
    current, error = _read_product(req, resource_url, auth, headers, deadline)
    if error:
        return None, "写入前核验失败，未提交修改：" + error, trace
    trace["preflight_state"] = product_state_snapshot(current)
    if _payload_matches(current, payload):
        trace.update(already_satisfied=True, final_state=product_state_snapshot(current))
        return current, None, trace

    fresh_read = True
    for phase in phases:
        if _payload_matches(current, phase):
            continue
        try:
            timeout = _request_timeout(deadline)
        except ProductOperationExpired as exc:
            trace["final_state"] = product_state_snapshot(current) if fresh_read else None
            trace["unconfirmed_write"] = not fresh_read
            return current, str(exc), trace
        phase_trace = {"payload": phase, "http_status": None}
        trace["phases"].append(phase_trace)
        try:
            resp = req.put(
                resource_url,
                auth=auth,
                json=phase,
                timeout=timeout,
                headers=headers,
            )
            changed, error = parse_wc_response(resp)
            if changed is not None and not isinstance(changed, dict):
                changed, error = None, "WC API 写入响应不是商品对象"
            phase_trace.update(http_status=resp.status_code, response=product_state_snapshot(changed))
        except req.RequestException as exc:
            error = f"连接失败: {exc}"
        if error:
            # The remote save may have committed before its hooks timed out.
            # Read back; never replay this phase on an ambiguous response.
            phase_trace["error"] = error
            current, read_error = _read_product(req, resource_url, auth, headers, deadline)
            trace["final_state"] = product_state_snapshot(current) if current else None
            if read_error or not _payload_matches(current, phase):
                trace["unconfirmed_write"] = True
                return current, "写入结果未确认：" + error + (
                    "；回读失败：" + read_error if read_error else "；回读未达到目标状态"
                ), trace
            phase_trace["verified_after_error"] = True
            fresh_read = True
        else:
            current = changed or {}
            fresh_read = False

    final, error = (current, None) if fresh_read else _read_product(req, resource_url, auth, headers, deadline)
    if error:
        trace["unconfirmed_write"] = True
        return None, f"写入后校验失败: {error}", trace
    final = final or {}
    trace["final_state"] = product_state_snapshot(final)
    mismatches = product_payload_mismatches(final, payload)
    if mismatches:
        return final, "写入未达到目标状态：" + "；".join(mismatches), trace
    return final, None, trace


def find_child_product(req, site, master_item, *, deadline=None):
    """Resolve a master product or variation on its actual child site."""
    child_url = (site["url"] or "").rstrip("/")
    auth = (site["consumer_key"], site["consumer_secret"])
    headers = {
        "User-Agent": "WooCommerce API Client-Python/3.0.0",
        "Accept": "application/json",
        "Cache-Control": "no-cache",
    }
    sku = (master_item.get("sku") or "").strip()
    name = html.unescape(master_item.get("name") or "").strip()
    params = {"per_page": 100, "status": "any"}
    if sku:
        params["sku"] = sku
    elif name:
        params["search"] = name
    else:
        return None, "商品没有 SKU 或名称，无法定位子站对应商品"

    # A master write has already completed. A transient read failure must not
    # replay it; retry only this read once with a smaller timeout budget. A
    # repeated failure remains pending and never triggers a blind child write.
    deadline = deadline if deadline is not None else new_product_operation_deadline()
    for attempt, read_limit in enumerate((20, 15)):
        try:
            resp = req.get(
                f"{child_url}/wp-json/wc/v3/products",
                auth=auth,
                params=params,
                timeout=_request_timeout(deadline, read_limit),
                headers=headers,
            )
            if attempt == 0 and resp.status_code in (502, 503, 504, 520, 521, 522, 523, 524, 525, 526):
                continue
            break
        except ProductOperationExpired as exc:
            return None, str(exc)
        except req.RequestException as exc:
            if attempt == 1:
                return None, f"查询子站失败（已重试只读核验）: {exc}"
    candidates, error = parse_wc_response(resp)
    if error:
        return None, f"查询子站失败: {error}"
    candidates = candidates or []

    if sku:
        for candidate in candidates:
            if (
                (candidate.get("sku") or "").strip().casefold()
                == sku.casefold()
            ):
                return candidate, None
    master_slug = (master_item.get("slug") or "").strip().casefold()
    if master_slug:
        for candidate in candidates:
            if (
                (candidate.get("slug") or "").strip().casefold()
                == master_slug
            ):
                return candidate, None
    normalized_name = " ".join(name.casefold().split())
    for candidate in candidates:
        candidate_name = " ".join(
            html.unescape(candidate.get("name") or "").casefold().split()
        )
        if candidate_name == normalized_name:
            return candidate, None
    return None, f'子站未找到对应商品（SKU={sku or "空"}）'


def wcms_stock_meta_update(child_item, master_item, payload):
    """Return companion stock metadata for the site's custom WCMS bridge.

    Some child stores intentionally restore WooCommerce's core stock fields
    from ``wcms_stock_*`` metadata after every product save.  Detect that
    contract from the child response and update the bridge metadata together
    with the requested core fields.  Stores without this metadata are left
    untouched.
    """
    if not any(
        key in payload
        for key in ("manage_stock", "stock_quantity", "stock_status")
    ):
        return []
    child_meta = {
        str(row.get("key")): row.get("value")
        for row in (child_item.get("meta_data") or [])
        if isinstance(row, dict)
    }
    if "wcms_stock_manage" not in child_meta:
        return []

    quantity = master_item.get("stock_quantity")
    try:
        quantity = int(quantity) if quantity is not None else 0
    except (TypeError, ValueError):
        quantity = 0
    status = master_item.get("stock_status") or "instock"
    return [
        {
            "key": "wcms_stock_manage",
            "value": "yes" if bool(master_item.get("manage_stock")) else "no",
        },
        {"key": "wcms_stock_qty", "value": max(0, quantity)},
        {"key": "wcms_stock_status", "value": status},
    ]


def verify_product_child_sync(req, site, master_item, payload, *, deadline=None):
    """Ensure that a master-routed write reaches the selected child site.

    WooCommerce Multistore does not reliably propagate stock and price changes
    made through the master's REST API.  Read the child first; if it has not
    converged, write the same whitelisted payload through the child's own REST
    credentials and require an authoritative read-back before reporting success.
    """
    if not site["product_master_id"]:
        return {
            "status": "not_applicable",
            "detail": "直连站点，无需子站同步校验",
            "state": None,
        }
    if not (
        site["url"] and site["consumer_key"] and site["consumer_secret"]
    ):
        return {
            "status": "error",
            "detail": "子站 WC API 凭据不完整，无法验证同步结果",
            "state": None,
        }
    deadline = deadline if deadline is not None else new_product_operation_deadline()
    child_item, error = find_child_product(req, site, master_item, deadline=deadline)
    if error:
        return {"status": "pending", "detail": error, "state": None}
    child_state = product_state_snapshot(child_item)
    mismatches = product_payload_mismatches(child_item, payload)
    if mismatches:
        child_url = (site["url"] or "").rstrip("/")
        child_id = child_item.get("id")
        parent_id = child_item.get("parent_id")
        if not child_id:
            return {
                "status": "error",
                "detail": "已定位子站商品，但响应缺少商品 ID，无法补写",
                "state": child_state,
            }
        if parent_id:
            resource_url = (
                f"{child_url}/wp-json/wc/v3/products/{parent_id}"
                f"/variations/{child_id}"
            )
        elif child_item.get("type") == "variation":
            return {
                "status": "error",
                "detail": "已定位子站变体，但响应缺少父商品 ID，无法补写",
                "state": child_state,
            }
        else:
            resource_url = f"{child_url}/wp-json/wc/v3/products/{child_id}"

        direct_payload = dict(payload)
        bridge_meta = wcms_stock_meta_update(child_item, master_item, payload)
        if bridge_meta:
            direct_payload["meta_data"] = bridge_meta

        final, write_error, write_trace = wc_product_update_verified(
            req,
            resource_url,
            (site["consumer_key"], site["consumer_secret"]),
            direct_payload,
            deadline=deadline,
        )
        if write_error:
            return {
                "status": "error",
                "detail": "子站自动同步未完成，直接补写也失败：" + write_error,
                "state": product_state_snapshot(final or child_item),
                "direct_update": True,
                "trace": write_trace,
            }
        return {
            "status": "verified",
            "detail": "主站写入成功；子站未自动同步，已直接补写并回读一致",
            "state": product_state_snapshot(final),
            "direct_update": True,
            "trace": write_trace,
        }
    return {
        "status": "verified",
        "detail": "商品主站与子站状态一致",
        "state": child_state,
    }


def product_operation_action(payload):
    if (
        payload.get("manage_stock") is False
        and payload.get("stock_status") == "outofstock"
    ):
        return "soldout"
    if payload.get("manage_stock") is True and "stock_quantity" in payload:
        return "stock_restore"
    if any(key in payload for key in ("regular_price", "sale_price")):
        return "price_update"
    return "product_update"
