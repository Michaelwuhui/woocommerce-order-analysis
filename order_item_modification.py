"""订单详情商品行更换口味/更换产品服务模块。

职责边界：
1. 权限严格限制为站点负责人本人（sites.manager == TRIM(users.name)）或内置超级管理员（username == 'admin'），禁止跨站修改他人订单。
2. 仅允许修改未发货、未进入仓库拣货/打包、未推送外部 WMS 的活跃订单（pending / processing / offline）。
3. 支持同款换口味（固定保持原成交金额不变）与跨产品更换（可选保持原成交金额不变或同步新产品站点售价），支持整行替换与部分数量拆分。
4. 先通过 WooCommerce REST API 原子修改源站订单并回读确认，再在同一本地事务中更新本地订单镜像与 OMS 多仓履约分仓预留。
"""

from __future__ import annotations

import html
import json
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any
from urllib.parse import urljoin, urlsplit

import requests

from external_operations import (
    SHIPMENT_LOCK_NAMESPACE,
    begin_operation,
    transition_operation,
)
from fulfillment_service import (
    DomainError,
    PLANNABLE_FULFILLMENT_STATUSES,
    _setting_enabled,
    _table_exists,
    plan_order,
    record_event,
)
from oid_utils import woo_post_id

MODIFIABLE_ORDER_STATUSES = {"pending", "processing", "offline"}
OPERATION_TYPE = "modify_order_item"
TWOPLACES = Decimal("0.01")

API_HEADERS = {
    "User-Agent": "WooCommerce API Client-Python/3.0.0",
    "Content-Type": "application/json",
    "Accept": "application/json",
}

STATUS_CN_LABELS = {
    "pending": "待付款",
    "processing": "处理中",
    "on-hold": "已发货",
    "shipped": "已发货",
    "completed": "已完成",
    "cancelled": "已取消",
    "refunded": "已退款",
    "failed": "失败",
    "checkout-draft": "草稿",
}

FLAVOR_ATTR_KEYWORDS = (
    "smak",
    "smaki",
    "flavor",
    "flavour",
    "taste",
    "variant",
    "íz",
    "iz",
    "příchuť",
    "prichut",
    "口味",
)


def _dec(value: Any, default: str = "0.00") -> Decimal:
    """将金额安全转换为保留两位小数的 Decimal。"""
    if value in (None, ""):
        return Decimal(default).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    try:
        return Decimal(str(value).strip()).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError, TypeError):
        return Decimal(default).quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def _clean_text(value: Any) -> str:
    """清理字符串并解码 HTML 实体。"""
    if value is None:
        return ""
    return html.unescape(str(value)).strip()


def can_user_modify_site_orders(conn, source_url: str | None, user, *, lock: bool = False) -> bool:
    """校验当前登录用户是否为该站点的负责人本人（或内置超级管理员 admin）。"""
    if not source_url or not user or not getattr(user, "is_authenticated", True):
        return False
    if getattr(user, "is_viewer", lambda: False)() or getattr(user, "role", None) == "viewer":
        return False

    pg = hasattr(conn, "_raw")
    if getattr(user, "username", None) == "admin":
        suffix = " FOR SHARE" if lock and pg else ""
        row = conn.execute(
            "SELECT id FROM sites WHERE url = ? AND COALESCE(api_status, '') != 'archived'" + suffix,
            (source_url,),
        ).fetchone()
        return row is not None

    try:
        user_id = int(getattr(user, "id", None))
    except (TypeError, ValueError):
        return False

    has_exclusions = _table_exists(conn, "user_site_exclusions")
    exclusion_sql = (
        " AND NOT EXISTS (SELECT 1 FROM user_site_exclusions e WHERE e.user_id = u.id AND e.site_id = s.id)"
        if has_exclusions
        else ""
    )
    suffix = " FOR SHARE OF s, u" if lock and pg else ""
    row = conn.execute(
        f"""SELECT s.id
            FROM sites s
            JOIN users u ON u.id = ?
            WHERE s.url = ?
              AND u.role IN ('admin', 'user')
              AND COALESCE(TRIM(u.name), '') != ''
              AND TRIM(COALESCE(s.manager, '')) = TRIM(u.name)
              AND COALESCE(s.api_status, '') != 'archived'
              {exclusion_sql}"""
        + suffix,
        (user_id, source_url),
    ).fetchone()
    return row is not None


def evaluate_order_modification_eligibility(
    conn,
    order: Any,
    user,
    *,
    lock: bool = False,
) -> dict[str, Any]:
    """评估订单当前是否允许修改商品行，返回权限归属、可改标记及不可改原因。"""
    if not order:
        return {
            "is_owner": False,
            "can_modify": False,
            "block_reason": "订单不存在",
            "code": "order_not_found",
        }

    order_id = str(order["id"])
    source_url = order["source"]
    is_owner = can_user_modify_site_orders(conn, source_url, user, lock=lock)
    if not is_owner:
        return {
            "is_owner": False,
            "can_modify": False,
            "block_reason": "只能修改本人负责站点的订单",
            "code": "site_owner_forbidden",
        }

    site_row = conn.execute("SELECT * FROM sites WHERE url = ?", (source_url,)).fetchone()
    if not site_row:
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": "站点配置不存在",
            "code": "site_not_found",
        }
    site = dict(site_row)
    if str(site.get("api_status") or "") == "archived":
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": "站点连接已归档移除，无法同步修改",
            "code": "site_archived",
        }
    if str(site.get("api_write_status") or "") == "error":
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": "该站点 API 密钥无写入权限，无法修改源站订单",
            "code": "site_write_forbidden",
        }
    if not site.get("consumer_key") or not site.get("consumer_secret"):
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": "站点缺少 API 凭据配置",
            "code": "site_credentials_missing",
        }

    order_dict = dict(order)
    if bool(order_dict.get("is_undelivered")) or bool(order_dict.get("is_problem_return")):
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": "订单已有未送达或问题退货记录，无法修改商品",
            "code": "order_returned",
        }

    status = str(order_dict.get("status") or "").strip()
    if status not in MODIFIABLE_ORDER_STATUSES:
        label = STATUS_CN_LABELS.get(status, status or "未知")
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": f"当前订单状态为「{label}」，仅待付款/处理中的未发货订单允许修改商品",
            "code": "order_status_locked",
        }

    if _table_exists(conn, "shipping_logs") and conn.execute(
        "SELECT 1 FROM shipping_logs WHERE order_id = ? LIMIT 1",
        (order_id,),
    ).fetchone():
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": "订单已有发货运单记录，无法修改商品",
            "code": "order_has_shipment",
        }

    if _table_exists(conn, "external_operations") and conn.execute(
        """SELECT 1 FROM external_operations
           WHERE order_id = ?
             AND operation_type IN ('ship_order', 'cancel_fulfillment_order')
             AND status NOT IN ('failed', 'cancelled')
           LIMIT 1""",
        (order_id,),
    ).fetchone():
        return {
            "is_owner": True,
            "can_modify": False,
            "block_reason": "订单存在待确认的发货或取消操作，请先核对处理结果",
            "code": "operation_pending",
        }

    if _table_exists(conn, "oms_fulfillments"):
        pg = hasattr(conn, "_raw")
        ff_rows = conn.execute(
            "SELECT * FROM oms_fulfillments WHERE order_id = ? ORDER BY id"
            + (" FOR UPDATE" if lock and pg else ""),
            (order_id,),
        ).fetchall()
        for row in ff_rows:
            ff = dict(row)
            ff_status = str(ff.get("status") or "")
            if ff_status in {"cancelled", "superseded"}:
                continue
            if ff.get("shipped_at") or ff.get("delivered_at") or ff_status in {
                "shipped",
                "delivered",
                "returning",
                "returned",
                "exception",
            }:
                return {
                    "is_owner": True,
                    "can_modify": False,
                    "block_reason": "仓库已有发货记录，无法修改商品",
                    "code": "fulfillment_shipped",
                }
            if ff.get("mode") == "external_wms" and (
                ff.get("submitted_at") or ff.get("accepted_at") or ff.get("external_pick_code")
            ):
                return {
                    "is_owner": True,
                    "can_modify": False,
                    "block_reason": "订单已提交外部仓（WMS），请先在多仓履约拦截后再处理",
                    "code": "external_wms_submitted",
                }
            if ff_status not in PLANNABLE_FULFILLMENT_STATUSES:
                return {
                    "is_owner": True,
                    "can_modify": False,
                    "block_reason": f"仓库履约已进入「{ff_status}」阶段，请先暂停仓库作业后再修改",
                    "code": "fulfillment_in_progress",
                }

        if _table_exists(conn, "oms_fulfillment_items") and conn.execute(
            """SELECT 1 FROM oms_fulfillment_items i
               JOIN oms_fulfillments f ON f.id = i.fulfillment_id
               WHERE f.order_id = ? AND i.fulfilled_qty > 0
               LIMIT 1""",
            (order_id,),
        ).fetchone():
            return {
                "is_owner": True,
                "can_modify": False,
                "block_reason": "订单已有出库数量记录，无法修改商品",
                "code": "fulfillment_item_shipped",
            }

        if _table_exists(conn, "oms_shipments") and conn.execute(
            """SELECT 1 FROM oms_shipments s
               JOIN oms_fulfillments f ON f.id = s.fulfillment_id
               WHERE f.order_id = ? AND (s.status != 'cancelled' OR s.shipped_at IS NOT NULL)
               LIMIT 1""",
            (order_id,),
        ).fetchone():
            return {
                "is_owner": True,
                "can_modify": False,
                "block_reason": "订单已生成包裹记录，无法修改商品",
                "code": "shipment_exists",
            }

    return {
        "is_owner": True,
        "can_modify": True,
        "block_reason": None,
        "code": None,
    }


def _wc_request(
    method: str,
    site: dict[str, Any],
    path: str,
    *,
    params: dict[str, Any] | None = None,
    json_body: dict[str, Any] | None = None,
    timeout: tuple[int, int] = (6, 20),
    session: Any = None,
) -> requests.Response:
    """向 WooCommerce 站点发起安全请求，仅允许同域 HTTPS 裸域/www 规范重定向保留鉴权。"""
    http = session or requests
    base = str(site["url"]).rstrip("/")
    url = f"{base}/wp-json/wc/v3/{path.lstrip('/')}"
    auth = (site["consumer_key"], site["consumer_secret"])
    origin = urlsplit(url)
    host_base = (origin.hostname or "").removeprefix("www.")
    allowed_hosts = {origin.hostname, host_base, f"www.{host_base}"}

    for _ in range(3):
        resp = http.request(
            method=method.upper(),
            url=url,
            params=params,
            json=json_body,
            auth=auth,
            headers=API_HEADERS,
            timeout=timeout,
            allow_redirects=False,
        )
        if resp.status_code not in {301, 302, 303, 307, 308}:
            return resp
        location = resp.headers.get("Location") or ""
        target = urljoin(url, location)
        parsed = urlsplit(target)
        if (
            parsed.scheme != "https"
            or parsed.hostname not in allowed_hosts
            or parsed.port != origin.port
            or parsed.username
            or parsed.password
            or target == url
        ):
            raise DomainError("站点发生非预期跨域重定向，已拒绝发送凭据", "unsafe_redirect")
        url = target
        params = None  # 重定向目标 URL 中已包含查询串
    raise DomainError("站点重定向次数过多，无法访问接口", "redirect_loop")


def _check_remote_order_unshipped(remote_order: dict[str, Any]) -> None:
    """校验源站实时订单仍处于未发货可修改状态。"""
    remote_status = str(remote_order.get("status") or "").strip()
    if remote_status not in MODIFIABLE_ORDER_STATUSES:
        label = STATUS_CN_LABELS.get(remote_status, remote_status or "未知")
        raise DomainError(
            f"源站订单当前状态已变为「{label}」，无法修改商品",
            "remote_status_locked",
        )
    if remote_order.get("tracking_number") or remote_order.get("shipment_tracking"):
        raise DomainError("源站订单已存在运单号，无法修改商品", "remote_has_tracking")
    for meta in remote_order.get("meta_data") or []:
        if not isinstance(meta, dict):
            continue
        if meta.get("key") in {"_wc_shipment_tracking_items", "_tracking_number", "tracking_number"}:
            if meta.get("value") not in (None, "", [], {}, "[]", "{}"):
                raise DomainError("源站订单已写入物流运单信息，无法修改商品", "remote_has_tracking")
    for item in remote_order.get("line_items") or []:
        if not isinstance(item, dict):
            continue
        for meta in item.get("meta_data") or []:
            if isinstance(meta, dict) and meta.get("key") in {
                "_vi_wot_order_item_tracking_data",
                "tracking_number",
            }:
                if meta.get("value") not in (None, "", [], {}, "[]", "{}"):
                    raise DomainError("源站订单商品行已写入运单信息，无法修改商品", "remote_has_tracking")


def _attribute_meta_key(attr: dict[str, Any]) -> str:
    """推导在 WooCommerce 订单行 meta_data 中使用的属性键名。"""
    slug = str(attr.get("slug") or "").strip()
    attr_id = int(attr.get("id") or 0)
    if slug:
        if attr_id > 0 and not slug.startswith("pa_"):
            return f"pa_{slug}"
        return slug
    name = _clean_text(attr.get("name"))
    return name or "flavor"


def _is_flavor_attr(attr: dict[str, Any]) -> bool:
    """判断属性是否为口味/规格类属性。"""
    slug = str(attr.get("slug") or "").lower()
    name = _clean_text(attr.get("name")).lower()
    return any(k in slug or k in name for k in FLAVOR_ATTR_KEYWORDS)


def _build_product_options(parent: dict[str, Any], variations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """根据父商品及其变体列表构建前端可选规格列表（支持具体变体与任意口味共用变体）。"""
    parent_id = int(parent["id"])
    parent_name = _clean_text(parent.get("name")) or f"商品 #{parent_id}"
    prod_type = str(parent.get("type") or "simple").strip()

    if prod_type != "variable" or not variations:
        price_str = str(parent.get("price") or parent.get("regular_price") or "").strip()
        stock_status = str(parent.get("stock_status") or "instock")
        return [
            {
                "option_key": f"{parent_id}:0:",
                "product_id": parent_id,
                "variation_id": 0,
                "flavor_label": "默认规格（单品）",
                "full_name": parent_name,
                "sku": _clean_text(parent.get("sku")),
                "price": str(_dec(price_str)) if price_str != "" else None,
                "stock_status": stock_status,
                "stock_quantity": parent.get("stock_quantity"),
                "in_stock": stock_status != "outofstock",
                "explicit_meta": [],
            }
        ]

    parent_attrs = [a for a in (parent.get("attributes") or []) if isinstance(a, dict)]
    parent_options_by_key: dict[str, tuple[dict[str, Any], list[str]]] = {}
    for pattr in parent_attrs:
        opts = [_clean_text(o) for o in (pattr.get("options") or []) if _clean_text(o)]
        if opts:
            by_id = f"id:{int(pattr.get('id') or 0)}"
            by_name = f"name:{_clean_text(pattr.get('name')).casefold()}"
            parent_options_by_key[by_id] = (pattr, opts)
            parent_options_by_key[by_name] = (pattr, opts)

    options: list[dict[str, Any]] = []
    for var in variations:
        if not isinstance(var, dict):
            continue
        var_status = str(var.get("status") or "publish").strip()
        if var_status not in {"publish", ""}:
            continue
        var_id = int(var.get("id") or 0)
        if var_id <= 0:
            continue

        price_raw = var.get("price")
        if price_raw in (None, ""):
            price_raw = var.get("regular_price") or parent.get("price") or ""
        price_formatted = str(_dec(price_raw)) if str(price_raw).strip() != "" else None
        stock_status = str(var.get("stock_status") or "instock")
        stock_qty = var.get("stock_quantity")
        sku = _clean_text(var.get("sku")) or _clean_text(parent.get("sku"))

        var_attrs = [a for a in (var.get("attributes") or []) if isinstance(a, dict)]
        fixed_parts: list[str] = []
        wildcard_attr: tuple[dict[str, Any], list[str]] | None = None

        var_attr_ids = {int(a.get("id") or 0) for a in var_attrs if int(a.get("id") or 0) > 0}
        var_attr_names = {_clean_text(a.get("name")).casefold() for a in var_attrs if _clean_text(a.get("name"))}

        for vattr in var_attrs:
            opt_val = _clean_text(vattr.get("option"))
            if opt_val:
                fixed_parts.append(opt_val)
            elif wildcard_attr is None:
                lookup_id = f"id:{int(vattr.get('id') or 0)}"
                lookup_name = f"name:{_clean_text(vattr.get('name')).casefold()}"
                wildcard_attr = parent_options_by_key.get(lookup_id) or parent_options_by_key.get(lookup_name)

        # WooCommerce 中“任意口味”变体有时直接在变体 attributes 数组中省略该属性
        if not fixed_parts and wildcard_attr is None:
            for pattr in parent_attrs:
                if not pattr.get("variation"):
                    continue
                pid = int(pattr.get("id") or 0)
                pname = _clean_text(pattr.get("name")).casefold()
                if (pid and pid in var_attr_ids) or (pname and pname in var_attr_names):
                    continue
                opts = [_clean_text(o) for o in (pattr.get("options") or []) if _clean_text(o)]
                if opts:
                    wildcard_attr = (pattr, opts)
                    break

        if wildcard_attr is not None:
            pattr, candidate_opts = wildcard_attr
            meta_key = _attribute_meta_key(pattr)
            for cand in candidate_opts:
                parts = fixed_parts + [cand]
                label = " / ".join(parts)
                options.append(
                    {
                        "option_key": f"{parent_id}:{var_id}:{cand}",
                        "product_id": parent_id,
                        "variation_id": var_id,
                        "flavor_label": label,
                        "full_name": f"{parent_name} - {label}",
                        "sku": sku,
                        "price": price_formatted,
                        "stock_status": stock_status,
                        "stock_quantity": stock_qty,
                        "in_stock": stock_status != "outofstock",
                        "explicit_meta": [{"key": meta_key, "value": cand}],
                    }
                )
        else:
            var_name = _clean_text(var.get("name"))
            if fixed_parts:
                label = " / ".join(fixed_parts)
            elif var_name and var_name.casefold() != parent_name.casefold():
                label = var_name
            else:
                label = f"变体 #{var_id}"
            full_name = (
                parent_name
                if label.casefold() in parent_name.casefold()
                else f"{parent_name} - {label}"
            )
            options.append(
                {
                    "option_key": f"{parent_id}:{var_id}:",
                    "product_id": parent_id,
                    "variation_id": var_id,
                    "flavor_label": label,
                    "full_name": full_name,
                    "sku": sku,
                    "price": price_formatted,
                    "stock_status": stock_status,
                    "stock_quantity": stock_qty,
                    "in_stock": stock_status != "outofstock",
                    "explicit_meta": [],
                }
            )

    # 有货排在前面，同状态下保持站点原有顺序
    options.sort(key=lambda item: (0 if item["in_stock"] else 1))
    return options


def _load_order_and_authorized_site(conn, order_id: str, user) -> tuple[dict[str, Any], dict[str, Any]]:
    """校验订单商品修改资格并返回 (order_dict, site_dict)。"""
    order_row = conn.execute("SELECT * FROM orders WHERE id = ?", (order_id,)).fetchone()
    if not order_row:
        raise DomainError("订单不存在", "order_not_found")
    check = evaluate_order_modification_eligibility(conn, order_row, user, lock=False)
    if not check["can_modify"]:
        raise DomainError(check["block_reason"] or "当前订单不可修改", check["code"] or "forbidden")
    site_row = conn.execute("SELECT * FROM sites WHERE url = ?", (order_row["source"],)).fetchone()
    if not site_row:
        raise DomainError("站点配置不存在", "site_not_found")
    return dict(order_row), dict(site_row)


def search_site_products_for_order(
    conn,
    order_id: str,
    *,
    user,
    query: str = "",
    page: int = 1,
    session: Any = None,
) -> tuple[dict[str, Any], int]:
    """按关键词读取订单所属站点的在售商品列表（用于跨款更换商品）。"""
    try:
        order, site = _load_order_and_authorized_site(conn, order_id, user)
    except DomainError as exc:
        status = 404 if exc.code in {"order_not_found", "site_not_found"} else (
            403 if exc.code in {"site_owner_forbidden", "site_write_forbidden"} else 409
        )
        return {"success": False, "error": str(exc), "code": exc.code}, status

    conn.rollback()
    q = (query or "").strip()
    page = max(1, int(page or 1))
    params: dict[str, Any] = {
        "status": "publish",
        "per_page": 40,
        "page": page,
        "_fields": "id,name,sku,type,status,price,regular_price,stock_status,stock_quantity,variations",
    }
    if q:
        params["search"] = q

    try:
        resp = _wc_request("GET", site, "products", params=params, timeout=(6, 20), session=session)
        if resp.status_code != 200:
            return {"success": False, "error": f"读取站点商品列表失败（HTTP {resp.status_code}）"}, 502
        items = resp.json()
        if not isinstance(items, list):
            return {"success": False, "error": "站点返回的商品列表格式异常"}, 502

        # 若搜索关键词未命中标题但可能是精确 SKU，补充一次按 SKU 查询
        if q and page == 1 and not items:
            sku_params = dict(params)
            sku_params.pop("search", None)
            sku_params["sku"] = q
            sku_resp = _wc_request("GET", site, "products", params=sku_params, timeout=(6, 15), session=session)
            if sku_resp.status_code == 200 and isinstance(sku_resp.json(), list):
                items = sku_resp.json()
    except DomainError as exc:
        return {"success": False, "error": str(exc), "code": exc.code}, 502
    except (requests.RequestException, ValueError):
        return {"success": False, "error": "连接站点读取商品超时或失败，请稍后重试"}, 502

    products = []
    for raw in items:
        if not isinstance(raw, dict) or not raw.get("id"):
            continue
        if str(raw.get("status") or "publish") not in {"publish", ""}:
            continue
        price_raw = raw.get("price") if raw.get("price") not in (None, "") else raw.get("regular_price")
        variations_list = raw.get("variations") if isinstance(raw.get("variations"), list) else []
        products.append(
            {
                "product_id": int(raw["id"]),
                "name": _clean_text(raw.get("name")) or f"商品 #{raw['id']}",
                "sku": _clean_text(raw.get("sku")),
                "type": str(raw.get("type") or "simple"),
                "price": str(_dec(price_raw)) if price_raw not in (None, "") else None,
                "stock_status": str(raw.get("stock_status") or "instock"),
                "in_stock": str(raw.get("stock_status") or "instock") != "outofstock",
                "variation_count": len(variations_list),
            }
        )

    total_pages_raw = str(resp.headers.get("X-WP-TotalPages") or "1").strip()
    total_pages = int(total_pages_raw) if total_pages_raw.isdigit() else 1
    return {
        "success": True,
        "currency": order.get("currency") or "",
        "page": page,
        "has_more": page < max(1, total_pages),
        "products": products,
    }, 200


def get_product_options_for_order(
    conn,
    order_id: str,
    *,
    user,
    product_id: int,
    session: Any = None,
) -> tuple[dict[str, Any], int]:
    """拉取站点指定产品详情及其全部可选口味/变体列表。"""
    try:
        product_id = int(product_id)
        if product_id <= 0:
            return {"success": False, "error": "无效的商品编号"}, 400
    except (TypeError, ValueError):
        return {"success": False, "error": "无效的商品编号"}, 400

    try:
        order, site = _load_order_and_authorized_site(conn, order_id, user)
    except DomainError as exc:
        status = 404 if exc.code in {"order_not_found", "site_not_found"} else (
            403 if exc.code in {"site_owner_forbidden", "site_write_forbidden"} else 409
        )
        return {"success": False, "error": str(exc), "code": exc.code}, status

    conn.rollback()
    fields = "id,name,sku,type,status,attributes,variations,price,regular_price,stock_status,stock_quantity"
    try:
        prod_resp = _wc_request(
            "GET",
            site,
            f"products/{product_id}",
            params={"_fields": fields},
            timeout=(6, 20),
            session=session,
        )
        if prod_resp.status_code == 404:
            return {"success": False, "error": "该商品在源站已不存在或已下架删除，请切换到「换其他产品」搜索选择"}, 404
        if prod_resp.status_code != 200:
            return {"success": False, "error": f"读取源站商品详情失败（HTTP {prod_resp.status_code}）"}, 502
        parent = prod_resp.json()
        if not isinstance(parent, dict) or int(parent.get("id") or 0) != product_id:
            return {"success": False, "error": "源站返回的商品信息格式异常"}, 502

        variations: list[dict[str, Any]] = []
        if str(parent.get("type") or "") == "variable":
            var_page = 1
            while var_page <= 10:
                var_resp = _wc_request(
                    "GET",
                    site,
                    f"products/{product_id}/variations",
                    params={
                        "per_page": 100,
                        "page": var_page,
                        "_fields": "id,name,sku,status,attributes,price,regular_price,stock_status,stock_quantity",
                    },
                    timeout=(6, 25),
                    session=session,
                )
                if var_resp.status_code != 200:
                    return {
                        "success": False,
                        "error": f"读取商品口味变体失败（HTTP {var_resp.status_code}）",
                    }, 502
                batch = var_resp.json()
                if not isinstance(batch, list):
                    return {"success": False, "error": "源站返回的变体列表格式异常"}, 502
                variations.extend(batch)
                total_pages_hdr = str(var_resp.headers.get("X-WP-TotalPages") or "").strip()
                if total_pages_hdr.isdigit():
                    if var_page >= int(total_pages_hdr):
                        break
                elif len(batch) < 100:
                    break
                var_page += 1
    except DomainError as exc:
        return {"success": False, "error": str(exc), "code": exc.code}, 502
    except (requests.RequestException, ValueError):
        return {"success": False, "error": "读取源站商品及口味超时或失败，请稍后重试"}, 502

    options = _build_product_options(parent, variations)
    return {
        "success": True,
        "currency": order.get("currency") or "",
        "product": {
            "product_id": product_id,
            "name": _clean_text(parent.get("name")) or f"商品 #{product_id}",
            "type": str(parent.get("type") or "simple"),
            "sku": _clean_text(parent.get("sku")),
        },
        "options": options,
    }, 200


def compute_line_item_replacement_amounts(
    *,
    orig_qty: int,
    replace_qty: int,
    orig_subtotal: Any,
    orig_total: Any,
    price_mode: str,
    target_unit_price: Any = None,
) -> dict[str, Any]:
    """计算整行替换或部分数量拆分时的精确金额（保留两位小数，keep 模式严格零尾差）。"""
    if orig_qty <= 0:
        raise DomainError("原订单行数量异常，无法修改", "invalid_orig_qty")
    if replace_qty <= 0 or replace_qty > orig_qty:
        raise DomainError(f"更换数量必须在 1 到 {orig_qty} 之间", "invalid_replace_qty")

    s_orig = _dec(orig_subtotal)
    t_orig = _dec(orig_total)
    remain_qty = orig_qty - replace_qty

    if price_mode not in {"keep", "sync"}:
        raise DomainError("无效的金额处理模式", "invalid_price_mode")

    if price_mode == "keep":
        if remain_qty == 0:
            rem_sub = Decimal("0.00")
            rem_tot = Decimal("0.00")
            new_sub = s_orig
            new_tot = t_orig
        else:
            rem_sub = (s_orig * Decimal(remain_qty) / Decimal(orig_qty)).quantize(
                TWOPLACES, rounding=ROUND_HALF_UP
            )
            rem_tot = (t_orig * Decimal(remain_qty) / Decimal(orig_qty)).quantize(
                TWOPLACES, rounding=ROUND_HALF_UP
            )
            new_sub = (s_orig - rem_sub).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
            new_tot = (t_orig - rem_tot).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    else:
        if target_unit_price in (None, ""):
            raise DomainError(
                "所选新商品在站点未设置有效售价，请选择「保持原成交金额不变」或先在站点设置售价",
                "target_price_missing",
            )
        unit_price = _dec(target_unit_price, default="-1")
        if unit_price < Decimal("0.00"):
            raise DomainError("所选新商品售价无效", "target_price_invalid")
        if remain_qty == 0:
            rem_sub = Decimal("0.00")
            rem_tot = Decimal("0.00")
        else:
            rem_sub = (s_orig * Decimal(remain_qty) / Decimal(orig_qty)).quantize(
                TWOPLACES, rounding=ROUND_HALF_UP
            )
            rem_tot = (t_orig * Decimal(remain_qty) / Decimal(orig_qty)).quantize(
                TWOPLACES, rounding=ROUND_HALF_UP
            )
        new_sub = (unit_price * Decimal(replace_qty)).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
        new_tot = new_sub

    delta_total = (rem_tot + new_tot - t_orig).quantize(TWOPLACES, rounding=ROUND_HALF_UP)
    return {
        "orig_qty": orig_qty,
        "replace_qty": replace_qty,
        "remain_qty": remain_qty,
        "remain_subtotal": f"{rem_sub:.2f}",
        "remain_total": f"{rem_tot:.2f}",
        "new_subtotal": f"{new_sub:.2f}",
        "new_total": f"{new_tot:.2f}",
        "delta_total": f"{delta_total:.2f}",
    }


def _extract_item_flavor_signature(item: dict[str, Any]) -> str:
    """提取订单商品行的口味签名，用于防呆判断是否选了完全相同的原口味。"""
    parts = []
    for meta in item.get("meta_data") or []:
        if not isinstance(meta, dict):
            continue
        key = str(meta.get("key") or "").strip().lower()
        if not key or key.startswith("_"):
            continue
        val = _clean_text(meta.get("display_value") or meta.get("value"))
        if val:
            parts.append(val.casefold())
    return "|".join(sorted(parts))


def _update_local_order_from_remote(conn, order_id: str, remote_order: dict[str, Any], source_url: str) -> None:
    """在当前事务中将源站返回的最新订单核心字段同步至本地 orders 表。"""
    pg = hasattr(conn, "_raw")

    def _money_param(raw_val: Any):
        if raw_val is None or raw_val == "":
            return None
        d = _dec(raw_val)
        return d if pg else float(d)

    existing_cols = {
        str(row["name"] if isinstance(row, dict) or hasattr(row, "keys") else row[1])
        for row in conn.execute("PRAGMA table_info(orders)").fetchall()
    }
    candidate_updates: dict[str, Any] = {
        "status": remote_order.get("status"),
        "currency": remote_order.get("currency"),
        "date_modified": remote_order.get("date_modified"),
        "date_modified_gmt": remote_order.get("date_modified_gmt"),
        "discount_total": _money_param(remote_order.get("discount_total")),
        "discount_tax": _money_param(remote_order.get("discount_tax")),
        "shipping_total": _money_param(remote_order.get("shipping_total")),
        "shipping_tax": _money_param(remote_order.get("shipping_tax")),
        "cart_tax": _money_param(remote_order.get("cart_tax")),
        "total": _money_param(remote_order.get("total")),
        "total_tax": _money_param(remote_order.get("total_tax")),
        "cart_hash": remote_order.get("cart_hash"),
        "line_items": json.dumps(remote_order.get("line_items") or [], ensure_ascii=False),
        "shipping_lines": json.dumps(remote_order.get("shipping_lines") or [], ensure_ascii=False),
        "fee_lines": json.dumps(remote_order.get("fee_lines") or [], ensure_ascii=False),
        "coupon_lines": json.dumps(remote_order.get("coupon_lines") or [], ensure_ascii=False),
        "tax_lines": json.dumps(remote_order.get("tax_lines") or [], ensure_ascii=False),
        "meta_data": json.dumps(remote_order.get("meta_data") or [], ensure_ascii=False),
        "updated_at": datetime.now().isoformat(),
    }
    assignments = []
    params = []
    for col, val in candidate_updates.items():
        if col in existing_cols and val is not None:
            assignments.append(f"{col} = ?")
            params.append(val)
    if assignments:
        params.append(order_id)
        conn.execute(f"UPDATE orders SET {', '.join(assignments)} WHERE id = ?", tuple(params))


def _remote_modification_verified(
    remote_order: dict[str, Any],
    *,
    orig_item_id: int,
    remain_qty: int,
    target_product_id: int,
    target_variation_id: int,
    replace_qty: int,
    prior_modified: str,
) -> bool:
    """当 PUT 请求发生网络读超时时，通过回读源站订单核验修改是否其实已在源站生效。"""
    items = [i for i in (remote_order.get("line_items") or []) if isinstance(i, dict)]
    orig_match = next((i for i in items if int(i.get("id") or 0) == orig_item_id), None)
    if remain_qty == 0:
        if orig_match is not None:
            return False
    else:
        if orig_match is None or int(orig_match.get("quantity") or 0) != remain_qty:
            return False

    has_new_target = any(
        int(i.get("id") or 0) != orig_item_id
        and int(i.get("product_id") or 0) == target_product_id
        and int(i.get("variation_id") or 0) == target_variation_id
        and int(i.get("quantity") or 0) == replace_qty
        for i in items
    )
    current_modified = str(
        remote_order.get("date_modified_gmt") or remote_order.get("date_modified") or ""
    )
    return has_new_target or (bool(prior_modified) and current_modified != prior_modified and remain_qty == 0)


def modify_order_line_item(
    conn,
    order_id: str,
    payload: dict[str, Any],
    *,
    user,
    session: Any = None,
) -> tuple[dict[str, Any], int]:
    """修改订单指定商品行的口味或产品，同步更新 WooCommerce 源站、订单备注与多仓履约库存预留。"""
    data = payload or {}
    try:
        line_item_id = int(data.get("line_item_id"))
        target_product_id = int(data.get("target_product_id"))
        target_variation_id = int(data.get("target_variation_id") or 0)
        replace_qty = int(data.get("replace_qty"))
    except (TypeError, ValueError):
        return {"success": False, "error": "提交的商品行或数量参数格式不正确"}, 400

    if line_item_id <= 0 or target_product_id <= 0 or target_variation_id < 0 or replace_qty <= 0:
        return {"success": False, "error": "商品编号与更换数量必须为正整数"}, 400

    expected_orig_qty = data.get("expected_original_qty")
    if expected_orig_qty is not None:
        try:
            expected_orig_qty = int(expected_orig_qty)
        except (TypeError, ValueError):
            return {"success": False, "error": "原数量校验参数无效"}, 400

    requested_price_mode = str(data.get("price_mode") or "keep").strip().lower()
    if requested_price_mode not in {"keep", "sync"}:
        return {"success": False, "error": "无效的金额处理方式"}, 400

    target_flavor = _clean_text(data.get("target_flavor"))

    try:
        order, site = _load_order_and_authorized_site(conn, order_id, user)
    except DomainError as exc:
        status = 404 if exc.code in {"order_not_found", "site_not_found"} else (
            403 if exc.code in {"site_owner_forbidden", "site_write_forbidden"} else 409
        )
        return {"success": False, "error": str(exc), "code": exc.code}, status

    woo_id = woo_post_id(order_id)
    conn.rollback()

    # 1. 远端预检：实时读取源站订单与目标商品/变体（不在持有数据库锁期间请求外部网络）
    try:
        order_resp = _wc_request("GET", site, f"orders/{woo_id}", timeout=(6, 15), session=session)
        if order_resp.status_code != 200:
            return {
                "success": False,
                "error": f"无法读取源站订单状态（HTTP {order_resp.status_code}），尚未修改任何数据",
            }, 502
        remote_order = order_resp.json()
        if not isinstance(remote_order, dict) or str(remote_order.get("id")) != str(woo_id):
            return {"success": False, "error": "源站返回的订单身份不匹配，已中止操作"}, 502
        _check_remote_order_unshipped(remote_order)
    except DomainError as exc:
        return {"success": False, "error": str(exc), "code": exc.code}, 409
    except (requests.RequestException, ValueError):
        return {"success": False, "error": "连接源站核对订单状态失败，请稍后重试；尚未修改任何数据"}, 502

    remote_items = [i for i in (remote_order.get("line_items") or []) if isinstance(i, dict)]
    orig_item = next((i for i in remote_items if int(i.get("id") or 0) == line_item_id), None)
    if not orig_item:
        return {"success": False, "error": "源站订单中未找到该商品明细行，可能已被修改，请刷新订单详情"}, 409

    orig_qty = int(orig_item.get("quantity") or 0)
    if expected_orig_qty is not None and orig_qty != expected_orig_qty:
        return {
            "success": False,
            "error": f"源站该商品行当前数量为 {orig_qty}（与页面显示的 {expected_orig_qty} 不一致），请刷新后重试",
        }, 409
    if replace_qty > orig_qty:
        return {"success": False, "error": f"更换数量（{replace_qty}）不能超过当前行订购数量（{orig_qty}）"}, 400

    orig_product_id = int(orig_item.get("product_id") or 0)
    orig_variation_id = int(orig_item.get("variation_id") or 0)
    is_same_product = target_product_id == orig_product_id
    # 同款产品不同口味固定保持原成交金额不变
    effective_price_mode = "keep" if is_same_product else requested_price_mode

    # 实时读取目标商品与变体，核实真实存在性、口味属性与站点售价
    try:
        prod_resp = _wc_request(
            "GET",
            site,
            f"products/{target_product_id}",
            params={"_fields": "id,name,sku,type,status,attributes,price,regular_price,stock_status"},
            timeout=(6, 20),
            session=session,
        )
        if prod_resp.status_code != 200:
            return {"success": False, "error": "所选目标商品在源站不存在或无法读取"}, 422
        target_parent = prod_resp.json()
        if not isinstance(target_parent, dict) or int(target_parent.get("id") or 0) != target_product_id:
            return {"success": False, "error": "源站返回的目标商品信息异常"}, 502

        prod_type = str(target_parent.get("type") or "simple")
        target_variations: list[dict[str, Any]] = []
        if prod_type == "variable":
            if target_variation_id <= 0:
                return {"success": False, "error": "该商品包含多个口味变体，请先选择具体口味"}, 400
            var_resp = _wc_request(
                "GET",
                site,
                f"products/{target_product_id}/variations/{target_variation_id}",
                params={"_fields": "id,name,sku,status,attributes,price,regular_price,stock_status"},
                timeout=(6, 20),
                session=session,
            )
            if var_resp.status_code != 200:
                return {"success": False, "error": "所选口味变体在源站不存在或已下架"}, 422
            var_obj = var_resp.json()
            if not isinstance(var_obj, dict) or int(var_obj.get("id") or 0) != target_variation_id:
                return {"success": False, "error": "源站返回的口味变体信息异常"}, 502
            target_variations = [var_obj]
        else:
            target_variation_id = 0
    except DomainError as exc:
        return {"success": False, "error": str(exc), "code": exc.code}, 502
    except (requests.RequestException, ValueError):
        return {"success": False, "error": "核验目标商品与口味信息超时，请稍后重试"}, 502

    candidate_options = _build_product_options(target_parent, target_variations)
    if not candidate_options:
        return {"success": False, "error": "所选商品没有可用的已发布规格"}, 422

    chosen_option = None
    if target_flavor:
        for opt in candidate_options:
            if (
                opt["variation_id"] == target_variation_id
                and opt["flavor_label"].casefold() == target_flavor.casefold()
            ):
                chosen_option = opt
                break
    if chosen_option is None:
        chosen_option = next(
            (opt for opt in candidate_options if opt["variation_id"] == target_variation_id),
            candidate_options[0],
        )

    # 防呆拦截：不允许换成与原商品完全相同的商品+变体+口味
    if is_same_product and target_variation_id == orig_variation_id:
        orig_sig = _extract_item_flavor_signature(orig_item)
        new_sig = (chosen_option["flavor_label"] or "").strip().casefold()
        if not chosen_option["explicit_meta"] or orig_sig == new_sig:
            return {"success": False, "error": "所选目标口味与当前订单行口味完全相同，无需修改"}, 400

    try:
        amounts = compute_line_item_replacement_amounts(
            orig_qty=orig_qty,
            replace_qty=replace_qty,
            orig_subtotal=orig_item.get("subtotal", orig_item.get("total", "0")),
            orig_total=orig_item.get("total", "0"),
            price_mode=effective_price_mode,
            target_unit_price=chosen_option.get("price"),
        )
    except DomainError as exc:
        return {"success": False, "error": str(exc), "code": exc.code}, 400

    remain_qty = amounts["remain_qty"]
    line_items_patch: list[dict[str, Any]] = []
    if remain_qty == 0:
        line_items_patch.append({"id": line_item_id, "quantity": 0})
    else:
        line_items_patch.append(
            {
                "id": line_item_id,
                "quantity": remain_qty,
                "subtotal": amounts["remain_subtotal"],
                "total": amounts["remain_total"],
            }
        )

    new_line_payload: dict[str, Any] = {
        "product_id": target_product_id,
        "quantity": replace_qty,
        "subtotal": amounts["new_subtotal"],
        "total": amounts["new_total"],
    }
    if target_variation_id > 0:
        new_line_payload["variation_id"] = target_variation_id
    if chosen_option.get("explicit_meta"):
        new_line_payload["meta_data"] = chosen_option["explicit_meta"]
    line_items_patch.append(new_line_payload)

    prior_modified = str(
        remote_order.get("date_modified_gmt") or remote_order.get("date_modified") or ""
    )

    # 2. 加锁复检本地状态并写入持久化操作记录（PostgreSQL）
    pg = hasattr(conn, "_raw")
    use_ledger = pg and _table_exists(conn, "external_operations")
    operation_id = None
    should_execute = True

    try:
        if pg:
            conn.execute(
                "SELECT pg_advisory_xact_lock(?, hashtext(?))",
                (SHIPMENT_LOCK_NAMESPACE, str(order_id)),
            )
            locked_order = conn.execute(
                "SELECT * FROM orders WHERE id = ? FOR UPDATE", (order_id,)
            ).fetchone()
        else:
            locked_order = conn.execute("SELECT * FROM orders WHERE id = ?", (order_id,)).fetchone()

        recheck = evaluate_order_modification_eligibility(conn, locked_order, user, lock=pg)
        if not recheck["can_modify"]:
            conn.rollback()
            code = recheck["code"] or "forbidden"
            return (
                {"success": False, "error": recheck["block_reason"] or "订单当前不可修改", "code": code},
                403 if code == "site_owner_forbidden" else 409,
            )

        if use_ledger:
            op = begin_operation(
                conn,
                operation_type=OPERATION_TYPE,
                order_id=str(order_id),
                site_id=int(site["id"]),
                request_payload={
                    "line_item_id": line_item_id,
                    "replace_qty": replace_qty,
                    "target_product_id": target_product_id,
                    "target_variation_id": target_variation_id,
                    "target_flavor": chosen_option["flavor_label"],
                    "price_mode": effective_price_mode,
                    "prior_modified": prior_modified,
                },
                created_by=str(getattr(user, "id", "") or ""),
            )
            operation_id = op["operation_id"]
            should_execute = bool(op.get("should_execute"))
            conn.commit()
        else:
            conn.rollback()
    except Exception:
        conn.rollback()
        raise

    # 3. 执行 WooCommerce REST API 订单商品修改
    updated_remote: dict[str, Any] | None = None
    uncertain = False

    if should_execute:
        try:
            put_resp = _wc_request(
                "PUT",
                site,
                f"orders/{woo_id}",
                json_body={"line_items": line_items_patch},
                timeout=(8, 30),
                session=session,
            )
            if put_resp.status_code == 200:
                body = put_resp.json()
                if isinstance(body, dict) and str(body.get("id")) == str(woo_id):
                    updated_remote = body
                else:
                    uncertain = True
            elif 400 <= put_resp.status_code < 500:
                err_msg = f"源站拒绝修改订单商品（HTTP {put_resp.status_code}）"
                try:
                    err_json = put_resp.json()
                    if isinstance(err_json, dict) and err_json.get("message"):
                        err_msg = f"源站拒绝修改：{err_json.get('message')}"
                except Exception:
                    pass
                if use_ledger and operation_id:
                    transition_operation(conn, operation_id, "failed", error=err_msg[:240])
                    conn.commit()
                return {"success": False, "error": err_msg}, 422
            else:
                uncertain = True
        except (requests.RequestException, ValueError, DomainError):
            uncertain = True

    if updated_remote is None:
        # 超时或状态待确认时回读源站订单验证是否已生效
        try:
            verify_resp = _wc_request("GET", site, f"orders/{woo_id}", timeout=(6, 15), session=session)
            if verify_resp.status_code == 200:
                verify_body = verify_resp.json()
                if isinstance(verify_body, dict) and _remote_modification_verified(
                    verify_body,
                    orig_item_id=line_item_id,
                    remain_qty=remain_qty,
                    target_product_id=target_product_id,
                    target_variation_id=target_variation_id,
                    replace_qty=replace_qty,
                    prior_modified=prior_modified,
                ):
                    updated_remote = verify_body
        except Exception:
            pass

    if updated_remote is None:
        if use_ledger and operation_id:
            cur_op = conn.execute(
                "SELECT status FROM external_operations WHERE operation_id = ? FOR UPDATE",
                (operation_id,),
            ).fetchone()
            if cur_op and cur_op["status"] == "pending":
                transition_operation(
                    conn,
                    operation_id,
                    "reconciliation_required" if uncertain else "failed",
                    error="源站订单修改响应超时，结果待核对",
                )
            conn.commit()
        return {
            "success": False,
            "uncertain": uncertain,
            "error": "提交源站修改时响应超时，未能确认结果；请先刷新订单详情查看是否已生效，避免重复修改",
        }, 504

    # 4. 在源站写入一条内部备注留痕（不向客户发邮件）
    operator_name = (
        getattr(user, "name", None) or getattr(user, "username", None) or f"用户#{getattr(user, 'id', '')}"
    ).strip()
    orig_display = _clean_text(orig_item.get("name")) or f"商品#{orig_product_id}"
    orig_flav_sig = _extract_item_flavor_signature(orig_item)
    if orig_flav_sig and orig_flav_sig not in orig_display.casefold():
        orig_display = f"{orig_display} ({orig_flav_sig})"
    target_display = chosen_option["full_name"]
    currency = updated_remote.get("currency") or order.get("currency") or ""
    old_order_total = _dec(remote_order.get("total"))
    new_order_total = _dec(updated_remote.get("total"))

    split_desc = (
        f"整行替换 ×{replace_qty}"
        if remain_qty == 0
        else f"拆分更换 {replace_qty}/{orig_qty}（原商品保留 ×{remain_qty}）"
    )
    price_desc = (
        f"保持原成交金额不变（订单总计 {new_order_total:.2f} {currency}）"
        if effective_price_mode == "keep"
        else f"按新商品站点价同步金额（订单总计 {old_order_total:.2f} -> {new_order_total:.2f} {currency}）"
    )
    note_text = (
        f"[订单系统] 负责人 {operator_name} 修改订单商品（{split_desc}）："
        f"将「{orig_display}」更换为「{target_display}」×{replace_qty}；{price_desc}。"
    )

    remote_note_created = None
    try:
        note_resp = _wc_request(
            "POST",
            site,
            f"orders/{woo_id}/notes",
            json_body={"note": note_text, "customer_note": False},
            timeout=(6, 15),
            session=session,
        )
        if note_resp.status_code in (200, 201):
            parsed_note = note_resp.json()
            if isinstance(parsed_note, dict) and parsed_note.get("id"):
                remote_note_created = parsed_note
    except Exception:
        remote_note_created = None

    # 5. 在单一本地事务中更新本地订单镜像、备注与 OMS 多仓履约库存预留
    fulfillment_summary = None
    try:
        if pg:
            conn.execute(
                "SELECT pg_advisory_xact_lock(?, hashtext(?))",
                (SHIPMENT_LOCK_NAMESPACE, str(order_id)),
            )
            conn.execute("SELECT id FROM orders WHERE id = ? FOR UPDATE", (order_id,)).fetchone()

        _update_local_order_from_remote(conn, str(order_id), updated_remote, site["url"])

        if remote_note_created and _table_exists(conn, "order_notes"):
            try:
                from sync_utils import upsert_order_notes_in_transaction

                local_note = dict(remote_note_created)
                local_note.update(
                    {
                        "_local_order_id": str(order_id),
                        "note": remote_note_created.get("note") or note_text,
                        "date_created": remote_note_created.get(
                            "date_created", datetime.now().isoformat()
                        ),
                        "customer_note": False,
                        "author": operator_name,
                        "added_by_user": True,
                    }
                )
                upsert_order_notes_in_transaction([local_note], conn)
            except Exception:
                pass

        if use_ledger and operation_id:
            cur_op = conn.execute(
                "SELECT status FROM external_operations WHERE operation_id = ? FOR UPDATE",
                (operation_id,),
            ).fetchone()
            if cur_op:
                if cur_op["status"] in {"pending", "reconciliation_required"}:
                    transition_operation(
                        conn,
                        operation_id,
                        "external_success",
                        external_reference=str(woo_id),
                        evidence={"new_total": f"{new_order_total:.2f}"},
                    )
                if cur_op["status"] != "local_committed":
                    transition_operation(conn, operation_id, "local_committed")

        actor = {
            "type": "user",
            "id": getattr(user, "id", None),
            "name": operator_name,
        }

        if _table_exists(conn, "oms_order_fulfillment_state"):
            has_existing_state = (
                conn.execute(
                    "SELECT 1 FROM oms_order_fulfillment_state WHERE order_id = ?",
                    (str(order_id),),
                ).fetchone()
                is not None
            )
            market = str(site.get("country") or "").upper()
            oms_enabled = _setting_enabled(conn, "oms_fulfillment_enabled", default=False)
            if has_existing_state or (oms_enabled and market in {"PL", "CZ", "HU"}):
                plan_res = plan_order(conn, str(order_id), actor=actor, commit=False)
                fulfillment_summary = {
                    "revision": plan_res.get("revision"),
                    "aggregate_status": plan_res.get("aggregate_status"),
                    "shortages": plan_res.get("shortages") or [],
                }

        if _table_exists(conn, "oms_events"):
            record_event(
                conn,
                "order",
                str(order_id),
                "order_item_modified",
                to_status=str(updated_remote.get("status") or order.get("status") or ""),
                actor=actor,
                reason=note_text,
                correlation_id=operation_id,
                payload={
                    "line_item_id": line_item_id,
                    "orig_qty": orig_qty,
                    "replace_qty": replace_qty,
                    "remain_qty": remain_qty,
                    "target_product_id": target_product_id,
                    "target_variation_id": target_variation_id,
                    "target_name": target_display,
                    "price_mode": effective_price_mode,
                    "old_total": f"{old_order_total:.2f}",
                    "new_total": f"{new_order_total:.2f}",
                },
            )

        conn.commit()
    except DomainError as exc:
        conn.rollback()
        return {
            "success": False,
            "error": f"源站订单已更新，但本地履约重算提示：{exc}。请刷新订单详情并在多仓履约核对",
            "code": exc.code,
        }, 409
    except Exception as exc:
        conn.rollback()
        return {
            "success": False,
            "error": f"源站订单已修改成功，但本地同步保存异常：{exc}。请点击快速同步更新本地记录",
        }, 500

    msg = f"已成功将商品更换为「{target_display}」×{replace_qty}，并已同步至站点"
    if fulfillment_summary:
        if fulfillment_summary.get("shortages"):
            msg += "（注意：本地仓库该新口味当前存在缺货/未映射，请在多仓履约查看）"
        else:
            msg += "及仓库预留库存"

    return {
        "success": True,
        "message": msg,
        "old_total": f"{old_order_total:.2f}",
        "new_total": f"{new_order_total:.2f}",
        "currency": currency,
        "price_mode": effective_price_mode,
        "fulfillment": fulfillment_summary,
    }, 200
