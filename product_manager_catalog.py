"""Read-only, directly sourced WooCommerce catalog pages for cross-site search.

This module deliberately does not use product Master routing: each site's
published catalog and each variation's own attributes are the search evidence.
The browser coordinates bounded page reads and owns cross-page completeness.
"""
import html
import math
import re
import unicodedata
from urllib.parse import urlsplit

from flask import Blueprint, current_app, jsonify, request
from flask_login import current_user, login_required
import requests
from werkzeug.exceptions import HTTPException

from product_recognition import parse_product_name


PER_PAGE = 50
_SITE_COLUMNS = (
    "id, url, manager, country, consumer_key, consumer_secret, product_master_id"
)
_BRAND_KEYS = {"brand", "brands", "marka", "znacka", "品牌"}


def _text(value):
    if value is None:
        return ""
    return " ".join(html.unescape(str(value)).split())


def normalize_catalog_text(value):
    """Normalize readable Woo text for case/diacritic insensitive matching."""
    value = re.sub(r"<[^>]*>", " ", _text(value)).casefold()
    value = value.translate(str.maketrans({
        "ł": "l", "ø": "o", "đ": "d", "ð": "d", "þ": "th", "æ": "ae", "œ": "oe",
    }))
    value = "".join(
        c for c in unicodedata.normalize("NFKD", value)
        if not unicodedata.combining(c)
    )
    return " ".join(re.sub(r"[\W_]+", " ", value, flags=re.UNICODE).split())


def _attribute_key(attribute):
    name = attribute.get("slug") or attribute.get("name") or ""
    key = normalize_catalog_text(name)
    if key.startswith("pa "):
        key = key[3:]
    return key.strip()


def _is_flavor_attribute(attribute):
    # Labels such as Flavour Profiles and Hungarian Ízcsoport occur on real
    # sites. Match taxonomy tokens rather than demanding one exact label.
    return any(
        bool(re.search(r"smak|flavo[u]?r|口味|味道|风味|香味|taste|aroma|\b(?:iz|izcsoport|prichut\w*)\b", normalize_catalog_text(value)))
        for value in (attribute.get("name", ""), attribute.get("slug", ""))
    )


def _unique(values):
    seen = set()
    output = []
    for value in values:
        value = _text(value)
        normalized = normalize_catalog_text(value)
        if normalized and normalized not in seen:
            seen.add(normalized)
            output.append(value)
    return output


class CatalogReadError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


def _positive_id(value):
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _validate_item(item, *, variation=False):
    if not isinstance(item, dict) or not _positive_id(item.get("id")):
        raise CatalogReadError("invalid_schema", "WC 返回了无效的商品身份，当前页尚未完整读取。")
    for key in ("name", "sku", "status", "stock_status", "permalink"):
        if item.get(key) is not None and not isinstance(item[key], str):
            raise CatalogReadError("invalid_schema", "WC 返回了无效的商品字段，当前页尚未完整读取。")
    if not _text(item.get("status")):
        raise CatalogReadError("invalid_schema", "WC 商品状态缺失，无法核对目录完整性。")
    if not variation and not isinstance(item.get("type"), str):
        raise CatalogReadError("invalid_schema", "WC 商品类型缺失，当前页尚未完整读取。")
    if not variation and item["type"] == "variable" and "variations" not in item:
        raise CatalogReadError("invalid_schema", "WC 父商品缺少变体身份列表，无法确认目录完整性。")
    attributes = item.get("attributes", [])
    if not isinstance(attributes, list):
        raise CatalogReadError("invalid_schema", "WC 属性格式异常，当前页尚未完整读取。")
    for attribute in attributes:
        if not isinstance(attribute, dict):
            raise CatalogReadError("invalid_schema", "WC 属性格式异常，当前页尚未完整读取。")
        for key in ("name", "slug", "option"):
            if attribute.get(key) is not None and not isinstance(attribute[key], str):
                raise CatalogReadError("invalid_schema", "WC 属性值异常，当前页尚未完整读取。")
        options = attribute.get("options", [])
        if not isinstance(options, list) or any(not isinstance(value, str) for value in options):
            raise CatalogReadError("invalid_schema", "WC 属性选项异常，当前页尚未完整读取。")
    brands = item.get("brands", [])
    if not isinstance(brands, list) or any(
        not isinstance(brand, dict) or not isinstance(brand.get("name", ""), str)
        for brand in brands
    ):
        raise CatalogReadError("invalid_schema", "WC 品牌格式异常，当前页尚未完整读取。")
    if "variations" in item and (
        not isinstance(item["variations"], list)
        or any(not _positive_id(value) for value in item["variations"])
        or len(set(item["variations"])) != len(item["variations"])
    ):
        raise CatalogReadError("invalid_schema", "WC 变体目录异常，当前页尚未完整读取。")


def _read_wc(site, path, params=None):
    """Only GET; credentials go through HTTP Basic auth and are never returned."""
    try:
        response = requests.get(
            f'{site["url"].rstrip("/")}/wp-json/wc/v3/{path}',
            auth=(site["consumer_key"], site["consumer_secret"]),
            params=params,
            timeout=(5, 25),
            headers={"Accept": "application/json", "User-Agent": "Woo-Analysis-Catalog/1.0"},
        )
    except requests.RequestException:
        # Exception strings can include full URLs or proxy credentials.
        raise CatalogReadError("connection_failed", "读取站点失败或超时，当前页尚未完整读取。") from None
    if not 200 <= response.status_code < 300:
        raise CatalogReadError(
            "upstream_http_error",
            f"站点 API 返回 HTTP {response.status_code}，当前页尚未完整读取。",
        )
    try:
        payload = response.json()
    except (ValueError, TypeError):
        raise CatalogReadError("invalid_json", "站点返回网页或无效 JSON，当前页尚未完整读取。") from None
    return payload, response.headers


def _pagination(headers, page, count):
    warnings = []
    values = {}
    for header, key in (("X-WP-Total", "total"), ("X-WP-TotalPages", "total_pages")):
        raw = headers.get(header)
        if raw is None:
            values[key] = None
            warnings.append(f"缺少 {header}。")
        elif not re.fullmatch(r"\d+", str(raw).strip()):
            raise CatalogReadError("invalid_pagination", f"站点分页头 {header} 格式异常，无法确认目录完整性。")
        else:
            values[key] = int(str(raw).strip())
    total, pages = values["total"], values["total_pages"]
    if count > PER_PAGE:
        raise CatalogReadError("invalid_pagination", "站点返回超过 50 条的单页，无法确认目录完整性。")
    if total is not None and pages is not None:
        expected_pages = math.ceil(total / PER_PAGE)
        # WordPress may report one page for an empty collection.
        if pages != expected_pages and not (total == 0 and pages == 1):
            raise CatalogReadError("invalid_pagination", "站点总条数与总页数矛盾，无法确认目录完整性。")
        expected_count = min(PER_PAGE, max(0, total - (page - 1) * PER_PAGE))
        if (page > max(pages, 1)) or count != expected_count:
            raise CatalogReadError("incomplete_page", "站点当前页条数与分页头不一致，目录可能在读取时变化，请重新加载。")
        more = page < pages
        mode = "headers"
    else:
        # Never interpret a missing header as a zero total.
        if total is not None:
            if page > max(math.ceil(total / PER_PAGE), 1):
                raise CatalogReadError("incomplete_page", "请求页超过站点总条数对应的终点，无法确认目录完整性。")
            expected_count = min(PER_PAGE, max(0, total - (page - 1) * PER_PAGE))
            if count != expected_count:
                raise CatalogReadError("incomplete_page", "站点当前页与总条数不一致，无法确认目录完整性。")
        if pages is not None and ((page > max(pages, 1)) or (page < pages and count != PER_PAGE)):
            raise CatalogReadError("incomplete_page", "站点当前页与总页数不一致，无法确认目录完整性。")
        if total is not None:
            more = page < math.ceil(total / PER_PAGE)
            warnings.append("按已知总条数确认分页终点。")
            mode = "total-header"
        elif pages is not None:
            if pages == 0 and count:
                raise CatalogReadError("invalid_pagination", "站点总页数为零但返回了商品，无法确认目录完整性。")
            more = page < pages
            warnings.append("按已知总页数确认分页终点。")
            mode = "pages-header"
        else:
            more = count == PER_PAGE
            warnings.append("按每页 50 条继续读取，短页结束；目录完整性依据分页长度。")
            mode = "length-fallback"
    return {
        "total": total, "total_pages": pages, "has_more": more,
        "next_page": page + 1 if more else None,
        "pagination_mode": mode, "warnings": warnings,
    }


def _attribute_identities(attribute):
    keys = set()
    for field in ("name", "slug"):
        value = normalize_catalog_text(attribute.get(field, ""))
        if value.startswith("pa "):
            value = value[3:]
        if value:
            keys.add(value)
    return keys


def _attributes(item, parent=None):
    parent_attributes = (parent or {}).get("attributes", [])
    parent_by_id = {a.get("id"): a for a in parent_attributes if a.get("id")}
    result = []
    for attribute in item.get("attributes", []):
        parent_attribute = parent_by_id.get(attribute.get("id"), {})
        if parent and not parent_attribute:
            identities = _attribute_identities(attribute)
            candidates = [
                candidate for candidate in parent_attributes
                if identities & _attribute_identities(candidate)
            ]
            if len(candidates) == 1:
                parent_attribute = candidates[0]
        # Woo variation attributes may omit the taxonomy name but retain its ID.
        name = _text(attribute.get("name") or parent_attribute.get("name"))
        slug = _text(attribute.get("slug") or parent_attribute.get("slug"))
        if parent and not name and not slug:
            raise CatalogReadError("invalid_schema", "WC 变体属性无法对应父商品属性，当前页尚未完整读取。")
        values = [attribute.get("option")] if "option" in attribute else attribute.get("options", [])
        resolved = {"id": attribute.get("id", 0), "name": name, "slug": slug, "values": _unique(values), "wildcard": False}
        if parent and "option" in attribute and not _text(attribute["option"]) and _is_flavor_attribute(resolved):
            # Woo's empty variation option means this same variation supports
            # every parent option for that dimension. Unlike sibling options,
            # these are genuine leaf capabilities and share one product ID.
            supported = _unique(parent_attribute.get("options", []))
            if not supported:
                raise CatalogReadError("unresolved_flavor_wildcard", "任意口味变体缺少可核实的父商品口味选项，当前页尚未完整读取。")
            resolved["values"] = supported
            resolved["wildcard"] = True
        result.append(resolved)
    return result


def _build_row(site_id, item, parent, recognition):
    attributes = _attributes(item, parent)
    brand_sources = list(item.get("brands", [])) + list((parent or {}).get("brands", []))
    brands = [brand.get("name") for brand in brand_sources]
    for attribute in attributes + (_attributes(parent) if parent else []):
        if _attribute_key(attribute) in _BRAND_KEYS:
            brands.extend(attribute["values"])
    flavors = []
    for attribute in attributes:
        if _is_flavor_attribute(attribute):
            flavors.extend(attribute["values"])
    product_name = _text((parent or item).get("name"))
    leaf_name = _text(item.get("name"))
    brands_cache, series_cache = recognition
    parsed = parse_product_name(leaf_name or product_name, brands_cache, series_cache)
    parent_parsed = parse_product_name(product_name, brands_cache, series_cache) if parent else {}
    if parsed.get("brand"):
        brands.append(parsed["brand"])
    elif parent_parsed.get("brand"):
        brands.append(parent_parsed["brand"])
    # A variable parent's inferred flavor can describe a sibling or the entire
    # range. Only the leaf's own name can supply inferred variation flavor.
    if not flavors and parsed.get("flavor") and (not parent or (leaf_name and leaf_name != product_name)):
        flavors.append(parsed["flavor"])
    attr_values = [value for attribute in attributes for value in attribute["values"]]
    display_name = leaf_name or product_name
    if parent and attr_values and (not leaf_name or leaf_name == product_name):
        display_name = f'{product_name} — {" / ".join(attr_values)}'
    parent_status = _text((parent or {}).get("status"))
    own_status = _text(item.get("status"))
    effective_status = parent_status if parent and parent_status != "publish" else own_status or parent_status
    return {
        "site_id": site_id,
        "product_id": (parent or item)["id"],
        "variation_id": item["id"] if parent else 0,
        "product_name": product_name,
        "name": display_name,
        "sku": _text(item.get("sku")),
        "brands": _unique(brands), "flavors": _unique(flavors),
        "flavor_scope": "any" if any(a["wildcard"] and _is_flavor_attribute(a) for a in attributes) else "specific",
        "attributes": attributes,
        "type": "variation" if parent else item["type"],
        "status": effective_status, "parent_status": parent_status,
        "variation_status": own_status if parent else "",
        "stock_state": _text(item.get("stock_status")),
        "stock_status": _text(item.get("stock_status")),
        "stock_quantity": item.get("stock_quantity"),
        "manage_stock": bool(item.get("manage_stock", False)),
        "price": _text(item.get("price")),
        "regular_price": _text(item.get("regular_price")),
        "sale_price": _text(item.get("sale_price")),
        "permalink": _text(item.get("permalink") or (parent or {}).get("permalink")),
        "puffs": parsed.get("puffs") or parent_parsed.get("puffs"),
        "series": parsed.get("series") or parent_parsed.get("series"),
    }


def _matches(row, keyword, parent=None, item=None):
    if not keyword:
        return True
    evidence = [row["name"], row["product_name"], row["sku"]]
    evidence.extend(value for attribute in row["attributes"] for value in attribute["values"])
    if parent:
        parent_flavor_options = [
            value for attribute in parent.get("attributes", [])
            if _is_flavor_attribute(attribute)
            and (attribute.get("variation") or len(attribute.get("options", [])) > 1)
            for value in attribute.get("options", [])
        ]
        if any(keyword in normalize_catalog_text(value) for value in parent_flavor_options):
            # Woo often prepends the parent title to all variation names. When
            # the query names a flavor dimension, remove that inherited title
            # before testing the leaf name; siblings need their own evidence.
            own_name = normalize_catalog_text((item or {}).get("name", ""))
            inherited_name = normalize_catalog_text(row["product_name"])
            own_name = own_name.replace(inherited_name, "", 1) if inherited_name else own_name
            evidence = [row["sku"], own_name]
            evidence.extend(value for attribute in row["attributes"] for value in attribute["values"])
    return any(keyword in normalize_catalog_text(value) for value in evidence)


def _configured(site):
    try:
        url = urlsplit(site["url"] or "")
    except ValueError:
        return False
    return bool(
        url.scheme in {"http", "https"} and url.hostname and not url.username and not url.password
        and not url.query and not url.fragment and _text(site["consumer_key"]) and _text(site["consumer_secret"])
    )


def _public_site_url(value):
    """A misconfigured site URL must not expose userinfo or query credentials."""
    try:
        url = urlsplit(value or "")
        if not url.hostname:
            return ""
        host = f"[{url.hostname}]" if ":" in url.hostname else url.hostname
        authority = f"{host}:{url.port}" if url.port else host
        return f"{url.scheme}://{authority}{url.path}".rstrip("/")
    except ValueError:
        return ""


def create_catalog_blueprint(get_db_connection, product_manager_required, recognition_loader=None):
    """Build API routes without importing the application's mutable global state.

    recognition_loader is an optional zero-argument callback returning either
    (brands_cache, series_cache) or {"brands": [...], "series": [...]}.
    """
    blueprint = Blueprint("product_manager_catalog", __name__)

    @blueprint.before_request
    def require_json_login():
        if not current_user.is_authenticated:
            return jsonify({"error": "请先登录。", "code": "authentication_required"}), 403

    @blueprint.errorhandler(Exception)
    def json_failure(exc):
        if isinstance(exc, HTTPException):
            return jsonify({"error": "请求无法处理。", "code": "request_failed"}), exc.code
        current_app.logger.exception("Catalog request could not be completed")
        return jsonify({"error": "目录读取暂时失败，请重新加载。", "code": "catalog_read_failed", "complete_page": False})

    def visible(site):
        if not current_user.product_manager_own_scoped():
            return True
        user_name = str(current_user.name or "").strip()
        return bool(user_name and user_name == str(site["manager"] or "").strip())

    def load_recognition():
        if recognition_loader is None:
            return [], []
        try:
            result = recognition_loader()
            if isinstance(result, dict):
                return result.get("brands", []), result.get("series", [])
            return result
        except Exception:
            current_app.logger.exception("Catalog recognition rules unavailable")
            raise CatalogReadError("recognition_unavailable", "产品识别规则暂时不可用，请重新加载。") from None

    @blueprint.get("/api/product-manager/catalog-sites")
    @login_required
    @product_manager_required
    def catalog_sites():
        conn = get_db_connection()
        try:
            sites = conn.execute(f"SELECT {_SITE_COLUMNS} FROM sites ORDER BY country, url").fetchall()
        finally:
            conn.close()
        return jsonify({"sites": [
            {
                "id": site["id"], "url": _public_site_url(site["url"]), "manager": site["manager"] or "",
                "country": site["country"] or "", "configured": _configured(site),
                "product_master_id": site["product_master_id"],
            }
            for site in sites if visible(site)
        ]})

    @blueprint.get("/api/product-manager/catalog-page")
    @login_required
    @product_manager_required
    def catalog_page():
        try:
            site_id = int(request.args.get("site_id", "0"))
            page = int(request.args.get("page", "1"))
            parent_id = int(request.args.get("parent_id", "0"))
            if site_id < 1 or page < 1 or parent_id < 0:
                raise ValueError
        except (ValueError, TypeError):
            return jsonify({"error": "site_id、page 必须为正整数，parent_id 必须为非负整数。", "code": "invalid_arguments"}), 400
        search = request.args.get("search", "")
        if len(search) > 200:
            return jsonify({"error": "查询词不能超过 200 个字符。", "code": "invalid_arguments"}), 400
        conn = get_db_connection()
        try:
            row = conn.execute(f"SELECT {_SITE_COLUMNS} FROM sites WHERE id = ?", (site_id,)).fetchone()
            site = dict(row) if row else None
        finally:
            conn.close()
        if site is None:
            return jsonify({"error": "站点不存在。", "code": "site_not_found"}), 404
        if not visible(site):
            return jsonify({"error": "仅站点负责人可读取该站点产品。", "code": "site_forbidden"}), 403
        response_base = {
            "site_id": site_id, "parent_id": parent_id, "page": page, "per_page": PER_PAGE,
            "rows": [], "variable_products": [], "complete_page": False,
        }
        try:
            if not _configured(site):
                raise CatalogReadError("site_unconfigured", "该站点未配置完整的直接读取 URL 和 WC API 凭据。")
            parent = None
            if parent_id:
                parent, _ = _read_wc(site, f"products/{parent_id}")
                _validate_item(parent)
                if parent["id"] != parent_id or parent["type"] != "variable":
                    raise CatalogReadError("invalid_parent", "站点父商品身份或类型不一致，变体尚未完整读取。")
            path = f"products/{parent_id}/variations" if parent_id else "products"
            items, headers = _read_wc(site, path, {
                "page": page, "per_page": PER_PAGE, "status": "any", "orderby": "id", "order": "asc",
            })
            if not isinstance(items, list):
                raise CatalogReadError("invalid_schema", "WC 返回的目录不是数组，当前页尚未完整读取。")
            ids = set()
            for item in items:
                _validate_item(item, variation=bool(parent))
                if item["id"] in ids:
                    raise CatalogReadError("duplicate_id", "WC 当前页有重复商品身份，无法确认目录完整性。")
                ids.add(item["id"])
            pagination = _pagination(headers, page, len(items))
            recognition = load_recognition()
            rows = []
            variables = []
            keyword = normalize_catalog_text(search)
            for item in items:
                if not parent and item["type"] == "variable":
                    variables.append({
                        "id": item["id"], "name": _text(item.get("name")), "sku": _text(item.get("sku")),
                        "variations_count": len(item.get("variations", [])),
                        "variation_ids": item.get("variations", []),
                    })
                else:
                    leaf = _build_row(site_id, item, parent, recognition)
                    if _matches(leaf, keyword, parent, item):
                        rows.append(leaf)
            return jsonify({
                **response_base, **pagination, "rows": rows, "variable_products": variables,
                "source_ids": [item["id"] for item in items], "scanned": len(items), "complete_page": True,
                "source_statuses": {str(item["id"]): _text(item["status"]) for item in items},
            })
        except CatalogReadError as exc:
            # HTTP 200 preserves actionable JSON through proxy/CDN error pages.
            return jsonify({**response_base, "error": str(exc), "code": exc.code})
        except Exception:
            current_app.logger.exception("Catalog page read failed for site %s parent %s page %s", site_id, parent_id, page)
            return jsonify({**response_base, "error": "当前页读取失败，请重新加载；目录尚未完整读取。", "code": "catalog_read_failed"})

    return blueprint
