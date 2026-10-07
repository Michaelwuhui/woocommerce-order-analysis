"""Exact-child product edits and durable clones from cross-site catalog rows."""
from decimal import Decimal, InvalidOperation
import html
import json
import secrets
import uuid

from flask import Blueprint, current_app, jsonify, request, session
from flask_login import current_user, login_required
import requests

from product_manager_catalog import _build_row, _configured, _public_site_url
from product_manager_service import (
    PRODUCT_EDIT_FIELDS, new_product_operation_deadline, product_payload_mismatches,
    product_state_snapshot, wc_product_update_verified, wcms_stock_meta_update,
    _read_product, _request_timeout, parse_wc_response, ProductOperationExpired,
)


SITE_FIELDS = "id,url,manager,country,consumer_key,consumer_secret,product_master_id"
CSRF_KEY = "product_catalog_edit_csrf"


class CatalogEditError(Exception):
    def __init__(self, code, message, status=400):
        super().__init__(message)
        self.code, self.status = code, status


def catalog_edit_csrf_token():
    if CSRF_KEY not in session:
        session[CSRF_KEY] = secrets.token_urlsafe(32)
    return session[CSRF_KEY]


def _id(value, zero=False):
    if type(value) is not int or value < (0 if zero else 1):
        raise CatalogEditError("INVALID_INPUT", "商品和站点 ID 必须为有效整数。")
    return value


def build_catalog_changes(changes):
    if not isinstance(changes, dict) or not changes or set(changes) - PRODUCT_EDIT_FIELDS:
        raise CatalogEditError("INVALID_CHANGES", "仅支持库存管理、数量、库存状态、原价和优惠价。")
    payload = dict(changes)
    if "manage_stock" in payload and type(payload["manage_stock"]) is not bool:
        raise CatalogEditError("INVALID_CHANGES", "库存管理必须为 true 或 false。")
    if "stock_quantity" in payload and (type(payload["stock_quantity"]) is not int or not 0 <= payload["stock_quantity"] <= 2147483647):
        raise CatalogEditError("INVALID_CHANGES", "库存数量必须为 0 至 2147483647 的整数。")
    if "stock_status" in payload and (not isinstance(payload["stock_status"], str) or payload["stock_status"] not in {"instock", "outofstock", "onbackorder"}):
        raise CatalogEditError("INVALID_CHANGES", "库存状态无效。")
    for field in ("regular_price", "sale_price"):
        if field not in payload:
            continue
        value = payload[field]
        if not isinstance(value, (str, int, float)) or isinstance(value, bool):
            raise CatalogEditError("INVALID_CHANGES", "价格必须为有限的非负数；明确传空字符串才会清空。")
        if value == "":
            continue
        try:
            text = str(value).strip()
            if len(text) > 64:
                raise InvalidOperation
            number = Decimal(text)
            if (not number.is_finite() or number < 0 or number.adjusted() > 12
                    or number.as_tuple().exponent < -8 or number.as_tuple().exponent > 12):
                raise InvalidOperation
        except (InvalidOperation, ValueError):
            raise CatalogEditError("INVALID_CHANGES", "价格必须为有限的非负数。") from None
        payload[field] = format(number, "f")
    return payload


def catalog_item_identity(item, parent_id=0):
    attributes = item.get("attributes", [])
    if not isinstance(attributes, list) or any(not isinstance(a, dict) for a in attributes):
        raise CatalogEditError("INVALID_REMOTE_IDENTITY", "站点商品属性格式异常，未提交修改。", 409)
    own_attributes = []
    for attribute in attributes:
        if "option" in attribute:
            values = [attribute["option"]]
        else:
            values = attribute.get("options", [])
        if not isinstance(values, list) or any(not isinstance(value, str) for value in values):
            raise CatalogEditError("INVALID_REMOTE_IDENTITY", "站点商品属性值异常，未提交修改。", 409)
        own_attributes.append({
            "name": html.unescape(str(attribute.get("name") or attribute.get("slug") or "")).strip(),
            "id": attribute.get("id", 0), "values": sorted(html.unescape(value).strip() for value in values),
        })
    own_attributes.sort(key=lambda attribute: json.dumps(attribute, sort_keys=True, ensure_ascii=False))
    return {"sku": str(item.get("sku") or "").strip(),
            "type": "variation" if parent_id else item.get("type"),
            "parent_id": parent_id, "attributes": own_attributes}


def catalog_item_before(item):
    return {field: item.get(field) for field in sorted(PRODUCT_EDIT_FIELDS)}


def catalog_stock_bridge(item):
    """Expose only the three public stock bridge flags for read-only recovery."""
    keys = {"wcms_stock_manage", "wcms_stock_qty", "wcms_stock_status"}
    metadata = item.get("meta_data") or []
    if not isinstance(metadata, list):
        return {"present": False, "valid": False, "values": {}}
    values, valid = {}, True
    for row in metadata:
        if not isinstance(row, dict) or not isinstance(row.get("key"), str):
            valid = False
            continue
        if row["key"] not in keys:
            continue
        value = row.get("value")
        scalar = value is None or type(value) in {str, int, float, bool}
        if type(value) is float and not Decimal(str(value)).is_finite():
            scalar = False
        if row["key"] in values:
            valid = False
            continue
        if not scalar or len(str(value)) > 64:
            valid = False
            value = None
        values[row["key"]] = value
    return {"present": "wcms_stock_manage" in values, "valid": valid, "values": values}


def _same_value(field, expected, actual):
    if field in {"regular_price", "sale_price"}:
        if expected in (None, "") or actual in (None, ""):
            return expected in (None, "") and actual in (None, "")
        try:
            first, second = Decimal(str(expected)), Decimal(str(actual))
            return first.is_finite() and second.is_finite() and first == second
        except InvalidOperation:
            return False
    if field == "manage_stock":
        return type(expected) is type(actual) and expected == actual
    if field == "stock_quantity":
        return type(expected) is type(actual) and expected == actual
    return expected == actual


def _check_expected(item, parent_id, identity, before, changes):
    if catalog_item_identity(item, parent_id) != identity:
        raise CatalogEditError("STALE_IDENTITY", "商品 SKU、类型或属性已变化，请刷新后重新确认。", 409)
    if any(field not in before or not _same_value(field, before[field], item.get(field)) for field in changes):
        raise CatalogEditError("STALE_STATE", "拟修改字段已被其他操作更新，请刷新后重新确认。", 409)


def _check_variation_parent(item, parent_id, parent_url):
    has_parent_id = "parent_id" in item
    if has_parent_id and (type(item["parent_id"]) is not int or item["parent_id"] != parent_id):
        raise CatalogEditError("RESOURCE_MISMATCH", "变体的实际父商品与选择不一致。", 409)
    links = item.get("_links")
    if links is None:
        if not has_parent_id:
            raise CatalogEditError("PARENT_EVIDENCE_MISSING", "变体缺少父商品身份和链接证据，未确认目标身份。", 409)
        return
    if not isinstance(links, dict):
        raise CatalogEditError("RESOURCE_MISMATCH", "变体父商品链接格式异常，未确认目标身份。", 409)
    if "up" not in links:
        if not has_parent_id:
            raise CatalogEditError("PARENT_EVIDENCE_MISSING", "变体缺少父商品身份和链接证据，未确认目标身份。", 409)
        return
    up = links["up"]
    if (not isinstance(up, list) or len(up) != 1 or not isinstance(up[0], dict)
            or not isinstance(up[0].get("href"), str)
            or up[0]["href"].strip().rstrip("/") != parent_url.rstrip("/")):
        raise CatalogEditError("RESOURCE_MISMATCH", "变体父商品链接与已核验实际站点不一致。", 409)


class _GuardedRequests:
    """Recheck the exact identity and selected old values at writer preflight."""
    def __init__(self, client, resource_url, parent_id, identity, before, changes, permission_check):
        self.client, self.resource_url, self.parent_id = client, resource_url, parent_id
        self.identity, self.before, self.changes = identity, before, changes
        self.permission_check = permission_check
        self.RequestException = client.RequestException
        self.has_put = False
        self.last_item = None
        self.expected_id = int(resource_url.rsplit("/", 1)[1])

    def get(self, url, **kwargs):
        response = self.client.get(url, **kwargs)
        if url == self.resource_url and response.status_code == 200:
            try:
                item = response.json()
            except ValueError:
                return response
            if isinstance(item, dict):
                self.last_item = item
                if (type(item.get("id")) is not int or item["id"] != self.expected_id
                        or (self.parent_id and "parent_id" in item and (type(item["parent_id"]) is not int or item["parent_id"] != self.parent_id))):
                    raise CatalogEditError("RESOURCE_MISMATCH", "核验响应商品 ID 或父商品 ID 不一致。", 409)
                if self.parent_id and item.get("type") not in {None, "variation"}:
                    raise CatalogEditError("RESOURCE_MISMATCH", "变体核验响应类型不一致。", 409)
                if self.parent_id:
                    _check_variation_parent(item, self.parent_id, self.resource_url.rsplit("/variations/", 1)[0])
                if self.has_put and catalog_item_identity(item, self.parent_id) != self.identity:
                    raise CatalogEditError("STALE_IDENTITY", "写入后的商品身份发生变化，结果尚未确认。", 409)
        return response

    def put(self, url, **kwargs):
        self.permission_check()
        if url != self.resource_url:
            raise CatalogEditError("RESOURCE_MISMATCH", "写入目标与已确认站点商品不一致。", 409)
        self.has_put = True
        return self.client.put(url, **kwargs)


def create_catalog_edit_blueprint(get_db_connection, product_manager_required,
                                  audit_writer, enqueue_clone_job=None, make_clone_suffix=None,
                                  requests_client=None, recognition_loader=None):
    client = requests_client or requests
    blueprint = Blueprint("product_manager_catalog_edit", __name__)

    @blueprint.before_request
    def boundary():
        if not current_user.is_authenticated:
            return jsonify(success=False, code="AUTH_REQUIRED", error="请先登录。", write_started=False), 403
        if request.method not in {"GET", "HEAD", "OPTIONS"}:
            supplied = request.headers.get("X-PM-CSRF", "")
            expected = session.get(CSRF_KEY, "")
            if not supplied or not expected or not secrets.compare_digest(supplied, expected):
                return jsonify(success=False, code="CSRF_REJECTED", error="页面已过期，请刷新后重试。", write_started=False), 403

    @blueprint.errorhandler(CatalogEditError)
    def known_error(exc):
        result = {"success": False, "code": exc.code, "error": str(exc), "write_started": getattr(exc, "write_started", False)}
        if hasattr(exc, "write_started"):
            result["write_started"] = exc.write_started
            result["verification_status"] = "unconfirmed" if exc.write_started else "failed"
        return jsonify(result), exc.status

    @blueprint.errorhandler(Exception)
    def unknown_error(_exc):
        current_app.logger.exception("Direct catalog operation failed")
        return jsonify(success=False, code="CATALOG_OPERATION_FAILED", error="操作暂时无法完成，请核对结果后重试。"), 200

    def own_site(site):
        if not current_user.product_manager_own_scoped():
            return True
        name = str(current_user.name or "").strip()
        return bool(name and name == str(site["manager"] or "").strip())

    def load_site(site_id, configured=True):
        connection = get_db_connection()
        try:
            actor = connection.execute("SELECT username,name,role,can_manage_products,can_manage_own_products FROM users WHERE id=?", (current_user.id,)).fetchone()
            row = connection.execute(f"SELECT {SITE_FIELDS} FROM sites WHERE id=?", (site_id,)).fetchone()
        finally:
            connection.close()
        if not actor or (actor["username"] != "admin" and actor["can_manage_products"] != 1):
            raise CatalogEditError("PERMISSION_REVOKED", "产品管理权限已撤销，未继续提交修改。", 403)
        if row is None:
            raise CatalogEditError("SITE_NOT_FOUND", "站点不存在。", 404)
        site = dict(row)
        scoped = actor["username"] != "admin" and (actor["role"] != "admin" or actor["can_manage_own_products"] == 1)
        name = str(actor["name"] or "").strip()
        if scoped and (not name or name != str(site["manager"] or "").strip()):
            raise CatalogEditError("SITE_FORBIDDEN", "仅站点负责人可修改该站点商品。", 403)
        if configured and not _configured(site):
            raise CatalogEditError("SITE_UNCONFIGURED", "实际站点未配置完整的直接读取和写入凭据。", 409)
        return site

    def recognition():
        return recognition_loader() if recognition_loader else ([], [])

    def read(site, product_id, variation_id, deadline=None):
        parent_url = f'{site["url"].rstrip("/")}/wp-json/wc/v3/products/{product_id}'
        def fetch(url, identity):
            try:
                response = client.get(url, auth=(site["consumer_key"], site["consumer_secret"]),
                                      timeout=_request_timeout(deadline) if deadline is not None else (5, 20),
                                      headers={"Accept": "application/json", "Cache-Control": "no-cache"})
            except (client.RequestException, ProductOperationExpired):
                raise CatalogEditError("READ_FAILED", "实际站点读取失败，未提交修改。", 409) from None
            if response.status_code != 200:
                raise CatalogEditError("READ_FAILED", f"实际站点读取返回 HTTP {response.status_code}，未提交修改。", 409)
            try:
                value = response.json()
            except ValueError:
                raise CatalogEditError("READ_FAILED", "实际站点未返回有效 JSON，未提交修改。", 409) from None
            if not isinstance(value, dict) or type(value.get("id")) is not int or value["id"] != identity:
                raise CatalogEditError("RESOURCE_MISMATCH", "实际站点商品身份不一致，未提交修改。", 409)
            return value
        parent = fetch(parent_url, product_id)
        if variation_id:
            if parent.get("type") != "variable":
                raise CatalogEditError("RESOURCE_MISMATCH", "目标父商品不是可变商品。", 409)
            resource_url = f"{parent_url}/variations/{variation_id}"
            item = fetch(resource_url, variation_id)
            # Standard Woo v3 omits parent_id from variation responses. The
            # exact verified parent route plus exact leaf ID still identifies it.
            _check_variation_parent(item, product_id, parent_url)
            if item.get("type") not in {None, "variation"}:
                raise CatalogEditError("RESOURCE_MISMATCH", "目标响应不是变体。", 409)
            return parent, item, resource_url
        if parent.get("type") not in {"simple", "variable", "external", "grouped"}:
            raise CatalogEditError("RESOURCE_MISMATCH", "目标商品类型无效。", 409)
        return None, parent, parent_url

    def routing(site):
        url = _public_site_url(site["url"])
        return {"mode": "direct", "site_id": site["id"], "site_url": url,
                "effective_url": url, "is_routed_to_master": False}

    def result(site, parent, item):
        return {"success": True, "identity": catalog_item_identity(item, parent["id"] if parent else 0),
                "before": catalog_item_before(item), "state": product_state_snapshot(item),
                "item": _build_row(site["id"], item, parent, recognition()), "routing": routing(site),
                "stock_scope": "parent" if item.get("manage_stock") == "parent" else "item",
                "stock_bridge": catalog_stock_bridge(item),
                "csrf_token": catalog_edit_csrf_token()}

    def audit(site, product_id, variation_id, payload, batch_id, state, status, trace, error=None):
        audit_writer(batch_id=batch_id, site={**site, "product_master_id": None}, effective_url=site["url"].rstrip("/"),
                     product_id=variation_id or product_id, parent_id=product_id if variation_id else None,
                     payload=payload, final_state=state, status=status, error=error,
                     child_verification={"status": "not_applicable", "detail": "直接写入所选实际站点"},
                     trace={**trace, "routing_mode": "direct", "configured_master_id": site["product_master_id"]})

    @blueprint.get("/api/product-manager/catalog-edit-config")
    @login_required
    @product_manager_required
    def config():
        connection = get_db_connection()
        try:
            sites = [dict(row) for row in connection.execute(f"SELECT {SITE_FIELDS} FROM sites ORDER BY country,url").fetchall()]
        finally:
            connection.close()
        visible = [{"id": site["id"], "url": _public_site_url(site["url"]),
                    "manager": site["manager"] or "", "country": site["country"] or "",
                    "can_edit": _configured(site), "can_clone": _configured(site)} for site in sites if own_site(site)]
        return jsonify(csrf_token=catalog_edit_csrf_token(), sites=visible,
                       clone_targets=[site for site in visible if site["can_clone"]])

    @blueprint.route("/api/product-manager/catalog-item", methods=["GET", "PUT"])
    @login_required
    @product_manager_required
    def item_api():
        if request.method == "GET":
            try:
                identifiers = {key: int(request.args.get(key, "0")) for key in ("site_id", "product_id", "variation_id")}
            except (ValueError, TypeError):
                raise CatalogEditError("INVALID_INPUT", "站点与商品 ID 无效。") from None
            body = identifiers
        else:
            body = request.get_json(silent=True)
            if not isinstance(body, dict):
                raise CatalogEditError("INVALID_INPUT", "请提交 JSON 对象。")
        site_id, product_id, variation_id = (_id(body.get("site_id")), _id(body.get("product_id")), _id(body.get("variation_id", 0), True))
        changes = None
        if request.method == "PUT":
            changes = build_catalog_changes(body.get("changes"))
            if not isinstance(body.get("expected_identity"), dict) or not isinstance(body.get("expected_before"), dict):
                raise CatalogEditError("EXPECTED_STATE_REQUIRED", "请先读取当前商品并确认身份和拟修改字段。")
        site = load_site(site_id)
        operation_deadline = new_product_operation_deadline() if request.method == "PUT" else None
        parent, item, resource_url = read(site, product_id, variation_id, operation_deadline)
        if request.method == "GET":
            return jsonify(result(site, parent, item))
        if item.get("type") in {"variable", "grouped", "external"} and not variation_id:
            raise CatalogEditError("LEAF_REQUIRED", "跨站编辑仅支持简单商品和明确选择的变体，不会自动修改兄弟变体。")
        _check_expected(item, product_id if variation_id else 0, body["expected_identity"], body["expected_before"], changes)
        batch_id = str(body.get("batch_id") or uuid.uuid4())[:100]
        direct_payload = dict(changes)
        metadata = []
        stock_changes = bool({"manage_stock", "stock_quantity", "stock_status"}.intersection(changes))
        bridge_required = False
        def recheck():
            latest = load_site(site_id)
            if any(latest[field] != site[field] for field in ("url", "consumer_key", "consumer_secret")):
                raise CatalogEditError("SITE_CHANGED", "站点配置已变化，请刷新后重新确认。", 409)
        guarded = _GuardedRequests(client, resource_url, product_id if variation_id else 0,
                                   body["expected_identity"], body["expected_before"], changes, recheck)
        def preflight(fresh):
            nonlocal bridge_required
            _check_expected(fresh, product_id if variation_id else 0, body["expected_identity"], body["expected_before"], changes)
            if (fresh.get("manage_stock") == "parent"
                    and {"manage_stock", "stock_quantity", "stock_status"}.intersection(changes)
                    and not (changes.get("manage_stock") is True and "stock_quantity" in changes)):
                raise CatalogEditError("INHERITED_STOCK", "该变体共用父商品库存。请明确启用独立库存并输入数量后再修改；本次未写入。", 409)
            bridge = catalog_stock_bridge(fresh)
            if stock_changes and not bridge["valid"]:
                raise CatalogEditError("BRIDGE_INVALID", "库存桥接元数据重复或格式异常，本次未写入。", 409)
            bridge_required = stock_changes and bridge["present"]
            if bridge_required:
                # Existing child bridges may restore public fields after each
                # save. Carry desired manage/quantity with the first core PUT;
                # any implicit status remains provisional until fresh GET.
                direct_payload["meta_data"] = wcms_stock_meta_update(fresh, {**fresh, **changes}, changes)

        def postflight(core, trace, deadline):
            """Reconcile bridge flags from authoritative Woo state, inside lease."""
            nonlocal metadata, bridge_required
            if not stock_changes:
                return core, None, trace
            headers = {"Accept": "application/json", "Content-Type": "application/json"}
            auth = (site["consumer_key"], site["consumer_secret"])
            stock_fields = ("manage_stock", "stock_quantity", "stock_status")

            def derive(current):
                nonlocal metadata
                if (type(current.get("manage_stock")) is not bool
                        or (current.get("stock_quantity") is not None and type(current["stock_quantity"]) is not int)
                        or not isinstance(current.get("stock_status"), str)
                        or current["stock_status"] not in {"instock", "outofstock", "onbackorder"}):
                    raise CatalogEditError("BRIDGE_CORE_INVALID", "最终库存状态格式无法核实，未继续修改桥接。", 409)
                contract = {"meta_data": [{"key": "wcms_stock_manage", "value": "yes"}]}
                metadata = wcms_stock_meta_update(contract, current, changes)
                expected = {row["key"]: row["value"] for row in metadata}
                trace["expected_stock_bridge"] = expected
                return expected

            def matches(current, expected):
                bridge = catalog_stock_bridge(current)
                return bridge["valid"] and bridge["present"] and all(
                    str(bridge["values"].get(key)) == str(value) for key, value in expected.items())

            def failed(current, message):
                trace.update(stock_bridge_error=True, unconfirmed_write=guarded.has_put,
                             final_state=product_state_snapshot(current) if current else None)
                return current, message, trace

            bridge = catalog_stock_bridge(core)
            bridge_required = bridge_required or bridge["present"]
            if not bridge_required:
                return core, None, trace
            expected = derive(core)
            if any(not _same_value(field, value, core.get(field)) for field, value in changes.items()):
                return failed(core, "核心字段尚未达到目标，未继续修改库存桥接。")
            if not bridge["valid"] or not bridge["present"]:
                return failed(core, "库存桥接证据缺失、重复或格式异常，结果未确认。")
            if matches(core, expected):
                trace["stock_bridge_verified"] = True
                return core, None, trace

            latest, error = _read_product(guarded, resource_url, auth, headers, deadline)
            if error:
                return failed(latest, "桥接写入前回读失败：" + error)
            expected = derive(latest)
            if (any(not _same_value(field, core.get(field), latest.get(field)) for field in stock_fields)
                    or any(not _same_value(field, value, latest.get(field)) for field, value in changes.items())):
                return failed(latest, "核心库存状态在桥接写入前发生变化，未继续写入。")
            evidence = catalog_stock_bridge(latest)
            if not evidence["valid"] or not evidence["present"]:
                return failed(latest, "桥接证据已变化，未继续写入。")
            if matches(latest, expected):
                trace.update(stock_bridge_verified=True, final_state=product_state_snapshot(latest))
                return latest, None, trace

            bridge_trace = {"payload": {"meta_data": metadata}, "http_status": None}
            trace["stock_bridge_phase"] = bridge_trace
            try:
                response = guarded.put(resource_url, auth=auth, json={"meta_data": metadata},
                                       timeout=_request_timeout(deadline), headers=headers)
                _response_item, put_error = parse_wc_response(response)
                bridge_trace["http_status"] = response.status_code
            except (client.RequestException, ProductOperationExpired) as exc:
                put_error = str(exc)
            if put_error:
                bridge_trace["error"] = put_error
            # One metadata PUT only. Even an ambiguous failure is read back,
            # never replayed; the whole stock state must remain unchanged.
            final, error = _read_product(guarded, resource_url, auth, headers, deadline)
            if error:
                return failed(final, "桥接写入后的完整回读失败：" + error)
            expected = derive(final)
            if (any(not _same_value(field, latest.get(field), final.get(field)) for field in stock_fields)
                    or any(not _same_value(field, value, final.get(field)) for field, value in changes.items())):
                return failed(final, "桥接写入后核心字段发生变化，整体结果未确认。")
            if not matches(final, expected):
                return failed(final, "库存桥接未达到最终核心状态，整体结果未确认。")
            trace.update(stock_bridge_verified=True, final_state=product_state_snapshot(final))
            if put_error:
                bridge_trace["verified_after_error"] = True
            return final, None, trace
        try:
            final, error, trace = wc_product_update_verified(guarded, resource_url,
                (site["consumer_key"], site["consumer_secret"]), direct_payload,
                deadline=operation_deadline, preflight_validator=preflight, postflight_hook=postflight)
        except CatalogEditError as exc:
            exc.write_started = guarded.has_put
            audit(site, product_id, variation_id, changes, batch_id, product_state_snapshot(guarded.last_item),
                  "unconfirmed" if guarded.has_put else "failed", {"direct_site": True, "write_started": guarded.has_put}, str(exc))
            raise
        if error:
            audit(site, product_id, variation_id, changes, batch_id, trace.get("final_state"),
                  "unconfirmed" if trace.get("unconfirmed_write") else "failed", trace, error)
            return jsonify(success=False, code="BRIDGE_NOT_VERIFIED" if trace.get("stock_bridge_error") else "WRITE_NOT_VERIFIED", error=error,
                           write_started=guarded.has_put,
                           verification_status="unconfirmed" if trace.get("unconfirmed_write") else "failed",
                           expected_stock_bridge=trace.get("expected_stock_bridge"),
                           stock_bridge=catalog_stock_bridge(final) if final else None,
                           state=trace.get("final_state"), routing=routing(site)), 409
        # The verified writer performs the authoritative GET after every write;
        # verify bridge metadata as well as the public field whitelist.
        if product_payload_mismatches(final, changes) or any(not _same_value(field, value, final.get(field)) for field, value in changes.items()):
            error = "最终回读字段不一致（继承父库存不等于独立库存管理），结果尚未确认。"
            audit(site, product_id, variation_id, changes, batch_id, product_state_snapshot(final), "failed", trace, error)
            return jsonify(success=False, code="WRITE_NOT_VERIFIED", error=error, write_started=guarded.has_put), 409
        final_meta = catalog_stock_bridge(final)["values"]
        if (metadata and not catalog_stock_bridge(final)["valid"]
                or any(str(final_meta.get(row["key"])) != str(row["value"]) for row in metadata)):
            error = "WCMS 库存元数据尚未达到目标状态，请核对结果。"
            audit(site, product_id, variation_id, changes, batch_id, product_state_snapshot(final), "unconfirmed", trace, error)
            return jsonify(success=False, code="BRIDGE_NOT_VERIFIED", error=error, write_started=guarded.has_put,
                           verification_status="unconfirmed", stock_bridge=catalog_stock_bridge(final),
                           expected_stock_bridge={row["key"]: row["value"] for row in metadata}), 409
        audit(site, product_id, variation_id, changes, batch_id, product_state_snapshot(final), "verified", trace)
        return jsonify({**result(site, parent, final), "verification": {"status": "verified", "direct_site": True}})

    @blueprint.post("/api/product-manager/catalog-clone")
    @login_required
    @product_manager_required
    def clone_api():
        if enqueue_clone_job is None or make_clone_suffix is None:
            raise CatalogEditError("CLONE_UNAVAILABLE", "克隆服务暂时不可用。", 409)
        body = request.get_json(silent=True)
        if not isinstance(body, dict):
            raise CatalogEditError("INVALID_INPUT", "请提交 JSON 对象。")
        source_id, target_id = _id(body.get("source_site_id")), _id(body.get("target_site_id"))
        if source_id == target_id:
            raise CatalogEditError("INVALID_INPUT", "源站点和目标站点不能相同。")
        ids = body.get("product_ids")
        if not isinstance(ids, list) or not ids or len(ids) > 50:
            raise CatalogEditError("INVALID_INPUT", "单个克隆任务请选择 1 至 50 个父商品。")
        ids = list(dict.fromkeys(_id(identity) for identity in ids))
        options = {"_catalog_direct_sites": True,
                   "include_variations": body.get("include_variations", True),
                   "include_images": body.get("include_images", True),
                   "status_on_target": body.get("status_on_target", "draft"),
                   "collision_mode": body.get("collision_mode", "skip_existing")}
        if any(type(options[field]) is not bool for field in ("include_variations", "include_images")):
            raise CatalogEditError("INVALID_INPUT", "克隆选项必须为布尔值。")
        if (not isinstance(options["status_on_target"], str) or options["status_on_target"] not in {"draft", "pending", "private", "publish"}
                or not isinstance(options["collision_mode"], str) or options["collision_mode"] not in {"skip_existing", "clone_as_new"}):
            raise CatalogEditError("INVALID_INPUT", "克隆状态或重复商品处理方式无效。")
        if options["collision_mode"] == "clone_as_new":
            options["status_on_target"] = "draft"
            options["clone_sku_suffix"] = make_clone_suffix()
        source, target = load_site(source_id), load_site(target_id)
        try:
            response = client.get(f'{source["url"].rstrip("/")}/wp-json/wc/v3/products',
                auth=(source["consumer_key"], source["consumer_secret"]), timeout=(5, 20),
                params={"include": ",".join(str(identity) for identity in ids), "per_page": 100,
                        "status": "any", "_fields": "id,type,name,sku,attributes"},
                headers={"Accept": "application/json", "Cache-Control": "no-cache"})
            source_items = response.json() if response.status_code == 200 else None
        except (client.RequestException, ValueError):
            source_items = None
        if (not isinstance(source_items, list) or len(source_items) != len(ids)
                or any(not isinstance(item, dict) or type(item.get("id")) is not int for item in source_items)
                or {item["id"] for item in source_items} != set(ids)):
            raise CatalogEditError("SOURCE_NOT_VERIFIED", "源商品目录读取不完整或 ID 不一致，未创建克隆任务。", 409)
        if any(item.get("type") not in {"simple", "variable", "external", "grouped"} for item in source_items):
            raise CatalogEditError("SOURCE_NOT_VERIFIED", "仅可克隆明确选择的父商品，不能将变体作为父商品。", 409)
        options["_catalog_source_identities"] = {str(item["id"]): catalog_item_identity(item) for item in source_items}
        options["_catalog_source_url"] = source["url"].rstrip("/")
        options["_catalog_target_url"] = target["url"].rstrip("/")
        # Recheck live permission and exact destination after the remote read.
        for selected in (source, target):
            latest = load_site(selected["id"])
            if any(latest[field] != selected[field] for field in ("url", "consumer_key", "consumer_secret")):
                raise CatalogEditError("SITE_CHANGED", "站点配置已变化，未创建克隆任务。", 409)
        connection = get_db_connection()
        try:
            job = enqueue_clone_job(connection, source_site_id=source_id, target_site_id=target_id,
                product_ids=ids, options=options, target_url=target["url"].rstrip("/"),
                created_by_id=str(current_user.id), created_by_name=current_user.name or current_user.username)
        finally:
            connection.close()
        return jsonify(job_id=job["id"], status=job["status"], total_count=job["total_count"],
                       target_url=_public_site_url(target["url"])), 202

    return blueprint
