"""Flask API/UI for order image notifications."""

from __future__ import annotations

import hashlib
import hmac
import base64
import json
import os
import re
import tempfile
import time
import uuid
from datetime import datetime, timezone
from functools import wraps
from pathlib import Path

from flask import Blueprint, abort, jsonify, render_template, request, send_file
from flask_login import current_user, login_required

from fulfillment_common import get_conn, json_dump, utcnow
from order_notification_service import (
    ACTIVE_STATUSES,
    AUTOMATIC_EVENT_TYPES,
    DEFAULT_POLICY,
    EVENT_TYPES,
    NotificationPermanent,
    authoritative_snapshot,
    create_job_for_order,
    create_test_send_job,
    flag_enabled,
    is_fallback_target,
    load_policy,
    notification_schema_exists,
    notification_summary,
    resolve_fallback_target,
    target_manager_names,
)
from order_notification_renderer import TEMPLATE_VERSION, render_order_cards
from order_notification_email import EmailRenderError, render_logged_admin_email
from order_notification_provider import (
    ProviderError,
    encrypt_managed_webhook,
    managed_webhook_ready,
    provider_for,
    resolve_target_webhook,
)


order_notification_bp = Blueprint("order_notification", __name__)
ALLOWED_SKEW_SECONDS = 300
MAX_EVENT_BYTES = 64 * 1024
CONFIGURABLE_STATUS_GROUPS = (
    "ready_statuses",
    "cancel_statuses",
    "hold_statuses",
    "excluded_statuses",
)
CONFIGURABLE_UPDATED_FIELDS = frozenset(DEFAULT_POLICY["updated_fields"])


def _is_super_admin() -> bool:
    """Match the order system's built-in super-admin boundary exactly."""
    return bool(
        current_user.is_authenticated
        and getattr(current_user, "username", None) == "admin"
    )


def notification_super_admin_required(function):
    @wraps(function)
    @login_required
    def wrapped(*args, **kwargs):
        if not _is_super_admin():
            return jsonify({"error": "订单群通知仅超级管理员可用"}), 403
        return function(*args, **kwargs)

    return wrapped


def _view_sources():
    from app import get_user_allowed_sources

    return get_user_allowed_sources(
        current_user.id,
        current_user.is_admin(),
        current_user.is_viewer(),
    )


def _can_view_order(conn, order_id: str) -> bool:
    row = conn.execute("SELECT source FROM orders WHERE id=?", (order_id,)).fetchone()
    allowed = _view_sources()
    return bool(row and (allowed is None or row["source"] in allowed))


def _can_edit_order(conn, order_id: str) -> bool:
    from app import get_user_editable_sources

    row = conn.execute("SELECT source FROM orders WHERE id=?", (order_id,)).fetchone()
    allowed = get_user_editable_sources(current_user)
    return bool(row and (allowed is None or row["source"] in allowed))


def _event_secret() -> bytes:
    name = os.environ.get("ORDER_NOTIFICATION_EVENT_SECRET_REF", "ORDER_NOTIFICATION_EVENT_SECRET")
    if not name or not name.replace("_", "").isalnum() or name.upper() != name:
        return b""
    return os.environ.get(name, "").encode("utf-8")


def _require_ajax():
    if request.headers.get("X-Requested-With") != "XMLHttpRequest":
        return jsonify({"error": "csrf_check_failed"}), 403
    return None


def _setting(conn, key: str, default: str = "") -> str:
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return str(row[0]) if row and row[0] is not None else default


def _target_dict(row) -> dict:
    target = dict(row)
    target["copy_to_fallback"] = bool(target.get("copy_to_fallback"))
    manager_names = list(target_manager_names(target))
    target.pop("manager_names_json", None)
    target["manager_names"] = manager_names
    secret_ref = str(target.pop("secret_ref", "") or "")
    secret_ciphertext = str(target.pop("secret_ciphertext", "") or "")
    secret_name = secret_ref[4:] if secret_ref.startswith("env:") else ""
    target["secret_ref_name"] = secret_name
    target["secret_source"] = (
        "managed" if secret_ciphertext else "environment" if secret_name else "none"
    )
    target["secret_configured"] = bool(secret_ciphertext or secret_name)
    target["secret_available"] = False
    target["secret_error"] = None
    if target["secret_configured"]:
        try:
            resolve_target_webhook(
                {
                    "secret_ciphertext": secret_ciphertext,
                    "secret_ref": secret_ref,
                }
            )
            target["secret_available"] = True
        except ProviderError as exc:
            target["secret_error"] = exc.code
    return target


def _configuration_snapshot(conn) -> dict:
    policy = load_policy(conn)
    targets = [
        _target_dict(row)
        for row in conn.execute(
            """SELECT id,name,channel_type,secret_ref,secret_ciphertext,
                      webhook_fingerprint,store_id,country_code,manager_scope,manager_names_json,warehouse_id,
                      shipping_method,environment,enabled,rate_limit_per_minute,
                      copy_to_fallback,
                      created_at,updated_at
               FROM notification_targets
               WHERE deleted_at IS NULL
               ORDER BY environment,name,id"""
        ).fetchall()
    ]
    sites = [
        dict(row)
        for row in conn.execute(
            "SELECT id,url,manager,country FROM sites ORDER BY country,url"
        ).fetchall()
    ]
    manager_site_counts = {}
    for site in sites:
        manager = str(site.get("manager") or "").strip()
        if manager:
            manager_site_counts[manager] = manager_site_counts.get(manager, 0) + 1
    managers = [
        {"name": name, "site_count": site_count}
        for name, site_count in sorted(
            manager_site_counts.items(), key=lambda item: item[0].casefold()
        )
    ]
    country_site_counts = {}
    for site in sites:
        country = str(site.get("country") or "").strip().upper()
        if country:
            country_site_counts[country] = country_site_counts.get(country, 0) + 1
    countries = [
        {"code": code, "site_count": site_count}
        for code, site_count in sorted(country_site_counts.items())
    ]
    warehouses = [
        dict(row)
        for row in conn.execute(
            "SELECT id,name,country FROM warehouses ORDER BY country,name,id"
        ).fetchall()
    ]
    recent_orders = _search_preview_orders(conn, status_filter="new")
    enabled = flag_enabled(conn, "order_notification_enabled")
    production_send = flag_enabled(conn, "order_notification_send_enabled")
    test_send = flag_enabled(conn, "order_notification_test_send_enabled")
    return {
        "enabled": enabled,
        "mode": "off" if not enabled else "preview",
        "production_send_enabled": production_send,
        "test_send_enabled": test_send,
        "managed_webhook_ready": managed_webhook_ready(),
        "sending_locked": True,
        "debounce_seconds": int(_setting(conn, "order_notification_debounce_seconds", "45") or 45),
        "retention_days": int(_setting(conn, "order_notification_image_retention_days", "30") or 30),
        "render_source": _setting(conn, "order_notification_render_source", "email"),
        "template_version": _setting(
            conn, "order_notification_template_version", TEMPLATE_VERSION
        ),
        "policy": policy,
        "targets": targets,
        "sites": sites,
        "countries": countries,
        "managers": managers,
        "warehouses": warehouses,
        "recent_orders": recent_orders,
        "preview_default_status": "new",
        "preview_order_statuses": _preview_order_statuses(conn),
        "enabled_target_count": sum(bool(item["enabled"]) for item in targets),
    }


def _like_pattern(value: str) -> str:
    """Escape a user search term for a parameterized SQLite LIKE expression."""
    escaped = value.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return f"%{escaped}%"


def _preview_order_statuses(conn) -> list[str]:
    rows = conn.execute(
        """SELECT DISTINCT LOWER(status) AS status
             FROM orders
            WHERE status IS NOT NULL AND TRIM(status)<>''"""
    ).fetchall()
    priority = {
        "processing": 0,
        "on-hold": 1,
        "offline": 2,
        "pending": 3,
        "failed": 4,
        "completed": 5,
        "shipped": 6,
        "delivered": 7,
        "cancelled": 8,
        "refunded": 9,
    }
    statuses = {str(row["status"] or "").strip() for row in rows}
    return sorted(statuses, key=lambda value: (priority.get(value, 100), value))


def _normal_preview_status_filter(value: str) -> str:
    status_filter = str(value or "new").strip().lower()
    if status_filter in {"new", "all"}:
        return status_filter
    if not re.fullmatch(r"[a-z0-9][a-z0-9_-]{0,39}", status_filter):
        raise ValueError("order_search_status_invalid")
    return status_filter


def _smart_new_order_condition(conn) -> tuple[str, list[object]]:
    """Select business-new orders using the order list's payment-aware labels."""
    policy = load_policy(conn)
    ready_statuses = [str(item).strip().lower() for item in policy.get("ready_statuses") or []]
    clauses = []
    params: list[object] = []
    if ready_statuses:
        clauses.append(
            "LOWER(COALESCE(orders.status,'')) IN ("
            + ",".join("?" for _ in ready_statuses)
            + ")"
        )
        params.extend(ready_statuses)
    # The main order list defines COD pending as "待处理", while online pending
    # means "待支付". Conversely, COD on-hold is already "已发货" and BACS
    # on-hold is "待转账确认", so no on-hold order belongs in the new-order view.
    clauses.append(
        """(LOWER(COALESCE(orders.status,''))='pending'
              AND LOWER(COALESCE(NULLIF(orders.payment_method,''),'cod'))='cod')"""
    )
    return "(" + " OR ".join(clauses or ["0"]) + ")", params


def _search_preview_orders(
    conn,
    *,
    site: str = "",
    query: str = "",
    status_filter: str = "new",
    limit: int = 50,
) -> list[dict]:
    conditions = []
    params: list[object] = []
    status_filter = _normal_preview_status_filter(status_filter)
    allowed = _view_sources()
    if allowed is not None:
        allowed = list(allowed)
        if not allowed:
            return []
        conditions.append("source IN (" + ",".join("?" for _ in allowed) + ")")
        params.extend(allowed)
    if site:
        conditions.append("source=?")
        params.append(site)
    if status_filter == "new":
        status_condition, status_params = _smart_new_order_condition(conn)
        conditions.append(status_condition)
        params.extend(status_params)
    elif status_filter != "all":
        conditions.append("LOWER(COALESCE(orders.status,''))=?")
        params.append(status_filter)
    if query:
        pattern = _like_pattern(query)
        conditions.append(
            "(id LIKE ? ESCAPE '\\' OR CAST(number AS TEXT) LIKE ? ESCAPE '\\')"
        )
        params.extend((pattern, pattern))
    where = "WHERE " + " AND ".join(conditions) if conditions else ""
    params.extend((query or "", query or "", int(limit)))
    rows = conn.execute(
        f"""SELECT id,number,status,payment_method,source,warehouse_id,date_modified
              FROM orders
              {where}
             ORDER BY CASE
                        WHEN id=? THEN 0
                        WHEN CAST(number AS TEXT)=? THEN 1
                        ELSE 2
                      END,
                      COALESCE(date_modified,updated_at) DESC
             LIMIT ?""",
        params,
    ).fetchall()
    return [dict(row) for row in rows]


def _status_list(value, key: str) -> list[str]:
    if not isinstance(value, list) or len(value) > 20:
        raise ValueError(f"{key}_invalid")
    cleaned = []
    for item in value:
        status = str(item).strip().lower()
        if not re.fullmatch(r"[a-z0-9][a-z0-9_-]{0,39}", status):
            raise ValueError(f"{key}_invalid")
        if status not in cleaned:
            cleaned.append(status)
    return cleaned


def verify_event_request(raw: bytes, timestamp: str | None, signature: str | None, *, now: int | None = None) -> tuple[bool, str]:
    secret = _event_secret()
    if not secret:
        return False, "event_secret_missing"
    try:
        stamp = int(timestamp or "")
    except ValueError:
        return False, "timestamp_invalid"
    current = int(time.time() if now is None else now)
    if abs(current - stamp) > ALLOWED_SKEW_SECONDS:
        return False, "timestamp_expired"
    supplied = str(signature or "")
    if supplied.startswith("sha256="):
        supplied = supplied[7:]
    expected = hmac.new(secret, str(stamp).encode("ascii") + b"." + raw, hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected, supplied):
        return False, "signature_invalid"
    return True, "ok"


@order_notification_bp.route("/api/v1/order-notification-events", methods=["POST"])
def receive_order_event():
    if request.content_length is not None and request.content_length > MAX_EVENT_BYTES:
        return jsonify({"accepted": False, "error": "payload_too_large"}), 413
    raw = request.get_data(cache=True)
    if len(raw) > MAX_EVENT_BYTES:
        return jsonify({"accepted": False, "error": "payload_too_large"}), 413
    valid, reason = verify_event_request(
        raw,
        request.headers.get("X-Webhook-Timestamp"),
        request.headers.get("X-Webhook-Signature"),
    )
    if not valid:
        status = 503 if reason == "event_secret_missing" else 401
        return jsonify({"accepted": False, "error": reason}), status
    try:
        data = json.loads(raw)
    except (TypeError, ValueError):
        return jsonify({"accepted": False, "error": "invalid_json"}), 400
    if not isinstance(data, dict):
        return jsonify({"accepted": False, "error": "payload_invalid"}), 400
    required = {"event_id", "event_type", "occurred_at", "store_id", "order_id", "source"}
    if not required.issubset(data):
        return jsonify({"accepted": False, "error": "missing_fields"}), 400
    event_id = str(data["event_id"])
    order_id = str(data["order_id"])
    store_id = str(data["store_id"])
    if not re.fullmatch(r"[A-Za-z0-9._:-]{1,128}", event_id):
        return jsonify({"accepted": False, "error": "event_id_invalid"}), 400
    if not re.fullmatch(r"[A-Za-z0-9._:-]{1,128}", order_id):
        return jsonify({"accepted": False, "error": "order_id_invalid"}), 400
    if not (1 <= len(store_id) <= 300):
        return jsonify({"accepted": False, "error": "store_id_invalid"}), 400
    if data["event_type"] not in EVENT_TYPES - {"MANUAL_RESEND"}:
        return jsonify({"accepted": False, "error": "event_type_invalid"}), 400
    if data["event_type"] not in AUTOMATIC_EVENT_TYPES:
        return jsonify({"accepted": False, "error": "event_type_not_enabled"}), 400
    if data["source"] != "hongkong-order-system":
        return jsonify({"accepted": False, "error": "source_invalid"}), 400
    try:
        occurred = datetime.fromisoformat(str(data["occurred_at"]).replace("Z", "+00:00"))
        if occurred.tzinfo is None:
            raise ValueError
        if abs((datetime.now(timezone.utc) - occurred.astimezone(timezone.utc)).total_seconds()) > 86400:
            return jsonify({"accepted": False, "error": "occurred_at_out_of_range"}), 400
    except (TypeError, ValueError):
        return jsonify({"accepted": False, "error": "occurred_at_invalid"}), 400

    payload_hash = hashlib.sha256(raw).hexdigest()
    conn = get_conn()
    try:
        if not notification_schema_exists(conn):
            return jsonify({"accepted": False, "error": "schema_missing"}), 503
        existing = conn.execute(
            "SELECT payload_hash,status FROM order_notification_event_inbox WHERE event_id=?",
            (event_id,),
        ).fetchone()
        if existing:
            if existing["payload_hash"] != payload_hash:
                return jsonify({"accepted": False, "error": "event_id_conflict"}), 409
            job = conn.execute(
                "SELECT id FROM order_notification_jobs WHERE event_id=? ORDER BY created_at LIMIT 1",
                (event_id,),
            ).fetchone()
            return jsonify({"accepted": True, "duplicate": True, "job_id": job["id"] if job else None}), 202
        conn.execute(
            """INSERT INTO order_notification_event_inbox
               (event_id,event_type,occurred_at,source,payload_hash,status)
               VALUES (?,?,?,?,?,'received')""",
            (
                event_id, data["event_type"], str(data["occurred_at"]),
                str(data["source"]), payload_hash,
            ),
        )
        snapshot_store = conn.execute(
            "SELECT source FROM orders WHERE id=?", (order_id,)
        ).fetchone()
        if not snapshot_store:
            raise NotificationPermanent("订单不存在", code="order_not_found")
        if snapshot_store["source"] != store_id:
            raise NotificationPermanent("事件店铺与权威订单不一致", code="store_mismatch")
        result = create_job_for_order(
            conn,
            order_id,
            event_id=event_id,
            allowed_event_types=AUTOMATIC_EVENT_TYPES,
            target_environment="production",
        )
        conn.execute(
            """UPDATE order_notification_event_inbox
               SET status=?,processed_at=? WHERE event_id=?""",
            ("queued" if result.get("created") else result.get("reason", "ignored"), utcnow(), event_id),
        )
        conn.commit()
        job_data = result.get("job") or {}
        return jsonify(
            {
                "accepted": True,
                "duplicate": bool(result.get("duplicate")),
                "job_id": job_data.get("id"),
                "queued": bool(result.get("created")),
                "reason": result.get("reason"),
            }
        ), 202
    except NotificationPermanent as exc:
        conn.execute(
            "UPDATE order_notification_event_inbox SET status='rejected',error_summary=?,processed_at=? WHERE event_id=?",
            (exc.code, utcnow(), str(data.get("event_id"))),
        )
        conn.commit()
        return jsonify({"accepted": False, "error": exc.code}), 400
    finally:
        conn.close()


@order_notification_bp.route("/api/order/<path:order_id>/notifications")
@notification_super_admin_required
def order_notifications(order_id):
    conn = get_conn()
    try:
        if not _can_view_order(conn, order_id):
            return jsonify({"error": "无权查看该站点订单"}), 403
        return jsonify(notification_summary(conn, order_id))
    finally:
        conn.close()


@order_notification_bp.route("/api/order/<path:order_id>/notifications/<job_id>/resend", methods=["POST"])
@notification_super_admin_required
def resend_notification(order_id, job_id):
    csrf = _require_ajax()
    if csrf:
        return csrf
    conn = get_conn()
    try:
        if not _can_edit_order(conn, order_id):
            return jsonify({"error": "无权重发该站点订单"}), 403
        result = create_job_for_order(
            conn,
            order_id,
            event_id="manual:" + uuid.uuid4().hex,
            resend_of=job_id,
            actor={"type": "user", "id": str(current_user.id)},
            target_environment="production",
        )
        return jsonify({"created": bool(result.get("created")), "job_id": (result.get("job") or {}).get("id")}), 202
    except NotificationPermanent as exc:
        return jsonify({"error": exc.code}), 400
    finally:
        conn.close()


@order_notification_bp.route("/api/order-notifications/config", methods=["GET", "POST"])
@notification_super_admin_required
def notification_configuration():
    conn = get_conn()
    try:
        if not notification_schema_exists(conn):
            return jsonify({"error": "schema_missing"}), 503
        if request.method == "GET":
            return jsonify(_configuration_snapshot(conn))
        csrf = _require_ajax()
        if csrf:
            return csrf
        data = request.get_json(silent=True)
        if not isinstance(data, dict):
            return jsonify({"error": "invalid_configuration"}), 400
        if {
            "production_send_enabled",
            "test_send_enabled",
            "order_notification_send_enabled",
            "order_notification_test_send_enabled",
        }.intersection(data):
            return jsonify({"error": "sending_locked"}), 403
        if not isinstance(data.get("enabled"), bool):
            return jsonify({"error": "enabled_invalid"}), 400
        debounce = int(data.get("debounce_seconds", 45))
        retention = int(data.get("retention_days", 30))
        render_source = str(data.get("render_source") or _setting(
            conn, "order_notification_render_source", "email"
        )).strip().lower()
        if not 0 <= debounce <= 600:
            return jsonify({"error": "debounce_out_of_range"}), 400
        if not 1 <= retention <= 365:
            return jsonify({"error": "retention_out_of_range"}), 400
        if render_source not in {"email", "system_card"}:
            return jsonify({"error": "render_source_invalid"}), 400
        policy_data = data.get("policy")
        if not isinstance(policy_data, dict):
            return jsonify({"error": "policy_invalid"}), 400
        policy = {
            key: _status_list(policy_data.get(key), key)
            for key in CONFIGURABLE_STATUS_GROUPS
        }
        occupied: dict[str, str] = {}
        for group in CONFIGURABLE_STATUS_GROUPS:
            for status in policy[group]:
                previous = occupied.setdefault(status, group)
                if previous != group:
                    return jsonify({"error": "status_groups_overlap", "status": status}), 400
        updated_fields = policy_data.get("updated_fields")
        if not isinstance(updated_fields, list) or not updated_fields:
            return jsonify({"error": "updated_fields_invalid"}), 400
        policy["updated_fields"] = []
        for field in updated_fields:
            name = str(field)
            if name not in CONFIGURABLE_UPDATED_FIELDS:
                return jsonify({"error": "updated_fields_invalid"}), 400
            if name not in policy["updated_fields"]:
                policy["updated_fields"].append(name)
        if not isinstance(policy_data.get("allow_site_cod_on_hold"), bool):
            return jsonify({"error": "allow_site_cod_on_hold_invalid"}), 400
        policy["allow_site_cod_on_hold"] = policy_data["allow_site_cod_on_hold"]

        before = _configuration_snapshot(conn)
        updates = {
            "order_notification_enabled": "1" if data["enabled"] else "0",
            "order_notification_debounce_seconds": str(debounce),
            "order_notification_image_retention_days": str(retention),
            "order_notification_render_source": render_source,
            "order_notification_policy_json": json_dump(policy),
        }
        for key, value in updates.items():
            conn.execute(
                """INSERT INTO settings (key,value) VALUES (?,?)
                   ON CONFLICT(key) DO UPDATE SET value=excluded.value""",
                (key, value),
            )
        after = _configuration_snapshot(conn)
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,actor_id,before_summary,after_summary)
               VALUES ('configuration_updated','notification_configuration','global','user',?,?,?)""",
            (
                str(current_user.id),
                json_dump(
                    {
                        "enabled": before["enabled"],
                        "debounce_seconds": before["debounce_seconds"],
                        "retention_days": before["retention_days"],
                        "render_source": before["render_source"],
                        "policy": before["policy"],
                    }
                ),
                json_dump(
                    {
                        "enabled": after["enabled"],
                        "debounce_seconds": after["debounce_seconds"],
                        "retention_days": after["retention_days"],
                        "render_source": after["render_source"],
                        "policy": after["policy"],
                    }
                ),
            ),
        )
        conn.commit()
        return jsonify(after)
    except (TypeError, ValueError) as exc:
        conn.rollback()
        message = str(exc)
        return jsonify({"error": message if message else "invalid_configuration"}), 400
    finally:
        conn.close()


@order_notification_bp.route("/api/order-notifications/preview", methods=["POST"])
@notification_super_admin_required
def notification_preview():
    csrf = _require_ajax()
    if csrf:
        return csrf
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return jsonify({"error": "payload_invalid"}), 400
    order_id = str(data.get("order_id") or "").strip()
    event_type = str(data.get("event_type") or "ORDER_READY")
    source = str(data.get("source") or "email").strip().lower()
    if not order_id or len(order_id) > 128:
        return jsonify({"error": "order_id_invalid"}), 400
    if event_type not in EVENT_TYPES - {"MANUAL_RESEND"}:
        return jsonify({"error": "event_type_invalid"}), 400
    if source not in {"email", "system_card"}:
        return jsonify({"error": "preview_source_invalid"}), 400
    if source == "email" and event_type != "ORDER_READY":
        return jsonify({"error": "email_source_only_supports_new_order"}), 400
    conn = get_conn()
    try:
        if not _can_view_order(conn, order_id):
            return jsonify({"error": "无权查看该站点订单"}), 403
        snapshot = authoritative_snapshot(conn, order_id)
        snapshot["notification_at"] = utcnow()
        snapshot["changes"] = []
        root = Path(
            os.environ.get("ORDER_NOTIFICATION_IMAGE_DIR")
            or (Path("var") / "order-cards")
        ).resolve()
        root.mkdir(parents=True, exist_ok=True)
        preview_id = "preview-" + uuid.uuid4().hex
        images = []
        source_metadata = None
        with tempfile.TemporaryDirectory(prefix=preview_id + "-", dir=root) as temp_dir:
            if source == "email":
                rendered, source_metadata = render_logged_admin_email(
                    conn, order_id, temp_dir, preview_id
                )
            else:
                rendered = render_order_cards(
                    snapshot,
                    event_type,
                    temp_dir,
                    preview_id,
                    template_version=_setting(
                        conn, "order_notification_template_version", TEMPLATE_VERSION
                    ),
                )
            if len(rendered) > 5:
                return jsonify({"error": "preview_too_many_pages"}), 413
            for item in rendered:
                raw = Path(item["path"]).read_bytes()
                images.append(
                    {
                        "page": item["page"],
                        "bytes": len(raw),
                        "data_url": "data:image/png;base64,"
                        + base64.b64encode(raw).decode("ascii"),
                    }
                )
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,actor_id,request_id,after_summary)
               VALUES ('preview_generated','order',?,'user',?,?,?)""",
            (
                order_id,
                str(current_user.id),
                preview_id,
                json_dump(
                    {
                        "event_type": event_type,
                        "preview_source": source,
                        "pages": len(images),
                        "email_log_id": source_metadata.get("log_id") if source_metadata else None,
                        "html_sha256": source_metadata.get("html_sha256") if source_metadata else None,
                    }
                ),
            ),
        )
        conn.commit()
        return jsonify(
            {
                "preview_id": preview_id,
                "order": {
                    "id": snapshot["order_id"],
                    "number": snapshot["number"],
                    "store": snapshot["store_label"],
                    "status": snapshot["status"],
                },
                "routing": {
                    "store_id": snapshot["store_id"],
                    "manager_name": snapshot.get("site_manager"),
                    "warehouse_id": snapshot["warehouse_id"],
                    "shipping_method": snapshot["shipping_method"],
                },
                "event_type": event_type,
                "source": source,
                "email": (
                    {
                        "log_id": source_metadata.get("log_id"),
                        "plugin": source_metadata.get("plugin"),
                        "source": source_metadata.get("source"),
                        "subject": source_metadata.get("subject"),
                        "sent_at": source_metadata.get("sent_at"),
                        "images_inlined": source_metadata.get("images_inlined", 0),
                        "images_removed": source_metadata.get("images_removed", 0),
                    }
                    if source_metadata
                    else None
                ),
                "images": images,
                "queued": False,
                "sent": False,
            }
        )
    except EmailRenderError as exc:
        status = 404 if exc.code in {
            "admin_new_order_email_not_found", "email_body_missing", "order_not_found"
        } else 503 if exc.code in {"playwright_missing", "chromium_missing", "beautifulsoup_missing"} else 502
        return jsonify({"error": exc.code, "message": str(exc), "retryable": exc.retryable}), status
    except NotificationPermanent as exc:
        return jsonify({"error": exc.code}), 404
    finally:
        conn.close()


@order_notification_bp.route("/api/order-notifications/test-send", methods=["POST"])
@notification_super_admin_required
def notification_test_send():
    """Queue the currently reviewed order for one isolated-group test send."""
    csrf = _require_ajax()
    if csrf:
        return csrf
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return jsonify({"error": "payload_invalid"}), 400
    order_id = str(data.get("order_id") or "").strip()
    target_id = str(data.get("target_id") or "").strip()
    preview_id = str(data.get("preview_id") or "").strip()
    source = str(data.get("source") or "email").strip().lower()
    if data.get("confirmed") is not True:
        return jsonify({"error": "test_send_confirmation_required"}), 400
    if not order_id or len(order_id) > 128:
        return jsonify({"error": "order_id_invalid"}), 400
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", target_id):
        return jsonify({"error": "test_target_invalid"}), 400
    if not re.fullmatch(r"preview-[a-f0-9]{32}", preview_id):
        return jsonify({"error": "preview_id_invalid"}), 400
    if source not in {"email", "system_card"}:
        return jsonify({"error": "render_source_invalid"}), 400

    conn = get_conn()
    try:
        if not _can_edit_order(conn, order_id):
            return jsonify({"error": "无权发送该站点订单"}), 403
        preview_row = conn.execute(
            """SELECT after_summary FROM notification_audit_logs
                WHERE action='preview_generated' AND object_type='order'
                  AND object_id=? AND actor_type='user' AND actor_id=?
                  AND request_id=? AND created_at>=datetime('now','-15 minutes')
                ORDER BY id DESC LIMIT 1""",
            (order_id, str(current_user.id), preview_id),
        ).fetchone()
        try:
            preview_summary = json.loads(preview_row["after_summary"])
        except (TypeError, ValueError):
            preview_summary = {}
        if not preview_row or preview_summary.get("preview_source") != source:
            return jsonify({"error": "preview_id_invalid"}), 409
        if not flag_enabled(conn, "order_notification_test_send_enabled"):
            return jsonify({"error": "test_send_flag_off"}), 409
        target_row = conn.execute(
            "SELECT * FROM notification_targets WHERE id=? AND enabled=1", (target_id,)
        ).fetchone()
        if not target_row:
            return jsonify({"error": "test_target_missing"}), 404
        target = dict(target_row)
        if target.get("environment") != "test":
            return jsonify({"error": "test_target_required"}), 400
        if target.get("channel_type") != "WECOM_BOT":
            return jsonify({"error": "test_target_channel_invalid"}), 400
        try:
            resolve_target_webhook(target)
        except ProviderError as exc:
            return jsonify({"error": exc.code}), 409

        result = create_test_send_job(
            conn,
            order_id,
            target_id=target_id,
            preview_id=preview_id,
            render_source=source,
            actor={"type": "user", "id": str(current_user.id)},
        )
        if result.get("reason") == "schema_missing":
            return jsonify({"error": "schema_missing"}), 503
        if result.get("reason") == "not_eligible":
            return jsonify({"error": "order_not_ready_for_test"}), 409
        job = result.get("job") or {}
        return jsonify(
            {
                "queued": bool(result.get("created")),
                "duplicate": bool(result.get("duplicate")),
                "sent": False,
                "job_id": job.get("id"),
                "status": job.get("status"),
                "target": {"id": target["id"], "name": target["name"]},
            }
        ), 202 if result.get("created") else 200
    except NotificationPermanent as exc:
        return jsonify({"error": exc.code}), 400
    finally:
        conn.close()


@order_notification_bp.route("/api/order-notifications/orders", methods=["GET"])
@notification_super_admin_required
def notification_order_search():
    site = str(request.args.get("site") or "").strip()
    query = str(request.args.get("q") or "").strip()
    try:
        status_filter = _normal_preview_status_filter(request.args.get("status") or "new")
    except ValueError:
        return jsonify({"error": "order_search_status_invalid"}), 400
    if query.startswith("#"):
        query = query[1:].strip()
    if len(query) > 128:
        return jsonify({"error": "order_search_query_too_long"}), 400
    if len(site) > 2048:
        return jsonify({"error": "order_search_site_invalid"}), 400
    conn = get_conn()
    try:
        if site:
            site_row = conn.execute("SELECT 1 FROM sites WHERE url=?", (site,)).fetchone()
            allowed = _view_sources()
            if not site_row or (allowed is not None and site not in allowed):
                return jsonify({"error": "order_search_site_invalid"}), 400
        orders = _search_preview_orders(
            conn, site=site, query=query, status_filter=status_filter
        )
        return jsonify(
            {
                "orders": orders,
                "site": site,
                "query": query,
                "status_filter": status_filter,
            }
        )
    finally:
        conn.close()


@order_notification_bp.route("/api/order-notifications/targets", methods=["GET", "POST"])
@notification_super_admin_required
def notification_targets():
    conn = get_conn()
    try:
        if request.method == "GET":
            rows = conn.execute(
                """SELECT id,name,channel_type,secret_ref,secret_ciphertext,
                          webhook_fingerprint,store_id,country_code,manager_scope,manager_names_json,
                          warehouse_id,shipping_method,
                          environment,enabled,rate_limit_per_minute,
                          copy_to_fallback,
                          created_at,updated_at
                   FROM notification_targets
                   WHERE deleted_at IS NULL
                   ORDER BY environment,name,id"""
            ).fetchall()
            return jsonify([_target_dict(row) for row in rows])
        csrf = _require_ajax()
        if csrf:
            return csrf
        data = request.get_json(silent=True) or {}
        channel = str(data.get("channel_type") or "")
        environment = str(data.get("environment") or "test")
        target_id = str(data.get("id") or uuid.uuid4().hex)
        if not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", target_id):
            return jsonify({"error": "target_id_invalid"}), 400
        existing = conn.execute(
            "SELECT * FROM notification_targets WHERE id=?", (target_id,)
        ).fetchone()
        if existing and existing["deleted_at"]:
            return jsonify({"error": "target_deleted"}), 410
        webhook_url = str(data.get("webhook_url") or "").strip()
        if len(webhook_url) > 2048:
            return jsonify({"error": "webhook_invalid"}), 400
        supplied_secret_ref = str(data.get("secret_ref") or "") or None
        secret_ref = existing["secret_ref"] if existing else None
        secret_ciphertext = existing["secret_ciphertext"] if existing else None
        fingerprint = existing["webhook_fingerprint"] if existing else None
        if channel not in {"WECOM_BOT", "MANUAL_WECHAT", "FAKE"}:
            return jsonify({"error": "channel_invalid"}), 400
        if environment not in {"test", "production"}:
            return jsonify({"error": "environment_invalid"}), 400
        if channel == "WECOM_BOT":
            if webhook_url:
                secret_ciphertext, fingerprint = encrypt_managed_webhook(webhook_url)
                secret_ref = None
            elif supplied_secret_ref:
                if not re.fullmatch(r"env:[A-Z][A-Z0-9_]{2,100}", supplied_secret_ref):
                    return jsonify({"error": "secret_ref_invalid"}), 400
                ref_name = supplied_secret_ref[4:]
                if environment == "test" and "TEST" not in ref_name:
                    return jsonify({"error": "test_secret_ref_must_be_isolated"}), 400
                if environment == "production" and "TEST" in ref_name:
                    return jsonify({"error": "production_cannot_use_test_secret"}), 400
                secret_ref = supplied_secret_ref
                secret_ciphertext = None
                fingerprint = None
            if not secret_ciphertext and not secret_ref:
                return jsonify({"error": "webhook_required"}), 400
        else:
            secret_ref = None
            secret_ciphertext = None
            fingerprint = None
        name = str(data.get("name") or "").strip()[:120]
        if not name:
            return jsonify({"error": "name_required"}), 400
        store_id = str(data.get("store_id") or "").strip() or None
        site_row = None
        if store_id:
            site_row = conn.execute(
                "SELECT manager,country FROM sites WHERE url=?", (store_id,)
            ).fetchone()
            if not site_row:
                return jsonify({"error": "store_invalid"}), 400
        country_code = str(data.get("country_code") or "").strip().upper() or None
        available_countries = {
            str(row[0]).strip().upper()
            for row in conn.execute(
                "SELECT DISTINCT country FROM sites WHERE TRIM(COALESCE(country,''))<>''"
            ).fetchall()
        }
        if country_code and country_code not in available_countries:
            return jsonify({"error": "country_invalid"}), 400
        if (
            site_row
            and country_code
            and str(site_row["country"] or "").strip().upper() != country_code
        ):
            return jsonify({"error": "store_country_mismatch"}), 400
        manager_scope = str(
            data.get("manager_scope")
            or (existing["manager_scope"] if existing else "all")
        ).strip().lower()
        if manager_scope not in {"all", "selected"}:
            return jsonify({"error": "manager_scope_invalid"}), 400
        raw_manager_names = data.get("manager_names")
        if raw_manager_names is None:
            manager_names = (
                list(target_manager_names(dict(existing))) if existing else []
            )
        elif isinstance(raw_manager_names, list):
            manager_names = sorted(
                {
                    str(value).strip()
                    for value in raw_manager_names
                    if isinstance(value, str) and str(value).strip()
                },
                key=str.casefold,
            )
        else:
            return jsonify({"error": "manager_names_invalid"}), 400
        if manager_scope == "all":
            manager_names = []
        elif not manager_names:
            return jsonify({"error": "manager_selection_required"}), 400
        if len(manager_names) > 100 or any(len(value) > 120 for value in manager_names):
            return jsonify({"error": "manager_names_invalid"}), 400
        available_managers = {
            str(row[0]).strip()
            for row in conn.execute(
                "SELECT DISTINCT manager FROM sites WHERE TRIM(COALESCE(manager,''))<>''"
            ).fetchall()
        }
        if any(value not in available_managers for value in manager_names):
            return jsonify({"error": "manager_invalid"}), 400
        if (
            site_row
            and manager_scope == "selected"
            and str(site_row["manager"] or "").strip() not in manager_names
        ):
            return jsonify({"error": "store_manager_mismatch"}), 400
        warehouse_id = data.get("warehouse_id")
        if warehouse_id in (None, ""):
            warehouse_id = None
        else:
            warehouse_id = int(warehouse_id)
            if not conn.execute(
                "SELECT 1 FROM warehouses WHERE id=?", (warehouse_id,)
            ).fetchone():
                return jsonify({"error": "warehouse_invalid"}), 400
        shipping_method = str(data.get("shipping_method") or "").strip()[:160] or None
        enabled = data.get("enabled", True)
        if not isinstance(enabled, bool):
            return jsonify({"error": "enabled_invalid"}), 400
        copy_to_fallback = data.get(
            "copy_to_fallback",
            bool(existing["copy_to_fallback"]) if existing else False,
        )
        if not isinstance(copy_to_fallback, bool):
            return jsonify({"error": "copy_to_fallback_invalid"}), 400
        candidate = {
            "id": target_id,
            "store_id": store_id,
            "country_code": country_code,
            "manager_scope": manager_scope,
            "manager_names_json": json.dumps(
                manager_names, ensure_ascii=False, separators=(",", ":")
            ),
            "warehouse_id": warehouse_id,
            "shipping_method": shipping_method,
            "environment": environment,
            "enabled": 1 if enabled else 0,
            "copy_to_fallback": 1 if copy_to_fallback else 0,
        }
        if copy_to_fallback and is_fallback_target(candidate):
            return jsonify({"error": "fallback_cannot_copy_to_itself"}), 400
        if enabled and copy_to_fallback:
            _, fallback_error = resolve_fallback_target(
                conn,
                environment=environment,
                exclude_target_id=target_id,
            )
            if fallback_error:
                return jsonify({"error": fallback_error}), 409
        if existing and existing["enabled"] and is_fallback_target(dict(existing)):
            remains_same_fallback = bool(
                enabled
                and environment == existing["environment"]
                and is_fallback_target(candidate)
            )
            if not remains_same_fallback:
                fallback_users = conn.execute(
                    """SELECT COUNT(*) FROM notification_targets
                         WHERE id<>? AND enabled=1 AND deleted_at IS NULL
                           AND environment=? AND copy_to_fallback=1""",
                    (target_id, existing["environment"]),
                ).fetchone()[0]
                if fallback_users:
                    _, alternate_error = resolve_fallback_target(
                        conn,
                        environment=existing["environment"],
                        exclude_target_id=target_id,
                    )
                    if alternate_error:
                        return jsonify(
                            {
                                "error": "fallback_target_in_use",
                                "dependent_targets": fallback_users,
                            }
                        ), 409
        if enabled:
            same_route_rows = conn.execute(
                """SELECT id FROM notification_targets
                   WHERE id<>? AND enabled=1
                     AND deleted_at IS NULL
                     AND environment=?
                     AND store_id IS ?
                     AND LOWER(COALESCE(country_code,''))=LOWER(COALESCE(?,''))
                     AND warehouse_id IS ?
                     AND LOWER(COALESCE(shipping_method,''))=LOWER(COALESCE(?,''))""",
                (
                    target_id, environment, store_id, country_code,
                    warehouse_id, shipping_method,
                ),
            ).fetchall()
            for route_row in same_route_rows:
                other = dict(conn.execute(
                    "SELECT * FROM notification_targets WHERE id=?",
                    (route_row["id"],),
                ).fetchone())
                other_scope = other.get("manager_scope") or "all"
                if manager_scope != other_scope:
                    continue
                if manager_scope == "all" or set(manager_names).intersection(
                    target_manager_names(other)
                ):
                    return jsonify({"error": "route_ambiguous"}), 409
        before = conn.execute(
            """SELECT id,name,channel_type,webhook_fingerprint,store_id,country_code,
                      manager_scope,manager_names_json,warehouse_id,shipping_method,
                      environment,enabled,rate_limit_per_minute,copy_to_fallback
                 FROM notification_targets WHERE id=?""",
            (target_id,),
        ).fetchone()
        conn.execute(
            """INSERT INTO notification_targets
               (id,name,channel_type,secret_ref,secret_ciphertext,webhook_fingerprint,
                store_id,country_code,manager_scope,manager_names_json,warehouse_id,shipping_method,
                environment,enabled,rate_limit_per_minute,copy_to_fallback,deleted_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,NULL)
               ON CONFLICT(id) DO UPDATE SET name=excluded.name,
                 channel_type=excluded.channel_type,secret_ref=excluded.secret_ref,
                 secret_ciphertext=excluded.secret_ciphertext,
                 webhook_fingerprint=excluded.webhook_fingerprint,
                 store_id=excluded.store_id,country_code=excluded.country_code,
                 manager_scope=excluded.manager_scope,
                 manager_names_json=excluded.manager_names_json,
                 warehouse_id=excluded.warehouse_id,
                 shipping_method=excluded.shipping_method,environment=excluded.environment,
                 enabled=excluded.enabled,rate_limit_per_minute=excluded.rate_limit_per_minute,
                 copy_to_fallback=excluded.copy_to_fallback,
                 updated_at=CURRENT_TIMESTAMP""",
            (
                target_id, name, channel, secret_ref, secret_ciphertext, fingerprint,
                store_id, country_code, manager_scope,
                json.dumps(manager_names, ensure_ascii=False, separators=(",", ":")),
                warehouse_id, shipping_method,
                environment, 1 if enabled else 0,
                max(1, min(15, int(data.get("rate_limit_per_minute") or 15))),
                1 if copy_to_fallback else 0,
            ),
        )
        after = conn.execute(
            """SELECT id,name,channel_type,webhook_fingerprint,store_id,country_code,
                      manager_scope,manager_names_json,warehouse_id,shipping_method,
                      environment,enabled,rate_limit_per_minute,copy_to_fallback
                 FROM notification_targets WHERE id=?""",
            (target_id,),
        ).fetchone()
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,actor_id,before_summary,after_summary)
               VALUES (?,?,?,?,?,?,?)""",
            (
                "target_updated" if before else "target_created",
                "notification_target", target_id, "user", str(current_user.id),
                json_dump(dict(before)) if before else None,
                json_dump(dict(after)),
            ),
        )
        conn.commit()
        return jsonify(_target_dict(conn.execute(
            "SELECT * FROM notification_targets WHERE id=?", (target_id,)
        ).fetchone())), 200 if before else 201
    except ProviderError as exc:
        conn.rollback()
        status = 503 if exc.code in {
            "webhook_master_key_missing", "webhook_master_key_invalid"
        } else 400
        return jsonify({"error": exc.code}), status
    except (TypeError, ValueError):
        conn.rollback()
        return jsonify({"error": "invalid_configuration"}), 400
    finally:
        conn.close()


@order_notification_bp.route(
    "/api/order-notifications/targets/<target_id>", methods=["DELETE"]
)
@notification_super_admin_required
def notification_target_delete(target_id):
    csrf = _require_ajax()
    if csrf:
        return csrf
    data = request.get_json(silent=True) or {}
    if data.get("confirmed") is not True:
        return jsonify({"error": "target_delete_confirmation_required"}), 400
    conn = get_conn()
    try:
        target = conn.execute(
            """SELECT * FROM notification_targets
                 WHERE id=? AND deleted_at IS NULL""",
            (target_id,),
        ).fetchone()
        if not target:
            return jsonify({"error": "target_missing"}), 404
        if target["enabled"] and is_fallback_target(dict(target)):
            fallback_users = conn.execute(
                """SELECT COUNT(*) FROM notification_targets
                     WHERE id<>? AND enabled=1 AND deleted_at IS NULL
                       AND environment=? AND copy_to_fallback=1""",
                (target_id, target["environment"]),
            ).fetchone()[0]
            if fallback_users:
                _, alternate_error = resolve_fallback_target(
                    conn,
                    environment=target["environment"],
                    exclude_target_id=target_id,
                )
                if alternate_error:
                    return jsonify(
                        {
                            "error": "fallback_target_in_use",
                            "dependent_targets": fallback_users,
                        }
                    ), 409
        placeholders = ",".join("?" for _ in ACTIVE_STATUSES)
        active = conn.execute(
            f"""SELECT COUNT(*) FROM order_notification_jobs
                  WHERE target_id=? AND status IN ({placeholders})""",
            (target_id, *sorted(ACTIVE_STATUSES)),
        ).fetchone()[0]
        if active:
            return jsonify({"error": "target_has_active_jobs", "active_jobs": active}), 409
        before = _target_dict(target)
        conn.execute(
            """UPDATE notification_targets
                  SET enabled=0,secret_ref=NULL,secret_ciphertext=NULL,
                      webhook_fingerprint=NULL,deleted_at=CURRENT_TIMESTAMP,
                      updated_at=CURRENT_TIMESTAMP
                WHERE id=?""",
            (target_id,),
        )
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,actor_id,before_summary,after_summary)
               VALUES ('target_deleted','notification_target',?,'user',?,?,?)""",
            (
                target_id,
                str(current_user.id),
                json_dump(before),
                json_dump({"deleted": True, "secret_cleared": True}),
            ),
        )
        conn.commit()
        return jsonify({"id": target_id, "deleted": True, "secret_cleared": True})
    finally:
        conn.close()


@order_notification_bp.route(
    "/api/order-notifications/targets/<target_id>/test", methods=["POST"]
)
@notification_super_admin_required
def notification_target_test_message(target_id):
    csrf = _require_ajax()
    if csrf:
        return csrf
    data = request.get_json(silent=True) or {}
    if data.get("confirmed") is not True:
        return jsonify({"error": "test_message_confirmation_required"}), 400
    conn = get_conn()
    request_id = "target-test-" + uuid.uuid4().hex
    try:
        row = conn.execute(
            """SELECT * FROM notification_targets
                 WHERE id=? AND deleted_at IS NULL""",
            (target_id,),
        ).fetchone()
        if not row:
            return jsonify({"error": "target_missing"}), 404
        target = dict(row)
        if target.get("channel_type") != "WECOM_BOT":
            return jsonify({"error": "test_target_channel_invalid"}), 400
        recent = conn.execute(
            """SELECT 1 FROM notification_audit_logs
                WHERE action='target_test_started'
                  AND object_type='notification_target' AND object_id=?
                  AND created_at>=datetime('now','-10 seconds')
                LIMIT 1""",
            (target_id,),
        ).fetchone()
        if recent:
            return jsonify({"error": "test_message_rate_limited"}), 429
        manager_names = target_manager_names(target)
        manager_label = (
            "全部负责人"
            if target.get("manager_scope") != "selected"
            else "、".join(manager_names)
        )
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,actor_id,request_id,after_summary)
               VALUES ('target_test_started','notification_target',?,'user',?,?,?)""",
            (
                target_id,
                str(current_user.id),
                request_id,
                json_dump({"target_name": target["name"], "managers": manager_label}),
            ),
        )
        conn.commit()
        content = (
            "✅ 订单系统群通知测试\n"
            f"目标：{str(target['name'])[:120]}\n"
            f"负责人：{manager_label[:500]}\n"
            "这是一条连接测试消息；收到即表示企业微信群机器人可用。\n"
            f"时间：{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}"
        )
        try:
            result = provider_for("WECOM_BOT").send_text(content, target)
        except ProviderError as exc:
            conn.execute(
                """INSERT INTO notification_audit_logs
                   (action,object_type,object_id,actor_type,actor_id,request_id,after_summary)
                   VALUES ('target_test_failed','notification_target',?,'user',?,?,?)""",
                (
                    target_id,
                    str(current_user.id),
                    request_id,
                    json_dump({"error": exc.code, "http_status": exc.http_status}),
                ),
            )
            conn.commit()
            return jsonify({"error": exc.code}), 503 if exc.retryable else 409
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,actor_id,request_id,after_summary)
               VALUES ('target_test_sent','notification_target',?,'user',?,?,?)""",
            (
                target_id,
                str(current_user.id),
                request_id,
                json_dump({"provider": result.get("provider"), "messages": 1}),
            ),
        )
        conn.commit()
        return jsonify(
            {
                "id": target_id,
                "sent": True,
                "provider": result.get("provider"),
                "webhook_fingerprint": target.get("webhook_fingerprint"),
            }
        )
    finally:
        conn.close()


def _owned_image_path(conn, job_id: str, page: int) -> Path | None:
    row = conn.execute(
        "SELECT order_id,image_paths_json FROM order_notification_jobs WHERE id=?", (job_id,)
    ).fetchone()
    if not row or not _can_view_order(conn, row["order_id"]):
        return None
    try:
        paths = json.loads(row["image_paths_json"] or "[]")
        value = Path(paths[page - 1]).resolve()
    except (ValueError, TypeError, IndexError):
        return None
    root = Path(os.environ.get("ORDER_NOTIFICATION_IMAGE_DIR") or (Path("var") / "order-cards")).resolve()
    try:
        value.relative_to(root)
    except ValueError:
        return None
    return value if value.is_file() else None


@order_notification_bp.route("/api/order-notifications/<job_id>/image/<int:page>")
@notification_super_admin_required
def notification_image(job_id, page):
    conn = get_conn()
    try:
        path = _owned_image_path(conn, job_id, page)
    finally:
        conn.close()
    if not path:
        abort(404)
    return send_file(path, mimetype="image/png", max_age=0, conditional=True)


@order_notification_bp.route("/order-notifications")
@notification_super_admin_required
def notification_dashboard():
    status = request.args.get("status", "").strip()
    params = []
    conditions = []
    if status:
        conditions.append("j.status=?")
        params.append(status)
    allowed = _view_sources()
    if allowed is not None:
        if not allowed:
            conditions.append("1=0")
        else:
            conditions.append("j.store_id IN (" + ",".join("?" for _ in allowed) + ")")
            params.extend(allowed)
    where = "WHERE " + " AND ".join(conditions) if conditions else ""
    conn = get_conn()
    try:
        rows = conn.execute(
            f"""SELECT j.id,j.event_type,j.store_id,j.order_id,j.status,j.template_version,
                       j.created_at,j.scheduled_at,j.sent_at,j.image_bytes,j.last_error_code,j.last_error_summary,
                       t.name AS target_name,t.channel_type,
                       q.status AS queue_status,q.attempts AS queue_attempts,
                       q.max_attempts AS queue_max_attempts,q.available_at AS next_retry_at
                FROM order_notification_jobs j
                LEFT JOIN notification_targets t ON t.id=j.target_id
                LEFT JOIN oms_integration_jobs q ON q.id=j.queue_job_id
                {where} ORDER BY COALESCE(j.queue_job_id,0) DESC,j.created_at DESC LIMIT 200""",
            params,
        ).fetchall()
        count_params = []
        count_where = ""
        if allowed is not None:
            if not allowed:
                count_where = "WHERE 1=0"
            else:
                count_where = "WHERE store_id IN (" + ",".join("?" for _ in allowed) + ")"
                count_params = list(allowed)
        counters = {
            row["status"]: row["n"]
            for row in conn.execute(
                f"SELECT status,COUNT(*) n FROM order_notification_jobs {count_where} GROUP BY status",
                count_params,
            ).fetchall()
        }
        control = _configuration_snapshot(conn)
        control["queue_count"] = conn.execute(
            """SELECT COUNT(*) FROM oms_integration_jobs
               WHERE job_type='ORDER_NOTIFICATION' AND status IN ('pending','running','retry')"""
        ).fetchone()[0]
        control["retry_count"] = int(counters.get("RETRY_WAIT", 0))
        control["manual_review_count"] = int(counters.get("MANUAL_REVIEW", 0))
        control["dead_letter_count"] = int(counters.get("DEAD_LETTER", 0))
        control["exception_count"] = (
            control["retry_count"]
            + control["manual_review_count"]
            + control["dead_letter_count"]
        )
        control["audit_count"] = conn.execute(
            "SELECT COUNT(*) FROM notification_audit_logs"
        ).fetchone()[0]
        return render_template(
            "order_notifications.html",
            jobs=[dict(row) for row in rows],
            counters=counters,
            selected_status=status,
            control=control,
        )
    finally:
        conn.close()
