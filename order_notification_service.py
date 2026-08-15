"""Order notification domain: authoritative snapshots, routing and jobs."""

from __future__ import annotations

import hashlib
import json
import os
import re
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from fulfillment_common import json_load, utcnow
from fulfillment_service import enqueue_job
from order_notification_provider import ProviderError, provider_for
from order_notification_renderer import TEMPLATE_VERSION, render_order_cards
from order_notification_email import EmailRenderError, render_logged_admin_email


EVENT_TYPES = {
    "ORDER_READY",
    "ORDER_UPDATED",
    "ORDER_CANCELLED",
    "ORDER_HOLD",
    "MANUAL_RESEND",
}
AUTOMATIC_EVENT_TYPES = frozenset({"ORDER_READY"})
NOTIFICATION_MAX_ATTEMPTS = 12
MAX_NOTIFICATION_RETRY_SECONDS = 1800
NOTIFICATION_ALERT_AFTER_ATTEMPTS = 3
NOTIFICATION_ALERT_MAX_ATTEMPTS = 6
TERMINAL_STATUSES = {"SENT", "READY_PREVIEW", "READY_MANUAL", "MANUAL_REVIEW", "DEAD_LETTER", "SKIPPED"}
ACTIVE_STATUSES = {
    "PENDING",
    "DEBOUNCING",
    "VALIDATING",
    "RENDERING",
    "READY_TO_SEND",
    "SENDING",
    "RETRY_WAIT",
}
DEFAULT_POLICY = {
    "ready_statuses": ["processing", "offline"],
    "cancel_statuses": ["cancelled", "refunded"],
    "hold_statuses": ["on-hold", "failed", "pending"],
    "excluded_statuses": ["checkout-draft", "trash", "cheat"],
    "allow_site_cod_on_hold": True,
    "updated_fields": [
        "status",
        "warehouse_id",
        "shipping_method",
        "payment_method",
        "currency",
        "total",
        "recipient",
        "items",
        "customer_note",
    ],
}


class NotificationRetry(RuntimeError):
    def __init__(self, message: str, *, code: str, delay_seconds: int = 30):
        super().__init__(message)
        self.code = code
        self.delay_seconds = delay_seconds


class NotificationPermanent(RuntimeError):
    def __init__(self, message: str, *, code: str):
        super().__init__(message)
        self.code = code


ALERT_REASON_LABELS = {
    "admin_new_order_email_not_found": "尚未找到管理员新订单邮件",
    "email_body_missing": "管理员新订单邮件正文不可用",
    "email_log_credentials_missing": "邮件日志只读凭据不可用",
    "email_log_unavailable": "暂时无法读取邮件日志",
    "email_log_list_failed": "邮件日志列表读取失败",
    "email_log_detail_failed": "邮件正文读取失败",
    "rate_limited": "通知发送达到限速",
    "connect_timeout": "连接企业微信超时",
    "connection_error": "连接企业微信失败",
    "provider_transient": "企业微信服务暂时不可用",
    "provider_busy": "企业微信服务繁忙",
    "order_sequence_wait": "正在等待同订单较早任务",
}


def _json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _setting(conn, key: str, default: str = "") -> str:
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return str(row[0]) if row and row[0] is not None else default


def flag_enabled(conn, key: str, default: bool = False) -> bool:
    fallback = "1" if default else "0"
    return _setting(conn, key, fallback).strip().lower() in {"1", "true", "yes", "on"}


def notification_schema_exists(conn) -> bool:
    return bool(
        conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='order_notification_jobs'"
        ).fetchone()
    )


def load_policy(conn) -> dict:
    raw = _setting(conn, "order_notification_policy_json", "")
    if not raw:
        return dict(DEFAULT_POLICY)
    try:
        configured = json.loads(raw)
    except (TypeError, ValueError):
        return dict(DEFAULT_POLICY)
    policy = dict(DEFAULT_POLICY)
    if isinstance(configured, dict):
        for key in policy:
            if key in configured:
                policy[key] = configured[key]
    return policy


def _parse(value: Any, default: Any):
    if isinstance(value, type(default)):
        return value
    try:
        parsed = json.loads(value or "")
        return parsed if isinstance(parsed, type(default)) else default
    except (TypeError, ValueError):
        return default


def _mask_name(value: str) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip())
    if not text:
        return ""
    parts = text.split(" ")
    return " ".join(part[0] + ("*" * max(1, min(4, len(part) - 1))) for part in parts)


def _mask_phone(value: str) -> str:
    text = str(value or "").strip()
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    return ("+" if text.startswith("+") else "") + "*** *** " + digits[-3:]


def _variation(item: dict) -> str:
    values = []
    for meta in item.get("meta_data") or []:
        if not isinstance(meta, dict) or str(meta.get("key", "")).startswith("_"):
            continue
        value = str(meta.get("display_value") or meta.get("value") or "").strip()
        if value and value not in values:
            values.append(value)
    return " / ".join(values[:3])


def authoritative_snapshot(conn, order_id: str) -> dict:
    """Read the minimum fulfillment card snapshot from the local authority."""
    order = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
    if not order:
        raise NotificationPermanent("订单不存在", code="order_not_found")
    row = dict(order)
    billing = _parse(row.get("billing"), {})
    shipping = _parse(row.get("shipping"), {})
    address = dict(billing)
    address.update({key: value for key, value in shipping.items() if value not in (None, "")})
    shipping_lines = _parse(row.get("shipping_lines"), [])
    line_items = _parse(row.get("line_items"), [])
    site = conn.execute(
        "SELECT id,url,manager,country,cod_on_hold_is_shipped FROM sites WHERE url=?",
        (row.get("source"),),
    ).fetchone()
    warehouse = None
    if row.get("warehouse_id"):
        warehouse = conn.execute(
            "SELECT id,name,country FROM warehouses WHERE id=?", (row["warehouse_id"],)
        ).fetchone()
    shipping_line = shipping_lines[0] if shipping_lines and isinstance(shipping_lines[0], dict) else {}
    source = str(row.get("source") or "")
    base_url = os.environ.get("ORDER_SYSTEM_BASE_URL", "").rstrip("/")
    detail_url = f"{base_url}/orders?order_id={order_id}" if base_url else f"/orders?order_id={order_id}"
    items = []
    for item in line_items:
        if not isinstance(item, dict):
            continue
        items.append(
            {
                "line_item_id": str(item.get("id") or ""),
                "sku": str(item.get("sku") or "")[:100],
                "name": str(item.get("name") or "")[:240],
                "variation": _variation(item)[:160],
                "quantity": int(item.get("quantity") or 0),
            }
        )
    return {
        "order_id": str(row["id"]),
        "woo_id": row.get("woo_id"),
        "number": str(row.get("number") or row["id"]),
        "store_id": source,
        "store_label": source.replace("https://", "").replace("http://", "").rstrip("/"),
        "site_id": site["id"] if site else None,
        "site_country": site["country"] if site else None,
        "site_manager": site["manager"] if site else None,
        "site_cod_on_hold_is_shipped": bool(site and site["cod_on_hold_is_shipped"]),
        "status": str(row.get("status") or ""),
        "created_at": row.get("date_created"),
        "date_modified": row.get("date_modified") or row.get("updated_at") or "",
        "warehouse_id": row.get("warehouse_id"),
        "warehouse_name": warehouse["name"] if warehouse else "",
        "currency": str(row.get("currency") or ""),
        "total": str(row.get("total") or "0"),
        "cod_amount": str(row.get("total") or "0") if str(row.get("payment_method") or "") == "cod" else "0",
        "payment_method": str(row.get("payment_method_title") or row.get("payment_method") or ""),
        "payment_method_code": str(row.get("payment_method") or ""),
        "set_paid": bool(row.get("set_paid")),
        "shipping_method": str(
            shipping_line.get("method_title") or shipping_line.get("method_id") or ""
        )[:160],
        "recipient": {
            "name_masked": _mask_name(
                " ".join(str(address.get(k) or "") for k in ("first_name", "last_name"))
            ),
            "phone_masked": _mask_phone(address.get("phone") or billing.get("phone") or ""),
            "city": str(address.get("city") or "")[:100],
            "postal_code": str(address.get("postcode") or "")[:30],
            "delivery_point": str(
                address.get("address_2") or shipping_line.get("instance_id") or ""
            )[:120],
        },
        "items": items,
        "customer_note": str(row.get("customer_note") or "")[:500],
        "internal_order_url": detail_url,
    }


def snapshot_hash(snapshot: dict) -> str:
    stable = dict(snapshot)
    stable.pop("notification_at", None)
    stable.pop("changes", None)
    return hashlib.sha256(_json(stable).encode("utf-8")).hexdigest()


def order_version(snapshot: dict) -> str:
    return f"{snapshot.get('date_modified') or 'unknown'}:{snapshot_hash(snapshot)[:16]}"


def fulfillment_class(snapshot: dict, policy: dict) -> str:
    status = snapshot.get("status")
    if status in set(policy.get("excluded_statuses") or []):
        return "excluded"
    if status in set(policy.get("cancel_statuses") or []):
        return "cancelled"
    if status in set(policy.get("ready_statuses") or []):
        return "ready"
    if (
        status == "on-hold"
        and policy.get("allow_site_cod_on_hold")
        and snapshot.get("site_cod_on_hold_is_shipped")
        and snapshot.get("payment_method_code") == "cod"
    ):
        return "ready"
    if status in set(policy.get("hold_statuses") or []):
        return "hold"
    return "excluded"


def _latest_effective(conn, order_id: str):
    return conn.execute(
        """SELECT * FROM order_notification_jobs
           WHERE order_id=? AND status IN ('SENT','READY_PREVIEW','READY_MANUAL','READY_TO_SEND')
           ORDER BY COALESCE(queue_job_id,0) DESC,created_at DESC LIMIT 1""",
        (order_id,),
    ).fetchone()


def _automatic_ready_exists(conn, order_id: str) -> bool:
    """Treat any prior non-test new-order job as consumed, including failures.

    A route failure or disabled-send preview must never be silently flushed later.
    Test-send jobs are excluded so reviewing an order does not consume its future
    production notification.
    """
    return bool(
        conn.execute(
            """SELECT 1
               FROM order_notification_jobs j
               WHERE j.order_id=? AND j.event_type='ORDER_READY'
                 AND NOT EXISTS (
                     SELECT 1 FROM notification_targets t
                     WHERE t.id=j.target_id AND t.environment='test'
                 )
               LIMIT 1""",
            (order_id,),
        ).fetchone()
    )


def _automatic_rollout_reason(conn, snapshot: dict) -> str | None:
    """Gate automation by a per-store WooCommerce ID activation watermark."""
    raw = _setting(conn, "order_notification_auto_watermarks_json", "")
    try:
        configured = json.loads(raw) if raw else {}
    except (TypeError, ValueError):
        return "rollout_watermark_invalid"
    if not isinstance(configured, dict):
        return "rollout_watermark_invalid"
    entry = configured.get(str(snapshot.get("store_id") or ""))
    if not isinstance(entry, dict):
        return "rollout_not_active"
    try:
        cutoff = int(entry["max_woo_id"])
        woo_id = int(snapshot["woo_id"])
    except (KeyError, TypeError, ValueError):
        return "rollout_watermark_invalid"
    if woo_id <= cutoff:
        return "before_activation_watermark"
    return None


def _changes(previous: dict | None, current: dict, fields: list[str]) -> list[dict]:
    if not previous:
        return []
    changes = []
    for field in fields:
        before, after = previous.get(field), current.get(field)
        if _json(before) != _json(after):
            changes.append(
                {
                    "field": field,
                    "before": before if isinstance(before, (str, int, float, type(None))) else "已变更",
                    "after": after if isinstance(after, (str, int, float, type(None))) else "已变更",
                }
            )
    return changes


def determine_event(conn, snapshot: dict, requested_event: str | None = None) -> tuple[str | None, list[dict]]:
    policy = load_policy(conn)
    state = fulfillment_class(snapshot, policy)
    latest = _latest_effective(conn, snapshot["order_id"])
    previous = json_load(latest["snapshot_json"], {}) if latest else None
    changes = _changes(previous, snapshot, list(policy.get("updated_fields") or []))

    if requested_event:
        if requested_event not in EVENT_TYPES - {"MANUAL_RESEND"}:
            raise NotificationPermanent("事件类型无效", code="event_type_invalid")
        allowed = {
            "ORDER_READY": state == "ready",
            "ORDER_UPDATED": bool(latest and state == "ready" and changes),
            "ORDER_CANCELLED": bool(latest and state == "cancelled"),
            "ORDER_HOLD": state == "hold",
        }
        return (requested_event, changes) if allowed.get(requested_event) else (None, [])
    if state == "cancelled":
        return ("ORDER_CANCELLED", changes) if latest else (None, [])
    if state == "hold":
        return "ORDER_HOLD", changes
    if state != "ready":
        return None, []
    if not latest:
        return "ORDER_READY", []
    if latest["snapshot_hash"] != snapshot_hash(snapshot) and changes:
        return "ORDER_UPDATED", changes
    return None, []


def resolve_target(
    conn,
    snapshot: dict,
    *,
    environment: str | None = None,
) -> tuple[dict | None, str | None]:
    """Resolve one route, optionally inside a hard test/production boundary."""
    if environment not in {None, "test", "production"}:
        raise NotificationPermanent("通知目标环境无效", code="target_environment_invalid")
    rows = conn.execute(
        """SELECT * FROM notification_targets
             WHERE enabled=1 AND deleted_at IS NULL"""
    ).fetchall()
    candidates = []
    for row in rows:
        target = dict(row)
        if environment is not None and target.get("environment") != environment:
            continue
        if target.get("store_id") and target["store_id"] != snapshot.get("store_id"):
            continue
        if not target_matches_manager(target, snapshot.get("site_manager")):
            continue
        if target.get("warehouse_id") is not None and target["warehouse_id"] != snapshot.get("warehouse_id"):
            continue
        if target.get("shipping_method") and target["shipping_method"].casefold() != str(snapshot.get("shipping_method") or "").casefold():
            continue
        # Every higher-level dimension outweighs every possible combination
        # below it. A site route therefore overrides a manager route, while a
        # manager route overrides warehouse/shipping fallbacks.
        score = (
            8 * bool(target.get("store_id"))
            + 4 * (target.get("manager_scope") == "selected")
            + 2 * (target.get("warehouse_id") is not None)
            + bool(target.get("shipping_method"))
        )
        candidates.append((score, target))
    if not candidates:
        return None, "route_missing"
    best_score = max(score for score, _ in candidates)
    best = [target for score, target in candidates if score == best_score]
    if len(best) != 1:
        return None, "route_ambiguous"
    return best[0], None


def target_manager_names(target: dict) -> tuple[str, ...]:
    values = json_load(target.get("manager_names_json"), []) or []
    if not isinstance(values, list):
        return ()
    return tuple(
        sorted(
            {
                str(value).strip()
                for value in values
                if isinstance(value, str) and str(value).strip()
            },
            key=str.casefold,
        )
    )


def target_matches_manager(target: dict, manager_name: str | None) -> bool:
    if target.get("manager_scope") != "selected":
        return True
    return str(manager_name or "").strip() in target_manager_names(target)


def target_matches_snapshot(target: dict, snapshot: dict) -> bool:
    """Return whether an explicitly selected target still covers this order."""
    if target.get("store_id") and target["store_id"] != snapshot.get("store_id"):
        return False
    if not target_matches_manager(target, snapshot.get("site_manager")):
        return False
    if (
        target.get("warehouse_id") is not None
        and target["warehouse_id"] != snapshot.get("warehouse_id")
    ):
        return False
    if target.get("shipping_method") and target["shipping_method"].casefold() != str(
        snapshot.get("shipping_method") or ""
    ).casefold():
        return False
    return True


def _schedule_after(seconds: int) -> str:
    return (datetime.now(timezone.utc) + timedelta(seconds=max(0, seconds))).replace(microsecond=0).isoformat()


def _notification_retry_delay(queue_job: dict, base_seconds: int = 30) -> int:
    """Exponential retry window capped at 30 minutes between attempts."""
    try:
        attempts = max(1, int(queue_job.get("attempts") or 1))
    except (TypeError, ValueError):
        attempts = 1
    exponent = min(attempts - 1, 6)
    return min(
        MAX_NOTIFICATION_RETRY_SECONDS,
        max(1, int(base_seconds)) * (2 ** exponent),
    )


def _alert_reason(error_code: str | None) -> str:
    code = re.sub(r"[^A-Za-z0-9_.-]", "", str(error_code or ""))[:80]
    if code in ALERT_REASON_LABELS:
        return ALERT_REASON_LABELS[code]
    return f"后台处理异常（{code}）" if code else "后台处理异常"


def enqueue_notification_failure_alert(
    conn,
    queue_job: dict,
    *,
    phase: str,
    error_code: str | None,
) -> dict:
    """Queue one deduplicated, privacy-minimized WeCom alert for a production order."""
    if phase not in {"delayed", "final"}:
        return {"created": False, "reason": "phase_invalid"}
    if queue_job.get("aggregate_type") != "order_notification" or not queue_job.get("aggregate_id"):
        return {"created": False, "reason": "not_notification"}
    attempts = int(queue_job.get("attempts") or 0)
    max_attempts = int(queue_job.get("max_attempts") or NOTIFICATION_MAX_ATTEMPTS)
    if phase == "delayed" and attempts < NOTIFICATION_ALERT_AFTER_ATTEMPTS:
        return {"created": False, "reason": "below_threshold"}
    row = conn.execute(
        """SELECT j.id,j.snapshot_json,j.target_id,t.channel_type,t.environment,t.enabled
             FROM order_notification_jobs j
             LEFT JOIN notification_targets t ON t.id=j.target_id
            WHERE j.id=?""",
        (queue_job["aggregate_id"],),
    ).fetchone()
    if not row:
        return {"created": False, "reason": "notification_missing"}
    current = dict(row)
    snapshot = json_load(current.get("snapshot_json"), {}) or {}
    if snapshot.get("_notification_mode") == "test_send":
        return {"created": False, "reason": "test_send_excluded"}
    if (
        current.get("channel_type") != "WECOM_BOT"
        or current.get("environment") != "production"
        or not current.get("enabled")
        or not flag_enabled(conn, "order_notification_send_enabled")
    ):
        return {"created": False, "reason": "production_wecom_not_enabled"}
    idempotency_key = f"order-notification-alert:{current['id']}:{phase}"
    existing = conn.execute(
        "SELECT id FROM oms_integration_jobs WHERE idempotency_key=?",
        (idempotency_key,),
    ).fetchone()
    payload = {
        "notification_job_id": current["id"],
        "phase": phase,
        "attempts": attempts,
        "max_attempts": max_attempts,
        "error_code": str(error_code or "failed")[:100],
    }
    alert_job_id = enqueue_job(
        conn,
        "ORDER_NOTIFICATION_ALERT",
        "order_notification_alert",
        current["id"],
        idempotency_key,
        payload,
        max_attempts=NOTIFICATION_ALERT_MAX_ATTEMPTS,
    )
    if not existing:
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,request_id,after_summary)
               VALUES ('notification_failure_alert_queued','order_notification_job',?,'system',?,?)""",
            (
                current["id"],
                phase,
                _json(
                    {
                        "alert_queue_job_id": alert_job_id,
                        "phase": phase,
                        "attempts": attempts,
                        "max_attempts": max_attempts,
                        "error_code": payload["error_code"],
                    }
                ),
            ),
        )
    return {"created": not bool(existing), "queue_job_id": alert_job_id, "phase": phase}


def process_notification_alert(conn, job: dict, payload: dict, *, session=None) -> dict:
    """Send one operational text alert without exposing customer/order contents."""
    notification_job_id = payload.get("notification_job_id") or job.get("aggregate_id")
    phase = str(payload.get("phase") or "")
    if phase not in {"delayed", "final"}:
        raise NotificationPermanent("通知异常告警阶段无效", code="alert_phase_invalid")
    already_sent = conn.execute(
        """SELECT 1 FROM notification_audit_logs
            WHERE action='notification_failure_alert_sent'
              AND object_type='order_notification_job' AND object_id=? AND request_id=?""",
        (notification_job_id, phase),
    ).fetchone()
    if already_sent:
        return {"noop": "already_sent", "phase": phase}
    row = conn.execute(
        """SELECT j.status,j.snapshot_json,t.*
             FROM order_notification_jobs j
             JOIN notification_targets t ON t.id=j.target_id
            WHERE j.id=?""",
        (notification_job_id,),
    ).fetchone()
    if not row:
        raise NotificationPermanent("原通知任务或目标不存在", code="alert_target_missing")
    current = dict(row)
    if current["status"] == "SENT":
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,request_id,after_summary)
               VALUES ('notification_failure_alert_skipped','order_notification_job',?,'system',?,?)""",
            (notification_job_id, phase, _json({"reason": "image_recovered"})),
        )
        conn.commit()
        return {"skipped": "image_recovered", "phase": phase}
    snapshot = json_load(current.get("snapshot_json"), {}) or {}
    if (
        snapshot.get("_notification_mode") == "test_send"
        or current.get("channel_type") != "WECOM_BOT"
        or current.get("environment") != "production"
        or not current.get("enabled")
        or not flag_enabled(conn, "order_notification_send_enabled")
    ):
        raise NotificationPermanent("异常告警生产发送条件不满足", code="alert_send_blocked")
    site = str(snapshot.get("store_label") or snapshot.get("store_id") or "未知站点")[:160]
    order_number = str(snapshot.get("number") or snapshot.get("woo_id") or "未知")[:80]
    attempts = max(0, int(payload.get("attempts") or 0))
    max_attempts = max(0, int(payload.get("max_attempts") or 0))
    reason = _alert_reason(payload.get("error_code"))
    if phase == "delayed":
        content = (
            "⚠️ 订单图片推送延迟\n"
            f"站点：{site}\n"
            f"订单：#{order_number}\n"
            f"状态：图片暂未推送，系统仍在自动重试（{attempts}/{max_attempts}）\n"
            f"原因：{reason}"
        )
    else:
        content = (
            "🚨 订单图片推送失败\n"
            f"站点：{site}\n"
            f"订单：#{order_number}\n"
            f"状态：已重试 {attempts}/{max_attempts} 次，需人工处理\n"
            f"原因：{reason}\n"
            "请登录订单系统的“群通知”页面查看。"
        )
    provider = provider_for("WECOM_BOT", session=session)
    try:
        result = provider.send_text(content, current)
    except ProviderError as exc:
        action = (
            "notification_failure_alert_unknown"
            if exc.unknown_outcome
            else "notification_failure_alert_retry"
            if exc.retryable
            else "notification_failure_alert_failed"
        )
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,request_id,after_summary)
               VALUES (?,'order_notification_job',?,'system',?,?)""",
            (action, notification_job_id, phase, _json({"provider_error_code": exc.code})),
        )
        conn.commit()
        if exc.retryable:
            raise NotificationRetry(
                "异常告警发送暂时失败",
                code=exc.code,
                delay_seconds=_notification_retry_delay(job),
            ) from exc
        raise NotificationPermanent("异常告警发送失败", code=exc.code) from exc
    conn.execute(
        """INSERT INTO notification_audit_logs
           (action,object_type,object_id,actor_type,request_id,after_summary)
           VALUES ('notification_failure_alert_sent','order_notification_job',?,'system',?,?)""",
        (
            notification_job_id,
            phase,
            _json(
                {
                    "phase": phase,
                    "attempts": attempts,
                    "max_attempts": max_attempts,
                    "provider": result.get("provider"),
                }
            ),
        ),
    )
    conn.commit()
    return {"accepted": True, "phase": phase, "provider": result.get("provider")}


def create_job_for_order(
    conn,
    order_id: str,
    *,
    event_id: str,
    requested_event: str | None = None,
    resend_of: str | None = None,
    actor: dict | None = None,
    allowed_event_types: frozenset[str] | set[str] | None = None,
    target_environment: str | None = None,
) -> dict:
    if not notification_schema_exists(conn):
        return {"created": False, "reason": "schema_missing"}
    if not flag_enabled(conn, "order_notification_enabled"):
        return {"created": False, "reason": "disabled"}
    snapshot = authoritative_snapshot(conn, order_id)
    automatic_new_order_only = (
        not resend_of
        and target_environment == "production"
        and allowed_event_types is not None
        and frozenset(allowed_event_types) == AUTOMATIC_EVENT_TYPES
    )
    if automatic_new_order_only:
        rollout_reason = _automatic_rollout_reason(conn, snapshot)
        if rollout_reason:
            return {"created": False, "reason": rollout_reason}
        if _automatic_ready_exists(conn, order_id):
            return {"created": False, "reason": "already_notified"}
    if resend_of:
        original = conn.execute(
            "SELECT * FROM order_notification_jobs WHERE id=? AND order_id=?", (resend_of, order_id)
        ).fetchone()
        if not original:
            raise NotificationPermanent("原通知不存在", code="resend_source_missing")
        event_type, changes = "MANUAL_RESEND", []
        resend_sequence = conn.execute(
            "SELECT COUNT(*) FROM order_notification_jobs WHERE resend_of=?", (resend_of,)
        ).fetchone()[0] + 1
    else:
        event_type, changes = determine_event(conn, snapshot, requested_event)
        resend_sequence = 0
    if not event_type:
        return {"created": False, "reason": "not_eligible"}
    if allowed_event_types is not None and event_type not in allowed_event_types:
        return {"created": False, "reason": "event_type_not_enabled"}
    target, route_error = resolve_target(
        conn, snapshot, environment=target_environment
    )
    version = order_version(snapshot)
    target_key = target["id"] if target else route_error
    idem_source = f"{snapshot['store_id']}|{order_id}|{event_type}|{version}|{target_key}|{resend_sequence}"
    idem = hashlib.sha256(idem_source.encode("utf-8")).hexdigest()
    existing = conn.execute(
        "SELECT * FROM order_notification_jobs WHERE idempotency_key=?", (idem,)
    ).fetchone()
    if existing:
        return {"created": False, "duplicate": True, "job": dict(existing)}

    job_id = uuid.uuid4().hex
    debounce = int(_setting(conn, "order_notification_debounce_seconds", "45") or 45)
    scheduled_at = _schedule_after(0 if resend_of else debounce)
    template = _setting(conn, "order_notification_template_version", TEMPLATE_VERSION)
    snap_hash = snapshot_hash(snapshot)
    status = "DEAD_LETTER" if route_error else ("PENDING" if debounce == 0 or resend_of else "DEBOUNCING")
    conn.execute(
        """INSERT INTO order_notification_jobs
           (id,event_id,event_type,store_id,order_id,order_version,target_id,
            idempotency_key,status,snapshot_json,snapshot_hash,changed_fields_json,
            template_version,scheduled_at,resend_of,resend_sequence,last_error_code,last_error_summary)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            job_id, event_id, event_type, snapshot["store_id"], order_id, version,
            target["id"] if target else None, idem, status, _json(snapshot), snap_hash,
            _json(changes), template, scheduled_at, resend_of, resend_sequence,
            route_error, "没有唯一通知目标" if route_error else None,
        ),
    )
    queue_job_id = None
    if target:
        queue_job_id = enqueue_job(
            conn,
            "ORDER_NOTIFICATION",
            "order_notification",
            job_id,
            f"order-notification:{idem}",
            {"notification_job_id": job_id},
            available_at=scheduled_at,
            max_attempts=NOTIFICATION_MAX_ATTEMPTS,
        )
        conn.execute(
            "UPDATE order_notification_jobs SET queue_job_id=? WHERE id=?", (queue_job_id, job_id)
        )
    actor = actor or {"type": "system", "id": None}
    conn.execute(
        """INSERT INTO notification_audit_logs
           (action,object_type,object_id,actor_type,actor_id,after_summary)
           VALUES ('job_created','order_notification_job',?,?,?,?)""",
        (
            job_id,
            actor.get("type", "system"),
            actor.get("id"),
            _json({"event_type": event_type, "status": status, "route_error": route_error}),
        ),
    )
    conn.commit()
    return {"created": True, "duplicate": False, "job": dict(conn.execute("SELECT * FROM order_notification_jobs WHERE id=?", (job_id,)).fetchone())}


def create_test_send_job(
    conn,
    order_id: str,
    *,
    target_id: str,
    preview_id: str,
    render_source: str,
    actor: dict,
) -> dict:
    """Queue one audited preview for an explicit isolated WeCom test target.

    This path deliberately does not require the automatic-card flag. It can
    therefore exercise the durable worker without opening the sync hook that
    would enqueue unrelated orders. The worker still requires the separately
    locked test-send flag before any provider call.
    """
    if not notification_schema_exists(conn):
        return {"created": False, "reason": "schema_missing"}
    if render_source not in {"email", "system_card"}:
        raise NotificationPermanent("测试图片来源无效", code="render_source_invalid")

    snapshot = authoritative_snapshot(conn, order_id)
    event_type, changes = determine_event(conn, snapshot, "ORDER_READY")
    if not event_type:
        return {"created": False, "reason": "not_eligible"}

    target_row = conn.execute(
        "SELECT * FROM notification_targets WHERE id=? AND enabled=1", (target_id,)
    ).fetchone()
    if not target_row:
        raise NotificationPermanent("测试目标不存在或未启用", code="test_target_missing")
    target = dict(target_row)
    if target.get("environment") != "test":
        raise NotificationPermanent("只允许隔离测试目标", code="test_target_required")
    if target.get("channel_type") != "WECOM_BOT":
        raise NotificationPermanent("测试群必须使用企业微信群机器人", code="test_target_channel_invalid")
    if not target_matches_snapshot(target, snapshot):
        raise NotificationPermanent("测试目标与订单路由不匹配", code="test_target_scope_mismatch")

    version = order_version(snapshot)
    idem_source = (
        f"TEST_SEND|{preview_id}|{snapshot['store_id']}|{order_id}|"
        f"{version}|{target_id}|{render_source}"
    )
    idem = hashlib.sha256(idem_source.encode("utf-8")).hexdigest()
    existing = conn.execute(
        "SELECT * FROM order_notification_jobs WHERE idempotency_key=?", (idem,)
    ).fetchone()
    if existing:
        return {"created": False, "duplicate": True, "job": dict(existing)}

    job_id = uuid.uuid4().hex
    scheduled_at = _schedule_after(0)
    template = _setting(conn, "order_notification_template_version", TEMPLATE_VERSION)
    snap_hash = snapshot_hash(snapshot)
    stored_snapshot = dict(snapshot)
    stored_snapshot.update(
        {
            "_notification_mode": "test_send",
            "_notification_render_source": render_source,
            "_notification_preview_id": preview_id,
        }
    )
    conn.execute(
        """INSERT INTO order_notification_jobs
           (id,event_id,event_type,store_id,order_id,order_version,target_id,
            idempotency_key,status,snapshot_json,snapshot_hash,changed_fields_json,
            template_version,scheduled_at,resend_sequence)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,0)""",
        (
            job_id,
            "test-send:" + preview_id,
            event_type,
            snapshot["store_id"],
            order_id,
            version,
            target_id,
            idem,
            "PENDING",
            _json(stored_snapshot),
            snap_hash,
            _json(changes),
            template,
            scheduled_at,
        ),
    )
    queue_job_id = enqueue_job(
        conn,
        "ORDER_NOTIFICATION",
        "order_notification",
        job_id,
        f"order-notification-test:{idem}",
        {"notification_job_id": job_id},
        available_at=scheduled_at,
        max_attempts=NOTIFICATION_MAX_ATTEMPTS,
    )
    conn.execute(
        "UPDATE order_notification_jobs SET queue_job_id=? WHERE id=?",
        (queue_job_id, job_id),
    )
    conn.execute(
        """INSERT INTO notification_audit_logs
           (action,object_type,object_id,actor_type,actor_id,request_id,after_summary)
           VALUES ('test_send_requested','order_notification_job',?,?,?,?,?)""",
        (
            job_id,
            actor.get("type", "user"),
            actor.get("id"),
            preview_id,
            _json(
                {
                    "order_id": order_id,
                    "target_id": target_id,
                    "render_source": render_source,
                    "status": "PENDING",
                }
            ),
        ),
    )
    conn.commit()
    job = conn.execute(
        "SELECT * FROM order_notification_jobs WHERE id=?", (job_id,)
    ).fetchone()
    return {"created": True, "duplicate": False, "job": dict(job)}


def enqueue_synced_orders(candidates: list[dict]) -> None:
    """Called after WooCommerce UPSERT commit; never breaks core sync."""
    if not candidates:
        return
    try:
        from fulfillment_common import get_conn

        conn = get_conn()
        try:
            if not notification_schema_exists(conn) or not flag_enabled(conn, "order_notification_enabled"):
                return
            for item in candidates:
                raw = f"{item.get('order_id')}|{item.get('date_modified')}|{item.get('status')}"
                event_id = "sync:" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:32]
                create_job_for_order(
                    conn,
                    str(item["order_id"]),
                    event_id=event_id,
                    allowed_event_types=AUTOMATIC_EVENT_TYPES,
                    target_environment="production",
                )
        finally:
            conn.close()
    except Exception as exc:
        print(f"[order-notification] enqueue skipped: {type(exc).__name__}")


def _set_job(conn, job_id: str, status: str, **updates) -> None:
    allowed = {
        "snapshot_json", "snapshot_hash", "order_version", "changed_fields_json", "image_paths_json",
        "sent_pages_json", "image_sha256", "image_width", "image_height", "image_bytes",
        "sent_at", "last_error_code", "last_error_summary", "template_version",
    }
    values = {key: value for key, value in updates.items() if key in allowed}
    assignments = ["status=?", "updated_at=CURRENT_TIMESTAMP"] + [f"{key}=?" for key in values]
    conn.execute(
        f"UPDATE order_notification_jobs SET {','.join(assignments)} WHERE id=?",
        (status, *values.values(), job_id),
    )


def _attempt(conn, job_id: str, result: str, *, http_status=None, code=None, summary=None) -> int:
    attempt_no = conn.execute(
        "SELECT COALESCE(MAX(attempt_no),0)+1 FROM order_notification_attempts WHERE job_id=?",
        (job_id,),
    ).fetchone()[0]
    now = utcnow()
    conn.execute(
        """INSERT INTO order_notification_attempts
           (job_id,attempt_no,started_at,finished_at,http_status,provider_error_code,response_summary,result)
           VALUES (?,?,?,?,?,?,?,?)""",
        (job_id, attempt_no, now, now, http_status, code, str(summary or "")[:500], result),
    )
    return attempt_no


def _rate_limited(conn, target: dict) -> bool:
    count = conn.execute(
        """SELECT COUNT(*) FROM order_notification_attempts a
           JOIN order_notification_jobs j ON j.id=a.job_id
           WHERE j.target_id=? AND a.result='SUCCESS'
             AND datetime(a.finished_at)>=datetime('now','-60 seconds')""",
        (target["id"],),
    ).fetchone()[0]
    return count >= int(target.get("rate_limit_per_minute") or 15)


def cleanup_expired_cards(conn, image_root: str | os.PathLike[str] | None = None) -> int:
    """Delete only expired private card files while preserving job/audit history."""
    days = max(1, min(365, int(_setting(conn, "order_notification_image_retention_days", "30") or 30)))
    root = Path(
        image_root
        or os.environ.get("ORDER_NOTIFICATION_IMAGE_DIR")
        or (Path("var") / "order-cards")
    ).resolve()
    rows = conn.execute(
        """SELECT id,image_paths_json FROM order_notification_jobs
           WHERE status IN ('SENT','READY_PREVIEW','READY_MANUAL','MANUAL_REVIEW','DEAD_LETTER','SKIPPED')
             AND image_paths_json IS NOT NULL AND image_paths_json<>'[]'
             AND datetime(updated_at)<datetime('now', ?)""",
        (f"-{days} days",),
    ).fetchall()
    deleted = 0
    for row in rows:
        remaining = []
        for raw_path in json_load(row["image_paths_json"], []) or []:
            try:
                path = Path(raw_path).resolve()
                path.relative_to(root)
            except (OSError, ValueError):
                remaining.append(raw_path)
                continue
            try:
                if path.is_file():
                    path.unlink()
                    deleted += 1
            except OSError:
                remaining.append(raw_path)
        conn.execute(
            "UPDATE order_notification_jobs SET image_paths_json=? WHERE id=?",
            (_json(remaining), row["id"]),
        )
    if deleted:
        conn.execute(
            """INSERT INTO notification_audit_logs
               (action,object_type,object_id,actor_type,after_summary)
               VALUES ('card_retention_cleanup','order_notification_job','batch','system',?)""",
            (_json({"deleted_files": deleted, "retention_days": days}),),
        )
    conn.commit()
    return deleted


def process_notification_job(conn, job: dict, payload: dict, *, session=None, output_dir: str | None = None) -> dict:
    card_dir = output_dir or os.environ.get("ORDER_NOTIFICATION_IMAGE_DIR") or str(Path("var") / "order-cards")
    cleanup_expired_cards(conn, card_dir)
    job_id = payload.get("notification_job_id") or job.get("aggregate_id")
    row = conn.execute("SELECT * FROM order_notification_jobs WHERE id=?", (job_id,)).fetchone()
    if not row:
        raise NotificationPermanent("通知任务不存在", code="notification_job_missing")
    current = dict(row)
    stored_snapshot = json_load(current.get("snapshot_json"), {}) or {}
    test_send_job = stored_snapshot.get("_notification_mode") == "test_send"
    render_source_override = (
        stored_snapshot.get("_notification_render_source") if test_send_job else None
    )
    if current["status"] in TERMINAL_STATUSES:
        return {"noop": current["status"]}
    sent_pages = set(json_load(current.get("sent_pages_json"), []) or [])
    stored_paths = json_load(current.get("image_paths_json"), []) or []

    earlier = conn.execute(
        """SELECT id FROM order_notification_jobs
           WHERE order_id=? AND id<>? AND queue_job_id IS NOT NULL
             AND queue_job_id < ? AND status IN
             ('PENDING','DEBOUNCING','VALIDATING','RENDERING','READY_TO_SEND','SENDING','RETRY_WAIT')
           ORDER BY queue_job_id LIMIT 1""",
        (current["order_id"], job_id, current["queue_job_id"] or 0),
    ).fetchone()
    if earlier and earlier["id"] != job_id:
        _set_job(conn, job_id, "RETRY_WAIT", last_error_code="order_sequence_wait", last_error_summary="等待同订单较早事件")
        conn.commit()
        raise NotificationRetry("等待同订单较早事件", code="order_sequence_wait", delay_seconds=15)

    _set_job(conn, job_id, "VALIDATING")
    fresh = authoritative_snapshot(conn, current["order_id"])
    fresh_hash = snapshot_hash(fresh)
    state = fulfillment_class(fresh, load_policy(conn))
    if current["event_type"] in {"ORDER_READY", "ORDER_UPDATED"} and state != "ready":
        _set_job(conn, job_id, "SKIPPED", last_error_code="no_longer_fulfillable", last_error_summary="防抖后订单不再可履约")
        conn.commit()
        return {"skipped": "no_longer_fulfillable"}
    if current["event_type"] == "MANUAL_RESEND" and state == "excluded":
        _set_job(conn, job_id, "SKIPPED", last_error_code="no_longer_notifiable", last_error_summary="订单当前状态不允许通知")
        conn.commit()
        return {"skipped": "no_longer_notifiable"}
    if current["event_type"] == "ORDER_CANCELLED" and state != "cancelled":
        _set_job(conn, job_id, "SKIPPED", last_error_code="no_longer_cancelled", last_error_summary="订单已离开取消状态")
        conn.commit()
        return {"skipped": "no_longer_cancelled"}
    if current["event_type"] == "ORDER_HOLD" and state != "hold":
        _set_job(conn, job_id, "SKIPPED", last_error_code="no_longer_hold", last_error_summary="订单已离开暂停状态")
        conn.commit()
        return {"skipped": "no_longer_hold"}

    # Once one page has been accepted, never regenerate the remaining pages:
    # doing so could create a mixed notification with different timestamps or
    # order data. A concurrent order change therefore requires operator review
    # and a subsequent audited event rather than an automatic partial resend.
    if sent_pages and fresh_hash != current["snapshot_hash"]:
        _set_job(
            conn,
            job_id,
            "MANUAL_REVIEW",
            last_error_code="order_changed_after_partial_send",
            last_error_summary="部分图片已发送后订单再次变化，禁止混合版本自动续发",
        )
        conn.commit()
        raise NotificationPermanent("部分发送后订单发生变化", code="order_changed_after_partial_send")

    # Debounce uses the latest authoritative values, not the event payload. If
    # another material edit landed during the window, rebuild the displayed
    # diff against the last effective notification instead of preserving the
    # stale event-time diff.
    changes = json_load(current.get("changed_fields_json"), []) or []
    if current["event_type"] == "ORDER_UPDATED":
        latest = _latest_effective(conn, current["order_id"])
        previous = json_load(latest["snapshot_json"], {}) if latest else None
        changes = _changes(previous, fresh, list(load_policy(conn).get("updated_fields") or []))
        if not changes:
            _set_job(conn, job_id, "SKIPPED", last_error_code="no_material_change", last_error_summary="防抖后没有需要通知的字段变化")
            conn.commit()
            return {"skipped": "no_material_change"}
    if sent_pages:
        if not stored_paths or any(not Path(path).is_file() for path in stored_paths):
            _set_job(conn, job_id, "MANUAL_REVIEW", last_error_code="partial_image_missing", last_error_summary="部分发送任务的原始图片缺失")
            conn.commit()
            raise NotificationPermanent("部分发送任务图片缺失", code="partial_image_missing")
        paths = stored_paths
        _set_job(conn, job_id, "READY_TO_SEND", last_error_code=None, last_error_summary=None)
    else:
        persisted_fresh = dict(fresh)
        if test_send_job:
            persisted_fresh.update(
                {
                    "_notification_mode": "test_send",
                    "_notification_render_source": render_source_override,
                    "_notification_preview_id": stored_snapshot.get(
                        "_notification_preview_id"
                    ),
                }
            )
        _set_job(
            conn,
            job_id,
            "RENDERING",
            snapshot_json=_json(persisted_fresh),
            snapshot_hash=fresh_hash,
            order_version=order_version(fresh),
            changed_fields_json=_json(changes),
            last_error_code=None,
            last_error_summary=None,
        )
        fresh["changes"] = changes
        fresh["notification_at"] = utcnow()
        render_source = render_source_override or _setting(
            conn, "order_notification_render_source", "email"
        )
        if render_source == "email" and current["event_type"] == "ORDER_READY":
            try:
                rendered, email_metadata = render_logged_admin_email(
                    conn, current["order_id"], card_dir, job_id
                )
            except EmailRenderError as exc:
                next_status = "RETRY_WAIT" if exc.retryable else "DEAD_LETTER"
                _set_job(
                    conn,
                    job_id,
                    next_status,
                    last_error_code=exc.code,
                    last_error_summary=str(exc)[:300],
                )
                conn.commit()
                if exc.retryable:
                    raise NotificationRetry(
                        str(exc),
                        code=exc.code,
                        delay_seconds=_notification_retry_delay(job),
                    ) from exc
                raise NotificationPermanent(str(exc), code=exc.code) from exc
            conn.execute(
                """INSERT INTO notification_audit_logs
                   (action,object_type,object_id,actor_type,after_summary)
                   VALUES ('email_source_rendered','order_notification_job',?,'system',?)""",
                (
                    job_id,
                    _json(
                        {
                            "email_log_id": email_metadata.get("log_id"),
                            "html_sha256": email_metadata.get("html_sha256"),
                            "images_inlined": email_metadata.get("images_inlined", 0),
                            "images_removed": email_metadata.get("images_removed", 0),
                            "template_version": email_metadata.get("template_version"),
                        }
                    ),
                ),
            )
            _set_job(
                conn,
                job_id,
                "RENDERING",
                template_version=email_metadata.get("template_version"),
            )
        else:
            rendered = render_order_cards(
                fresh,
                current["event_type"],
                card_dir,
                job_id,
                template_version=current["template_version"],
            )
        paths = [item["path"] for item in rendered]
        aggregate_sha = hashlib.sha256("".join(item["sha256"] for item in rendered).encode()).hexdigest()
        _set_job(
            conn,
            job_id,
            "READY_TO_SEND",
            image_paths_json=_json(paths),
            image_sha256=aggregate_sha,
            image_width=max(item["width"] for item in rendered),
            image_height=max(item["height"] for item in rendered),
            image_bytes=sum(item["bytes"] for item in rendered),
        )
    target_row = conn.execute("SELECT * FROM notification_targets WHERE id=? AND enabled=1", (current["target_id"],)).fetchone()
    if not target_row:
        _set_job(conn, job_id, "DEAD_LETTER", last_error_code="target_disabled", last_error_summary="通知目标不存在或已禁用")
        conn.commit()
        raise NotificationPermanent("通知目标不存在或已禁用", code="target_disabled")
    target = dict(target_row)
    channel = target["channel_type"]

    if (
        channel == "WECOM_BOT"
        and not test_send_job
        and target.get("environment") != "production"
    ):
        _set_job(
            conn,
            job_id,
            "DEAD_LETTER",
            last_error_code="automatic_target_environment_invalid",
            last_error_summary="非测试任务禁止发送到测试群",
        )
        conn.commit()
        raise NotificationPermanent(
            "非测试任务禁止发送到测试群",
            code="automatic_target_environment_invalid",
        )

    if test_send_job and (
        target.get("environment") != "test"
        or channel != "WECOM_BOT"
        or not target_matches_snapshot(target, fresh)
    ):
        _set_job(
            conn,
            job_id,
            "DEAD_LETTER",
            last_error_code="test_target_changed",
            last_error_summary="隔离测试目标已改变或不再匹配订单，已阻止发送",
        )
        conn.commit()
        raise NotificationPermanent(
            "隔离测试目标已改变，已阻止发送", code="test_target_changed"
        )

    if channel == "WECOM_BOT":
        flag = "order_notification_send_enabled" if target["environment"] == "production" else "order_notification_test_send_enabled"
        if not flag_enabled(conn, flag):
            _attempt(conn, job_id, "BLOCKED", code="feature_flag_off", summary="企业微信发送开关关闭")
            # Preview is terminal: enabling the flag later must not flush old
            # shadow-mode orders into a live group. An audited manual resend
            # creates a fresh task when operators intentionally want delivery.
            _set_job(conn, job_id, "READY_PREVIEW", last_error_code="feature_flag_off", last_error_summary="已生成影子卡片，未发送")
            conn.commit()
            return {"blocked": "feature_flag_off", "images": len(paths)}
    if channel == "MANUAL_WECHAT":
        provider_for(channel).send_images(paths, target)
        _attempt(conn, job_id, "MANUAL_READY", summary="图片仅供人工下载/转发")
        _set_job(conn, job_id, "READY_MANUAL")
        conn.commit()
        return {"manual_ready": True, "images": len(paths)}

    provider = provider_for(channel, session=session)
    for page_no, path in enumerate(paths, 1):
        if page_no in sent_pages:
            continue
        if _rate_limited(conn, target):
            _set_job(conn, job_id, "RETRY_WAIT", last_error_code="rate_limited", last_error_summary="应用侧限流，任务保留")
            conn.commit()
            raise NotificationRetry("应用侧限流", code="rate_limited", delay_seconds=10)
        _set_job(conn, job_id, "SENDING")
        conn.commit()
        try:
            result = provider.send_images([path], target)
        except ProviderError as exc:
            attempt_result = (
                "BLOCKED" if exc.unknown_outcome
                else "RETRYABLE" if exc.retryable
                else "PERMANENT_FAILURE"
            )
            _attempt(conn, job_id, attempt_result, http_status=exc.http_status, code=exc.code, summary=str(exc))
            status = (
                "MANUAL_REVIEW" if exc.unknown_outcome
                else "RETRY_WAIT" if exc.retryable
                else "DEAD_LETTER"
            )
            _set_job(conn, job_id, status, last_error_code=exc.code, last_error_summary=str(exc)[:300])
            conn.commit()
            if exc.retryable:
                raise NotificationRetry(
                    str(exc),
                    code=exc.code,
                    delay_seconds=_notification_retry_delay(job),
                ) from exc
            raise NotificationPermanent(str(exc), code=exc.code) from exc
        _attempt(conn, job_id, "SUCCESS", summary=_json(result))
        sent_pages.add(page_no)
        _set_job(conn, job_id, "SENDING", sent_pages_json=_json(sorted(sent_pages)))
        conn.commit()

    _set_job(conn, job_id, "SENT", sent_at=utcnow(), last_error_code=None, last_error_summary=None)
    conn.commit()
    return {"accepted": True, "images": len(paths), "provider": channel}


def notification_summary(conn, order_id: str) -> dict:
    rows = conn.execute(
        """SELECT j.*,t.name AS target_name,t.channel_type,t.environment
           FROM order_notification_jobs j LEFT JOIN notification_targets t ON t.id=j.target_id
           WHERE j.order_id=? ORDER BY COALESCE(j.queue_job_id,0) DESC,j.created_at DESC""",
        (order_id,),
    ).fetchall()
    jobs = []
    for row in rows:
        item = dict(row)
        item.pop("snapshot_json", None)
        image_paths = json_load(item.pop("image_paths_json", None), []) or []
        item["image_count"] = len(image_paths)
        item["changes"] = json_load(item.pop("changed_fields_json", None), []) or []
        item["attempts"] = [
            dict(a)
            for a in conn.execute(
                """SELECT attempt_no,started_at,finished_at,http_status,provider_error_code,response_summary,result
                   FROM order_notification_attempts WHERE job_id=? ORDER BY attempt_no""",
                (item["id"],),
            ).fetchall()
        ]
        jobs.append(item)
    return {"order_id": order_id, "latest": jobs[0] if jobs else None, "jobs": jobs}
