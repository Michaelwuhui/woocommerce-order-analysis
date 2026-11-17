"""WooCommerce 订单实时 Webhook 同步模块（配合每小时自动增量同步兜底）。

核心职责：
1. 管理各活跃 WooCommerce 站点的 `order.created` 与 `order.updated` 受管 Webhook（自动创建、密钥管理、停用自愈、重复清理）。
2. 提供毫秒级响应的 HMAC-SHA256 验签回调端点，支持 WooCommerce Ping 握手、重放去重与同单密集更新防抖合并。
3. 联动现有 Celery `sync_fetch` / `sync_write` 队列回读权威单笔订单与备注，统一经 `upsert_orders_in_transaction` 与 `run_post_commit_sync_actions` 落库并触发多仓履约锁库存、企微通知及运单对账。
"""

from __future__ import annotations

import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
import hashlib
import hmac
import json
import logging
import os
import secrets
from typing import Any, Iterable
from urllib.parse import parse_qs, urljoin, urlsplit

from flask import Blueprint, current_app, jsonify, request
from flask_login import current_user, login_required
import requests

import db_backend
from oid_utils import make_oid
from site_connection_service import ARCHIVED_STATUS, is_site_archived
from sync_site_health import clear_site_failure
from sync_utils import (
    run_post_commit_sync_actions,
    upsert_order_notes_in_transaction,
    upsert_orders_in_transaction,
)

logger = logging.getLogger(__name__)

woo_order_webhook_bp = Blueprint("woo_order_webhooks", __name__)

MANAGED_TOPICS: tuple[str, ...] = ("order.created", "order.updated")
WEBHOOK_NAME_PREFIX = "[CFWorks OMS]"
TOPIC_LABELS: dict[str, str] = {
    "order.created": "[CFWorks OMS] 订单创建实时同步",
    "order.updated": "[CFWorks OMS] 订单更新实时同步",
}
WEBHOOK_ENABLE_KEY = "woo_webhook_sync_enabled"
DEFAULT_PUBLIC_BASE_URL = "https://cfworks.vip"
MAX_WEBHOOK_BYTES = 2 * 1024 * 1024
WC_CONNECT_TIMEOUT = 5.0
WC_READ_TIMEOUT = 20.0
HEALTHY_REVERIFY_SECONDS = 6 * 3600
IGNORED_ORDER_STATUSES = {"checkout-draft", "auto-draft"}


def _is_postgres(conn) -> bool:
    """判断当前数据库连接是否为 PostgreSQL 适配器连接。"""
    return hasattr(conn, "_raw")


def get_db_conn(db_path: str | None = None):
    """获取兼容 SQLite 与 PostgreSQL 的数据库连接。"""
    target = db_path or os.getenv("WOO_SQLITE_PATH", "woocommerce_orders.db")
    conn = db_backend.connect(target, timeout=30)
    if not _is_postgres(conn):
        import sqlite3 as _stdlib_sqlite3

        conn.row_factory = _stdlib_sqlite3.Row
    return conn


def public_base_url() -> str:
    """返回供外部 WooCommerce 站点回调的权威公网 HTTPS 根地址。"""
    raw = (
        os.getenv("WOO_WEBHOOK_PUBLIC_BASE_URL") or DEFAULT_PUBLIC_BASE_URL
    ).strip().rstrip("/")
    return raw or DEFAULT_PUBLIC_BASE_URL


def build_delivery_url(site_id: int, base_url: str | None = None) -> str:
    """构造指定站点的受管 Webhook 投递回调 URL。"""
    root = (base_url or public_base_url()).strip().rstrip("/")
    return f"{root}/api/webhooks/woocommerce/orders/{int(site_id)}"


def webhook_tables_exist(conn) -> bool:
    """检查 Webhook 管理表是否已完成初始化或迁移。"""
    if _is_postgres(conn):
        row = conn.execute(
            "SELECT to_regclass('public.woo_site_webhooks'), "
            "to_regclass('public.woo_webhook_deliveries')"
        ).fetchone()
        return bool(row and row[0] and row[1])
    rows = {
        str(r[0])
        for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' "
            "AND name IN ('woo_site_webhooks','woo_webhook_deliveries')"
        ).fetchall()
    }
    return {"woo_site_webhooks", "woo_webhook_deliveries"}.issubset(rows)


def init_woo_order_webhook_tables(conn) -> None:
    """在 SQLite 测试或回退环境中初始化 Webhook 相关表结构。"""
    if _is_postgres(conn):
        return
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS woo_site_webhooks (
            site_id INTEGER NOT NULL,
            topic TEXT NOT NULL,
            secret TEXT NOT NULL,
            previous_secret TEXT,
            remote_webhook_id INTEGER,
            delivery_url TEXT NOT NULL,
            remote_status TEXT NOT NULL DEFAULT 'unknown',
            last_synced_at TEXT,
            last_ping_at TEXT,
            last_delivery_at TEXT,
            last_error TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (site_id, topic)
        )
        """
    )
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_woo_site_webhooks_status
        ON woo_site_webhooks(remote_status, last_synced_at)
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS woo_webhook_deliveries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            delivery_key TEXT NOT NULL UNIQUE,
            site_id INTEGER NOT NULL,
            woo_order_id INTEGER,
            topic TEXT NOT NULL,
            event TEXT,
            delivery_id TEXT,
            payload_hash TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'queued',
            order_oid TEXT,
            order_status TEXT,
            latency_ms INTEGER,
            error_message TEXT,
            received_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            processed_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_woo_webhook_deliveries_received
        ON woo_webhook_deliveries(received_at DESC)
        """
    )
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_woo_webhook_deliveries_site_received
        ON woo_webhook_deliveries(site_id, received_at DESC)
        """
    )
    conn.execute(
        """
        INSERT INTO settings (key, value)
        VALUES ('woo_webhook_sync_enabled', 'true')
        ON CONFLICT(key) DO NOTHING
        """
    )
    conn.commit()


def is_webhook_sync_enabled(conn) -> bool:
    """读取全局实时 Webhook 订单同步开关（默认开启）。"""
    row = conn.execute(
        "SELECT value FROM settings WHERE key=?", (WEBHOOK_ENABLE_KEY,)
    ).fetchone()
    if not row:
        return True
    return str(row[0] or "").strip().lower() in {"1", "true", "yes", "on"}


def set_webhook_sync_enabled(conn, enabled: bool) -> bool:
    """持久化更新全局实时 Webhook 订单同步开关。"""
    val = "true" if enabled else "false"
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (WEBHOOK_ENABLE_KEY, val),
    )
    conn.commit()
    return bool(enabled)


def compute_wc_signature(raw_body: bytes, secret: str) -> str:
    """按 WooCommerce 原生规则计算 Base64 编码的 HMAC-SHA256 签名。"""
    mac = hmac.new(secret.encode("utf-8"), raw_body, hashlib.sha256).digest()
    return base64.b64encode(mac).decode("ascii")


def verify_wc_signature(
    raw_body: bytes, signature: str | None, secrets_list: Iterable[str]
) -> bool:
    """恒定时间校验 WooCommerce Webhook 请求头签名（兼容密钥平滑轮换）。"""
    supplied = str(signature or "").strip()
    if not supplied:
        return False
    for sec in secrets_list:
        clean_sec = str(sec or "").strip()
        if not clean_sec:
            continue
        expected = compute_wc_signature(raw_body, clean_sec)
        if hmac.compare_digest(expected, supplied):
            return True
    return False


def _parse_utc_timestamp(value: Any) -> datetime | None:
    """将数据库返回的时间字符串或 datetime 统一解析为 UTC 时间对象。"""
    if not value:
        return None
    if isinstance(value, datetime):
        return (
            value.astimezone(timezone.utc)
            if value.tzinfo is not None
            else value.replace(tzinfo=timezone.utc)
        )
    text = str(value).strip()
    if not text:
        return None
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except ValueError:
        return None


def _friendly_http_error(exc: Exception) -> str:
    """将请求异常转换为简洁直观的简体中文诊断信息。"""
    if isinstance(exc, requests.exceptions.SSLError):
        return "站点 HTTPS / 源站证书连接失败"
    if isinstance(exc, requests.Timeout):
        return "请求源站接口超时"
    if isinstance(exc, requests.TooManyRedirects):
        return "站点存在循环重定向，请检查域名配置"
    if isinstance(exc, requests.ConnectionError):
        msg = str(exc).lower()
        if "name or service not known" in msg or "nameresolutionerror" in msg or "gaierror" in msg:
            return "站点域名 DNS 解析失败"
        return "无法连接到源站服务器"
    return f"网络请求异常（{type(exc).__name__}）"


def _same_store_alias(origin_url: str, redirect_url: str) -> str | None:
    """若重定向仅发生在同一 HTTPS 域名的裸域与 www 之间，返回修正后的目标 URL。"""
    origin = urlsplit(origin_url)
    target = urlsplit(urljoin(origin_url, redirect_url))
    if (
        origin.scheme == target.scheme == "https"
        and origin.hostname
        and target.hostname
        and origin.hostname != target.hostname
        and origin.hostname.removeprefix("www.") == target.hostname.removeprefix("www.")
        and origin.port == target.port
        and not target.username
        and not target.password
    ):
        return origin._replace(netloc=target.netloc).geturl()
    return None


class WooWebhookApiError(RuntimeError):
    """调用源站 WooCommerce REST API 失败时抛出的受控中文异常。"""

    def __init__(self, message: str, *, status_code: int | None = None):
        super().__init__(message)
        self.status_code = status_code


def _wc_rest_json(
    session: requests.Session,
    method: str,
    url: str,
    *,
    auth: tuple[str, str],
    params: dict[str, Any] | None = None,
    json_body: dict[str, Any] | None = None,
    timeout: tuple[float, float] = (WC_CONNECT_TIMEOUT, WC_READ_TIMEOUT),
) -> Any:
    """调用源站 WooCommerce REST API，保留 POST/PUT 方法与认证头穿过裸域/www 跳转。"""
    headers = {
        "Accept": "application/json",
        "User-Agent": "WooCommerce API Client-Python/3.0.0",
    }
    try:
        response = session.request(
            method=method.upper(),
            url=url,
            auth=auth,
            params=params,
            json=json_body,
            headers=headers,
            timeout=timeout,
            allow_redirects=False,
        )
        if response.status_code in {301, 302, 303, 307, 308}:
            location = response.headers.get("Location") or ""
            alias_url = _same_store_alias(url, location)
            if alias_url:
                response = session.request(
                    method=method.upper(),
                    url=alias_url,
                    auth=auth,
                    params=params,
                    json=json_body,
                    headers=headers,
                    timeout=timeout,
                    allow_redirects=False,
                )
            else:
                response = session.request(
                    method=method.upper(),
                    url=url,
                    auth=auth,
                    params=params,
                    json=json_body,
                    headers=headers,
                    timeout=timeout,
                    allow_redirects=True,
                )
    except requests.RequestException as exc:
        raise WooWebhookApiError(_friendly_http_error(exc)) from exc

    status = int(response.status_code)
    if status == 404:
        raise WooWebhookApiError("源站接口或目标资源不存在（HTTP 404）", status_code=404)
    if status in {401, 403}:
        raise WooWebhookApiError(f"源站 REST API 鉴权失败（HTTP {status}）", status_code=status)
    if status < 200 or status >= 300:
        raise WooWebhookApiError(f"源站接口返回异常状态（HTTP {status}）", status_code=status)
    if "text/html" in (response.headers.get("Content-Type") or "").lower():
        raise WooWebhookApiError("源站返回网页而非 JSON，可能被防火墙拦截或站点维护中", status_code=status)
    try:
        return response.json()
    except ValueError as exc:
        raise WooWebhookApiError("源站接口返回了无效的 JSON 数据", status_code=status) from exc


def _is_ping_request(req) -> bool:
    """判断是否为 WooCommerce 保存/创建 Webhook 时发送的连通性握手请求。"""
    event = (req.headers.get("X-WC-Webhook-Event") or "").strip().lower()
    topic = (req.headers.get("X-WC-Webhook-Topic") or "").strip().lower()
    if event == "ping" or topic == "action.woocommerce_webhook_payload":
        return True
    ctype = (req.headers.get("Content-Type") or "").lower()
    if "application/x-www-form-urlencoded" in ctype:
        raw_text = req.get_data(cache=True).decode("utf-8", errors="ignore").strip()
        parsed = parse_qs(raw_text, keep_blank_values=True)
        if "webhook_id" in parsed and len(parsed) <= 2:
            return True
    return False


def _redis_client():
    """获取 Celery 共用的 Redis 客户端用于毫秒级防抖排队锁。"""
    import redis
    from celery_app import BROKER_URL

    return redis.Redis.from_url(BROKER_URL, socket_connect_timeout=1.5, socket_timeout=1.5)


def dispatch_webhook_order_fetch(
    site_id: int,
    woo_order_id: int,
    *,
    delivery_row_id: int | None = None,
    db_path: str | None = None,
) -> str:
    """对同站点同订单触发防抖合并异步拉取；若队列瞬时异常则保留数据库 queued 待补偿。"""
    site_id = int(site_id)
    woo_order_id = int(woo_order_id)

    if not db_backend.is_postgres_backend():
        return "deferred_local_mode"

    broker = None
    try:
        broker = _redis_client()
        seq_key = f"woo-analysis:wh-seq:{site_id}:{woo_order_id}"
        lock_key = f"woo-analysis:wh-queued:{site_id}:{woo_order_id}"
        pipe = broker.pipeline()
        pipe.incr(seq_key)
        pipe.expire(seq_key, 600)
        pipe.execute()

        acquired = bool(broker.set(lock_key, "1", nx=True, ex=180))
        if not acquired:
            if delivery_row_id is not None:
                conn = get_db_conn(db_path)
                try:
                    conn.execute(
                        "UPDATE woo_webhook_deliveries SET status='coalesced' "
                        "WHERE id=? AND status='queued'",
                        (int(delivery_row_id),),
                    )
                    conn.commit()
                finally:
                    conn.close()
            return "coalesced"

        try:
            from sync_tasks import fetch_webhook_order

            fetch_webhook_order.apply_async(args=[site_id, woo_order_id])
            return "dispatched"
        except Exception:
            broker.delete(lock_key)
            raise
    except Exception as exc:
        logger.warning(
            "Webhook 异步任务投递暂缓，将由后台巡检自动补偿: site_id=%s order=%s err=%s",
            site_id,
            woo_order_id,
            type(exc).__name__,
        )
        return "deferred"
    finally:
        if broker is not None:
            try:
                broker.close()
            except Exception:
                pass


def fetch_single_order_from_woo(
    site_url: str,
    consumer_key: str,
    consumer_secret: str,
    site_id: int,
    woo_order_id: int,
    *,
    session: requests.Session | None = None,
) -> dict[str, Any]:
    """回源站拉取权威单笔订单及最近备注。"""
    clean_url = str(site_url or "").strip().rstrip("/")
    own_session = session is None
    sess = session or requests.Session()
    auth = (str(consumer_key), str(consumer_secret))
    try:
        order_url = urljoin(clean_url + "/", f"wp-json/wc/v3/orders/{int(woo_order_id)}")
        try:
            order = _wc_rest_json(sess, "GET", order_url, auth=auth)
        except WooWebhookApiError as exc:
            if exc.status_code == 404:
                return {"not_found": True, "order": None, "notes": []}
            raise
        if not isinstance(order, dict) or order.get("id") is None:
            raise WooWebhookApiError("源站返回的单笔订单数据结构无效")
        order["source"] = clean_url

        notes: list[dict[str, Any]] = []
        if str(order.get("status") or "") not in IGNORED_ORDER_STATUSES:
            notes_url = urljoin(
                clean_url + "/", f"wp-json/wc/v3/orders/{int(woo_order_id)}/notes"
            )
            try:
                raw_notes = _wc_rest_json(
                    sess, "GET", notes_url, auth=auth, params={"per_page": 20}
                )
                if isinstance(raw_notes, list):
                    local_oid = make_oid(site_id, order["id"])
                    for note in raw_notes:
                        if isinstance(note, dict):
                            item = dict(note)
                            item["_local_order_id"] = local_oid
                            notes.append(item)
            except Exception as note_exc:
                logger.warning(
                    "单笔订单备注拉取失败（不影响主订单入库）: site_id=%s order=%s err=%s",
                    site_id,
                    woo_order_id,
                    note_exc,
                )
        return {"not_found": False, "order": order, "notes": notes}
    finally:
        if own_session:
            sess.close()


def apply_webhook_order_write(
    site_id: int,
    woo_order_id: int,
    order: dict[str, Any] | None,
    notes: list[dict[str, Any]] | None = None,
    *,
    not_found: bool = False,
    db_path: str | None = None,
) -> dict[str, Any]:
    """在串行写事务中持久化单笔订单并更新 Webhook 投递记录，随后触发下游动作。"""
    site_id = int(site_id)
    woo_order_id = int(woo_order_id)
    conn = get_db_conn(db_path)
    planning_candidates: list[dict[str, Any]] = []
    now_utc = datetime.now(timezone.utc)
    oid = make_oid(site_id, woo_order_id)
    try:
        if is_site_archived(conn, site_id):
            if webhook_tables_exist(conn):
                conn.execute(
                    "UPDATE woo_webhook_deliveries "
                    "SET status='ignored', error_message='站点已归档', processed_at=CURRENT_TIMESTAMP "
                    "WHERE site_id=? AND woo_order_id=? AND status IN ('queued','coalesced','fetching','error')",
                    (site_id, woo_order_id),
                )
                conn.commit()
            return {"written": 0, "changed": 0, "ignored": "site_archived"}

        if not_found or not order:
            if webhook_tables_exist(conn):
                conn.execute(
                    "UPDATE woo_webhook_deliveries "
                    "SET status='ignored', error_message='源站订单不存在或已删除', processed_at=CURRENT_TIMESTAMP "
                    "WHERE site_id=? AND woo_order_id=? AND status IN ('queued','coalesced','fetching','error')",
                    (site_id, woo_order_id),
                )
                conn.commit()
            return {"written": 0, "changed": 0, "ignored": "order_not_found"}

        order_status = str(order.get("status") or "").strip()
        if order_status in IGNORED_ORDER_STATUSES:
            if webhook_tables_exist(conn):
                conn.execute(
                    "UPDATE woo_webhook_deliveries "
                    "SET status='ignored', order_status=?, processed_at=CURRENT_TIMESTAMP "
                    "WHERE site_id=? AND woo_order_id=? AND status IN ('queued','coalesced','fetching','error')",
                    (order_status, site_id, woo_order_id),
                )
                conn.commit()
            return {"written": 0, "changed": 0, "ignored": "draft_order"}

        write_res = upsert_orders_in_transaction([order], conn)
        if notes:
            upsert_order_notes_in_transaction(notes, conn)
        planning_candidates = write_res.get("planning_candidates") or []

        if _is_postgres(conn):
            clear_site_failure(conn, site_id)

        if webhook_tables_exist(conn):
            pending_rows = conn.execute(
                "SELECT id, received_at FROM woo_webhook_deliveries "
                "WHERE site_id=? AND woo_order_id=? "
                "AND status IN ('queued','coalesced','fetching','error')",
                (site_id, woo_order_id),
            ).fetchall()
            for row in pending_rows:
                recv_dt = _parse_utc_timestamp(row["received_at"])
                latency_ms = None
                if recv_dt is not None:
                    latency_ms = max(
                        0, int((now_utc - recv_dt).total_seconds() * 1000)
                    )
                conn.execute(
                    "UPDATE woo_webhook_deliveries "
                    "SET status='synced', order_oid=?, order_status=?, latency_ms=?, "
                    "    error_message=NULL, processed_at=CURRENT_TIMESTAMP "
                    "WHERE id=?",
                    (oid, order_status, latency_ms, row["id"]),
                )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    if (
        os.getenv("WOO_SYNC_POST_COMMIT_ACTIONS_ENABLED", "1") == "1"
        and planning_candidates
    ):
        run_post_commit_sync_actions(planning_candidates, strict=False)

    return {
        "written": int(write_res.get("written") or 0),
        "inserted": int(write_res.get("inserted") or 0),
        "changed": int(write_res.get("changed") or 0),
        "order_oid": oid,
        "order_status": order_status,
    }


def mark_webhook_delivery_error(
    site_id: int, woo_order_id: int, message: str, *, db_path: str | None = None
) -> None:
    """当单笔订单多次回读仍失败时记录错误信息（后续仍由每小时自动同步兜底补齐）。"""
    conn = get_db_conn(db_path)
    try:
        if not webhook_tables_exist(conn):
            return
        conn.execute(
            "UPDATE woo_webhook_deliveries "
            "SET status='error', error_message=?, processed_at=CURRENT_TIMESTAMP "
            "WHERE site_id=? AND woo_order_id=? AND status IN ('queued','coalesced','fetching')",
            (str(message or "")[:500], int(site_id), int(woo_order_id)),
        )
        conn.commit()
    except Exception:
        conn.rollback()
    finally:
        conn.close()


def recover_pending_webhook_deliveries(
    *, limit: int = 50, db_path: str | None = None
) -> int:
    """补偿恢复因队列瞬断滞留在 queued / fetching 状态的实时推送订单。"""
    if not db_backend.is_postgres_backend():
        return 0
    conn = get_db_conn(db_path)
    pairs: list[tuple[int, int]] = []
    try:
        if not webhook_tables_exist(conn):
            return 0
        rows = conn.execute(
            """
            SELECT DISTINCT site_id, woo_order_id
            FROM woo_webhook_deliveries
            WHERE woo_order_id IS NOT NULL
              AND (
                  (status='queued' AND received_at <= CURRENT_TIMESTAMP - interval '20 seconds')
                  OR (status='fetching' AND received_at <= CURRENT_TIMESTAMP - interval '180 seconds')
              )
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
        pairs = [(int(r["site_id"]), int(r["woo_order_id"])) for r in rows]
    finally:
        conn.close()

    recovered = 0
    for s_id, w_id in pairs:
        res = dispatch_webhook_order_fetch(s_id, w_id, db_path=db_path)
        if res == "dispatched":
            recovered += 1
    return recovered


def prune_old_webhook_deliveries(
    *, retention_days: int = 14, db_path: str | None = None
) -> int:
    """定期清理超过保留天数的已完成投递记录，保持收件箱表轻量高效。"""
    conn = get_db_conn(db_path)
    try:
        if not webhook_tables_exist(conn):
            return 0
        days = max(1, int(retention_days))
        if _is_postgres(conn):
            cur = conn.execute(
                "DELETE FROM woo_webhook_deliveries "
                "WHERE status IN ('ping','synced','ignored','coalesced','error') "
                "AND received_at < CURRENT_TIMESTAMP - (? * interval '1 day')",
                (days,),
            )
        else:
            cur = conn.execute(
                "DELETE FROM woo_webhook_deliveries "
                "WHERE status IN ('ping','synced','ignored','coalesced','error') "
                "AND received_at < datetime('now', ?)",
                (f"-{days} days",),
            )
        deleted = max(0, int(cur.rowcount or 0))
        conn.commit()
        return deleted
    except Exception:
        conn.rollback()
        return 0
    finally:
        conn.close()


def _ensure_local_site_secrets(
    conn,
    site_id: int,
    delivery_url: str,
    *,
    rotate_secret: bool = False,
) -> tuple[str, dict[str, dict[str, Any]]]:
    """确保站点本地数据库已提前持久化统一的 HMAC 密钥（防止远端创建时同步发 Ping 查不到密钥）。"""
    rows = conn.execute(
        "SELECT * FROM woo_site_webhooks WHERE site_id=?", (int(site_id),)
    ).fetchall()
    existing = {str(r["topic"]): dict(r) for r in rows}

    current_secret = None
    for topic in MANAGED_TOPICS:
        cand = str((existing.get(topic) or {}).get("secret") or "").strip()
        if len(cand) >= 32:
            current_secret = cand
            break

    if not current_secret or rotate_secret:
        new_secret = secrets.token_hex(32)
        prev_secret = current_secret
        secret_to_use = new_secret
    else:
        secret_to_use = current_secret
        prev_secret = None

    for topic in MANAGED_TOPICS:
        row = existing.get(topic)
        if not row:
            conn.execute(
                """
                INSERT INTO woo_site_webhooks
                    (site_id, topic, secret, previous_secret, delivery_url, remote_status)
                VALUES (?, ?, ?, ?, ?, 'unknown')
                ON CONFLICT(site_id, topic) DO UPDATE SET
                    secret=excluded.secret,
                    delivery_url=excluded.delivery_url,
                    updated_at=CURRENT_TIMESTAMP
                """,
                (int(site_id), topic, secret_to_use, prev_secret, delivery_url),
            )
        elif rotate_secret or row.get("secret") != secret_to_use or row.get("delivery_url") != delivery_url:
            conn.execute(
                """
                UPDATE woo_site_webhooks
                SET previous_secret=?,
                    secret=?,
                    delivery_url=?,
                    updated_at=CURRENT_TIMESTAMP
                WHERE site_id=? AND topic=?
                """,
                (
                    row.get("secret") if rotate_secret else row.get("previous_secret"),
                    secret_to_use,
                    delivery_url,
                    int(site_id),
                    topic,
                ),
            )
    conn.commit()
    refreshed = conn.execute(
        "SELECT * FROM woo_site_webhooks WHERE site_id=?", (int(site_id),)
    ).fetchall()
    return secret_to_use, {str(r["topic"]): dict(r) for r in refreshed}


def _recently_verified_healthy(
    local_map: dict[str, dict[str, Any]], expected_url: str
) -> bool:
    """判断该站点的两个受管 Topic 是否近期刚核对为正常激活状态（避免频繁巡检消耗源站资源）。"""
    now_utc = datetime.now(timezone.utc)
    for topic in MANAGED_TOPICS:
        row = local_map.get(topic)
        if not row:
            return False
        if str(row.get("remote_status") or "") != "active":
            return False
        if not row.get("remote_webhook_id"):
            return False
        if str(row.get("delivery_url") or "").rstrip("/") != expected_url.rstrip("/"):
            return False
        synced_dt = _parse_utc_timestamp(row.get("last_synced_at"))
        if not synced_dt or (now_utc - synced_dt).total_seconds() > HEALTHY_REVERIFY_SECONDS:
            return False
    return True


def provision_site_webhooks(
    site_id: int,
    *,
    force_verify: bool = True,
    rotate_secret: bool = False,
    base_url: str | None = None,
    session: requests.Session | None = None,
    db_path: str | None = None,
) -> dict[str, Any]:
    """为单个活跃 WooCommerce 站点创建、修复或核对 `order.created` 与 `order.updated` Webhook。"""
    site_id = int(site_id)
    delivery_url = build_delivery_url(site_id, base_url=base_url)

    conn = get_db_conn(db_path)
    try:
        if not webhook_tables_exist(conn):
            return {"site_id": site_id, "ok": False, "error": "Webhook 数据表尚未迁移"}
        site = conn.execute("SELECT * FROM sites WHERE id=?", (site_id,)).fetchone()
        if not site:
            return {"site_id": site_id, "ok": False, "error": "站点不存在"}
        if is_site_archived(site):
            conn.execute(
                "UPDATE woo_site_webhooks SET remote_status='archived', updated_at=CURRENT_TIMESTAMP "
                "WHERE site_id=?",
                (site_id,),
            )
            conn.commit()
            return {"site_id": site_id, "ok": True, "skipped": "archived", "url": site["url"]}

        site_url = str(site["url"] or "").strip().rstrip("/")
        ck = str(site["consumer_key"] or "").strip()
        cs = str(site["consumer_secret"] or "").strip()
        if not site_url or not ck or not cs:
            conn.execute(
                "UPDATE woo_site_webhooks SET remote_status='error', last_error='站点缺少 API 密钥配置', "
                "last_synced_at=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP WHERE site_id=?",
                (site_id,),
            )
            conn.commit()
            return {"site_id": site_id, "url": site_url, "ok": False, "error": "站点缺少 API 密钥配置"}

        secret, local_map = _ensure_local_site_secrets(
            conn, site_id, delivery_url, rotate_secret=rotate_secret
        )
        if not force_verify and not rotate_secret and _recently_verified_healthy(local_map, delivery_url):
            return {
                "site_id": site_id,
                "url": site_url,
                "ok": True,
                "cached": True,
                "topics": {t: "active" for t in MANAGED_TOPICS},
            }
    finally:
        conn.close()

    own_session = session is None
    sess = session or requests.Session()
    auth = (ck, cs)
    webhooks_endpoint = urljoin(site_url + "/", "wp-json/wc/v3/webhooks")
    topic_results: dict[str, Any] = {}
    try:
        remote_list = _wc_rest_json(
            sess,
            "GET",
            webhooks_endpoint,
            auth=auth,
            params={"per_page": 100, "status": "all"},
        )
        if not isinstance(remote_list, list):
            raise WooWebhookApiError("源站返回的 Webhook 列表格式无效")

        expected_norm = delivery_url.rstrip("/")
        for topic in MANAGED_TOPICS:
            local_row = local_map.get(topic) or {}
            local_remote_id = local_row.get("remote_webhook_id")
            matching = [
                item
                for item in remote_list
                if isinstance(item, dict)
                and str(item.get("topic") or "").strip() == topic
                and (
                    str(item.get("delivery_url") or "").strip().rstrip("/") == expected_norm
                    or (
                        str(item.get("name") or "").startswith(WEBHOOK_NAME_PREFIX)
                        and "/api/webhooks/woocommerce/orders/" in str(item.get("delivery_url") or "")
                    )
                )
            ]

            primary = None
            if local_remote_id:
                for item in matching:
                    if int(item.get("id") or 0) == int(local_remote_id):
                        primary = item
                        break
            if primary is None and matching:
                primary = matching[0]

            # 清理同一站点同一事件的重复受管项
            for item in matching:
                if primary is not None and item.get("id") != primary.get("id"):
                    dup_id = item.get("id")
                    if dup_id:
                        try:
                            _wc_rest_json(
                                sess,
                                "DELETE",
                                f"{webhooks_endpoint}/{int(dup_id)}",
                                auth=auth,
                                params={"force": "true"},
                            )
                        except Exception:
                            pass

            if primary is None:
                created = _wc_rest_json(
                    sess,
                    "POST",
                    webhooks_endpoint,
                    auth=auth,
                    json_body={
                        "name": TOPIC_LABELS[topic],
                        "topic": topic,
                        "delivery_url": delivery_url,
                        "secret": secret,
                        "status": "active",
                        "api_version": "wp_api_v3",
                    },
                )
                remote_id = int(created.get("id"))
                action = "created"
            else:
                remote_id = int(primary.get("id"))
                remote_status = str(primary.get("status") or "").strip().lower()
                remote_deliv = str(primary.get("delivery_url") or "").strip().rstrip("/")
                needs_update = (
                    rotate_secret
                    or remote_status != "active"
                    or remote_deliv != expected_norm
                    or str(local_row.get("remote_status") or "") != "active"
                    or int(local_row.get("remote_webhook_id") or 0) != remote_id
                )
                if needs_update:
                    _wc_rest_json(
                        sess,
                        "PUT",
                        f"{webhooks_endpoint}/{remote_id}",
                        auth=auth,
                        json_body={
                            "name": TOPIC_LABELS[topic],
                            "topic": topic,
                            "delivery_url": delivery_url,
                            "secret": secret,
                            "status": "active",
                        },
                    )
                    action = "updated"
                else:
                    action = "verified"

            topic_results[topic] = {"remote_id": remote_id, "action": action, "status": "active"}

        conn = get_db_conn(db_path)
        try:
            for topic, info in topic_results.items():
                conn.execute(
                    """
                    UPDATE woo_site_webhooks
                    SET remote_webhook_id=?,
                        delivery_url=?,
                        remote_status='active',
                        last_synced_at=CURRENT_TIMESTAMP,
                        last_error=NULL,
                        updated_at=CURRENT_TIMESTAMP
                    WHERE site_id=? AND topic=?
                    """,
                    (info["remote_id"], delivery_url, site_id, topic),
                )
            conn.commit()
        finally:
            conn.close()

        return {
            "site_id": site_id,
            "url": site_url,
            "ok": True,
            "topics": topic_results,
        }
    except Exception as exc:
        err_msg = str(exc) if isinstance(exc, WooWebhookApiError) else _friendly_http_error(exc)
        conn = get_db_conn(db_path)
        try:
            conn.execute(
                """
                UPDATE woo_site_webhooks
                SET remote_status='error',
                    last_error=?,
                    last_synced_at=CURRENT_TIMESTAMP,
                    updated_at=CURRENT_TIMESTAMP
                WHERE site_id=?
                """,
                (err_msg[:500], site_id),
            )
            conn.commit()
        finally:
            conn.close()
        return {
            "site_id": site_id,
            "url": site_url,
            "ok": False,
            "error": err_msg,
        }
    finally:
        if own_session:
            sess.close()


def provision_all_active_sites(
    *,
    site_ids: Iterable[int] | None = None,
    force_verify: bool = True,
    max_workers: int = 8,
    base_url: str | None = None,
    db_path: str | None = None,
) -> dict[str, Any]:
    """并发为所有活跃站点（或指定站点列表）批量部署与巡检修复 Webhook。"""
    conn = get_db_conn(db_path)
    try:
        if not webhook_tables_exist(conn):
            return {"ok": False, "error": "Webhook 数据表尚未迁移", "total": 0, "succeeded": 0, "failed": 0}
        if site_ids is not None:
            clean_ids = sorted({int(x) for x in site_ids if int(x) > 0})
            if not clean_ids:
                return {"ok": True, "total": 0, "succeeded": 0, "failed": 0, "results": []}
            marks = ",".join("?" for _ in clean_ids)
            rows = conn.execute(
                f"SELECT id FROM sites WHERE id IN ({marks}) "
                "AND COALESCE(api_status, '') <> ? ORDER BY id",
                (*clean_ids, ARCHIVED_STATUS),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT id FROM sites WHERE COALESCE(api_status, '') <> ? ORDER BY id",
                (ARCHIVED_STATUS,),
            ).fetchall()
        target_ids = [int(r["id"]) for r in rows]
    finally:
        conn.close()

    if not target_ids:
        return {"ok": True, "total": 0, "succeeded": 0, "failed": 0, "results": []}

    workers = max(1, min(int(max_workers or 8), len(target_ids)))
    results: list[dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_map = {
            pool.submit(
                provision_site_webhooks,
                s_id,
                force_verify=force_verify,
                base_url=base_url,
                db_path=db_path,
            ): s_id
            for s_id in target_ids
        }
        for fut in as_completed(future_map):
            s_id = future_map[fut]
            try:
                results.append(fut.result())
            except Exception as exc:
                results.append({"site_id": s_id, "ok": False, "error": str(exc)[:300]})

    results.sort(key=lambda item: int(item.get("site_id") or 0))
    succeeded = sum(1 for item in results if item.get("ok"))
    failed = len(results) - succeeded
    return {
        "ok": failed == 0,
        "total": len(results),
        "succeeded": succeeded,
        "failed": failed,
        "results": results,
    }


def cleanup_remote_site_webhooks_best_effort(
    site_url: str,
    consumer_key: str,
    consumer_secret: str,
    remote_ids: Iterable[int],
) -> None:
    """在移除站点连接后，短超时尽力清理源站上的受管 Webhook（不阻塞主业务）。"""
    clean_url = str(site_url or "").strip().rstrip("/")
    ck = str(consumer_key or "").strip()
    cs = str(consumer_secret or "").strip()
    ids = [int(x) for x in (remote_ids or []) if x]
    if not clean_url or not ck or not cs or not ids:
        return
    sess = requests.Session()
    try:
        for wid in ids:
            try:
                _wc_rest_json(
                    sess,
                    "DELETE",
                    urljoin(clean_url + "/", f"wp-json/wc/v3/webhooks/{wid}"),
                    auth=(ck, cs),
                    params={"force": "true"},
                    timeout=(2.5, 4.0),
                )
            except Exception:
                pass
    finally:
        sess.close()


def _user_can_manage_all(user) -> bool:
    """判断登录用户是否具备全站同步管理权限。"""
    return bool(
        getattr(user, "is_authenticated", False)
        and user.is_admin()
        and user.can_manage_users()
    )


def _allowed_site_ids_for_user(conn, user) -> set[int] | None:
    """返回当前登录用户可管理的站点 ID 集合；若为全局超管则返回 None 表示全部。"""
    if _user_can_manage_all(user):
        return None
    if not getattr(user, "is_authenticated", False):
        return set()
    can_own = getattr(user, "can_manage_own_site_sync", None)
    if not callable(can_own) or not can_own():
        return set()
    rows = conn.execute(
        """
        SELECT s.id
        FROM sites s
        JOIN users u ON u.id = ?
        WHERE COALESCE(s.api_status, '') <> ?
          AND TRIM(COALESCE(u.name, '')) <> ''
          AND TRIM(COALESCE(s.manager, '')) = TRIM(u.name)
        """,
        (user.id, ARCHIVED_STATUS),
    ).fetchall()
    return {int(r[0]) for r in rows}


@woo_order_webhook_bp.route(
    "/api/webhooks/woocommerce/orders/<int:site_id>", methods=["POST"]
)
def receive_woocommerce_order_webhook(site_id: int):
    """接收源站 WooCommerce 推送的订单新建/更新事件（HMAC 验签 + 毫秒级入队返回）。"""
    if request.content_length is not None and request.content_length > MAX_WEBHOOK_BYTES:
        return jsonify({"ok": False, "error": "payload_too_large"}), 413

    raw = request.get_data(cache=True)
    if len(raw) > MAX_WEBHOOK_BYTES:
        return jsonify({"ok": False, "error": "payload_too_large"}), 413

    conn = get_db_conn()
    try:
        if not webhook_tables_exist(conn):
            return jsonify({"ok": False, "error": "webhook_schema_not_ready"}), 503

        site = conn.execute(
            "SELECT id, url, api_status FROM sites WHERE id=?", (int(site_id),)
        ).fetchone()
        if not site:
            return jsonify({"ok": False, "error": "site_not_found"}), 404

        if _is_ping_request(request):
            if not is_site_archived(site):
                conn.execute(
                    "UPDATE woo_site_webhooks "
                    "SET last_ping_at=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP "
                    "WHERE site_id=?",
                    (int(site_id),),
                )
                conn.commit()
            return jsonify({"ok": True, "ping": True}), 200

        wh_rows = conn.execute(
            "SELECT topic, secret, previous_secret FROM woo_site_webhooks WHERE site_id=?",
            (int(site_id),),
        ).fetchall()
        if not wh_rows:
            return jsonify({"ok": False, "error": "webhook_not_configured"}), 401

        candidate_secrets: list[str] = []
        for r in wh_rows:
            for key in ("secret", "previous_secret"):
                val = str(r[key] or "").strip()
                if val and val not in candidate_secrets:
                    candidate_secrets.append(val)

        sig_header = request.headers.get("X-WC-Webhook-Signature")
        if not verify_wc_signature(raw, sig_header, candidate_secrets):
            return jsonify({"ok": False, "error": "invalid_signature"}), 401

        if is_site_archived(site):
            return jsonify({"ok": True, "ignored": "site_archived"}), 200

        if not is_webhook_sync_enabled(conn):
            return jsonify({"ok": True, "ignored": "webhook_sync_disabled"}), 200

        topic = (request.headers.get("X-WC-Webhook-Topic") or "").strip().lower()
        event = (request.headers.get("X-WC-Webhook-Event") or "").strip().lower()
        delivery_id = (request.headers.get("X-WC-Webhook-Delivery-ID") or "").strip()[:120]

        if topic and topic not in MANAGED_TOPICS:
            return jsonify({"ok": True, "ignored": "unsupported_topic"}), 200
        effective_topic = topic or "order.updated"

        try:
            payload = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, ValueError):
            return jsonify({"ok": False, "error": "invalid_json"}), 400
        if not isinstance(payload, dict) or payload.get("id") is None:
            return jsonify({"ok": True, "ignored": "missing_order_id"}), 200

        try:
            woo_order_id = int(payload["id"])
        except (TypeError, ValueError):
            return jsonify({"ok": False, "error": "invalid_order_id"}), 400
        if woo_order_id <= 0:
            return jsonify({"ok": False, "error": "invalid_order_id"}), 400

        payload_hash = hashlib.sha256(raw).hexdigest()
        if delivery_id:
            delivery_key = f"{site_id}:{effective_topic}:{delivery_id}:{payload_hash[:16]}"
        else:
            delivery_key = f"{site_id}:{effective_topic}:{woo_order_id}:{payload_hash[:24]}"

        raw_status = str(payload.get("status") or "").strip().lower()
        initial_status = "ignored" if raw_status in IGNORED_ORDER_STATUSES else "queued"

        cur = conn.execute(
            """
            INSERT INTO woo_webhook_deliveries
                (delivery_key, site_id, woo_order_id, topic, event, delivery_id,
                 payload_hash, status, order_status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(delivery_key) DO NOTHING
            """,
            (
                delivery_key,
                int(site_id),
                woo_order_id,
                effective_topic,
                event or None,
                delivery_id or None,
                payload_hash,
                initial_status,
                raw_status or None,
            ),
        )
        inserted = bool(cur.rowcount and int(cur.rowcount) > 0)
        conn.execute(
            """
            UPDATE woo_site_webhooks
            SET last_delivery_at=CURRENT_TIMESTAMP,
                remote_status='active',
                last_error=NULL,
                updated_at=CURRENT_TIMESTAMP
            WHERE site_id=? AND topic=?
            """,
            (int(site_id), effective_topic),
        )
        delivery_row_id = None
        if inserted and initial_status == "queued":
            id_row = conn.execute(
                "SELECT id FROM woo_webhook_deliveries WHERE delivery_key=?",
                (delivery_key,),
            ).fetchone()
            delivery_row_id = int(id_row["id"]) if id_row else None
        conn.commit()
    finally:
        conn.close()

    if not inserted:
        return jsonify({"ok": True, "duplicate": True}), 200
    if initial_status == "ignored":
        return jsonify({"ok": True, "ignored": "draft_order"}), 200

    dispatch_state = dispatch_webhook_order_fetch(
        site_id, woo_order_id, delivery_row_id=delivery_row_id
    )
    return (
        jsonify(
            {
                "ok": True,
                "queued": True,
                "dispatch": dispatch_state,
                "site_id": int(site_id),
                "woo_order_id": woo_order_id,
            }
        ),
        200,
    )


@woo_order_webhook_bp.route("/api/settings/webhooks/status", methods=["GET"])
@login_required
def get_webhooks_dashboard_status():
    """查询实时 Webhook 同步总览、各站点健康状态与最近实时推送流水。"""
    conn = get_db_conn()
    try:
        allowed_ids = _allowed_site_ids_for_user(conn, current_user)
        if allowed_ids is not None and not allowed_ids:
            return jsonify({"error": "无权查看同步设置"}), 403

        if not webhook_tables_exist(conn):
            return jsonify({"ready": False, "enabled": False, "sites": [], "recent_deliveries": []})

        enabled = is_webhook_sync_enabled(conn)
        auto_en_row = conn.execute(
            "SELECT value FROM settings WHERE key='autosync_enabled'"
        ).fetchone()
        auto_iv_row = conn.execute(
            "SELECT value FROM settings WHERE key='autosync_interval'"
        ).fetchone()
        autosync_enabled = bool(
            auto_en_row and str(auto_en_row[0]).strip().lower() == "true"
        )
        try:
            autosync_interval = int(auto_iv_row[0]) if auto_iv_row else 3600
        except (TypeError, ValueError):
            autosync_interval = 3600

        site_rows = conn.execute(
            "SELECT id, url, manager, country, COALESCE(api_status, 'ok') AS api_status "
            "FROM sites WHERE COALESCE(api_status, '') <> ? ORDER BY id",
            (ARCHIVED_STATUS,),
        ).fetchall()
        if allowed_ids is not None:
            site_rows = [r for r in site_rows if int(r["id"]) in allowed_ids]

        wh_rows = conn.execute(
            """
            SELECT site_id, topic, remote_webhook_id, delivery_url, remote_status,
                   last_synced_at, last_ping_at, last_delivery_at, last_error
            FROM woo_site_webhooks
            """
        ).fetchall()
        wh_by_site: dict[int, dict[str, dict[str, Any]]] = {}
        for row in wh_rows:
            wh_by_site.setdefault(int(row["site_id"]), {})[str(row["topic"])] = dict(row)

        sites_payload = []
        healthy_count = 0
        error_count = 0
        missing_count = 0

        for s in site_rows:
            sid = int(s["id"])
            tmap = wh_by_site.get(sid, {})
            created_row = tmap.get("order.created") or {}
            updated_row = tmap.get("order.updated") or {}
            c_status = str(created_row.get("remote_status") or "missing")
            u_status = str(updated_row.get("remote_status") or "missing")
            if c_status == "active" and u_status == "active":
                overall = "active"
                healthy_count += 1
            elif "error" in {c_status, u_status} or "disabled" in {c_status, u_status}:
                overall = "error"
                error_count += 1
            else:
                overall = "missing"
                missing_count += 1

            last_delivery = max(
                str(created_row.get("last_delivery_at") or ""),
                str(updated_row.get("last_delivery_at") or ""),
            ) or None
            last_synced = max(
                str(created_row.get("last_synced_at") or ""),
                str(updated_row.get("last_synced_at") or ""),
            ) or None
            last_err = created_row.get("last_error") or updated_row.get("last_error")

            sites_payload.append(
                {
                    "site_id": sid,
                    "url": s["url"],
                    "manager": s["manager"] or "",
                    "country": s["country"] or "",
                    "overall_status": overall,
                    "created_status": c_status,
                    "updated_status": u_status,
                    "created_remote_id": created_row.get("remote_webhook_id"),
                    "updated_remote_id": updated_row.get("remote_webhook_id"),
                    "last_synced_at": last_synced,
                    "last_delivery_at": last_delivery,
                    "last_error": last_err,
                }
            )

        if _is_postgres(conn):
            cond_1h = "received_at >= CURRENT_TIMESTAMP - interval '1 hour'"
            cond_24h = "received_at >= CURRENT_TIMESTAMP - interval '24 hours'"
        else:
            cond_1h = "received_at >= datetime('now', '-1 hour')"
            cond_24h = "received_at >= datetime('now', '-24 hours')"

        scope_sql = ""
        scope_params: tuple[Any, ...] = ()
        if allowed_ids is not None:
            marks = ",".join("?" for _ in allowed_ids)
            scope_sql = f" AND d.site_id IN ({marks})"
            scope_params = tuple(sorted(allowed_ids))

        stats_1h = conn.execute(
            f"""
            SELECT COUNT(*) AS total,
                   SUM(CASE WHEN d.status='synced' THEN 1 ELSE 0 END) AS synced,
                   COUNT(DISTINCT CASE WHEN d.status='synced' THEN d.order_oid END) AS distinct_orders
            FROM woo_webhook_deliveries d
            WHERE {cond_1h}{scope_sql}
            """,
            scope_params,
        ).fetchone()
        stats_24h = conn.execute(
            f"""
            SELECT COUNT(*) AS total,
                   SUM(CASE WHEN d.status='synced' THEN 1 ELSE 0 END) AS synced,
                   COUNT(DISTINCT CASE WHEN d.status='synced' THEN d.order_oid END) AS distinct_orders,
                   AVG(CASE WHEN d.status='synced' AND d.latency_ms IS NOT NULL THEN d.latency_ms END) AS avg_ms
            FROM woo_webhook_deliveries d
            WHERE {cond_24h}{scope_sql}
            """,
            scope_params,
        ).fetchone()

        recent_rows = conn.execute(
            f"""
            SELECT d.id, d.site_id, s.url AS site_url, d.woo_order_id, d.order_oid,
                   d.topic, d.status, d.order_status, d.latency_ms,
                   d.error_message, d.received_at, d.processed_at
            FROM woo_webhook_deliveries d
            LEFT JOIN sites s ON s.id = d.site_id
            WHERE 1=1{scope_sql}
            ORDER BY d.id DESC
            LIMIT 25
            """,
            scope_params,
        ).fetchall()

        avg_ms_raw = stats_24h["avg_ms"] if stats_24h else None
        avg_latency_ms = int(round(float(avg_ms_raw))) if avg_ms_raw is not None else None

        return jsonify(
            {
                "ready": True,
                "enabled": enabled,
                "can_manage_all": allowed_ids is None,
                "public_base_url": public_base_url(),
                "autosync_fallback": {
                    "enabled": autosync_enabled,
                    "interval": autosync_interval,
                },
                "summary": {
                    "total_active_sites": len(site_rows),
                    "healthy_sites": healthy_count,
                    "error_sites": error_count,
                    "missing_sites": missing_count,
                    "deliveries_1h": int(stats_1h["total"] or 0) if stats_1h else 0,
                    "synced_1h": int(stats_1h["synced"] or 0) if stats_1h else 0,
                    "distinct_orders_1h": int(stats_1h["distinct_orders"] or 0) if stats_1h else 0,
                    "deliveries_24h": int(stats_24h["total"] or 0) if stats_24h else 0,
                    "synced_24h": int(stats_24h["synced"] or 0) if stats_24h else 0,
                    "distinct_orders_24h": int(stats_24h["distinct_orders"] or 0) if stats_24h else 0,
                    "avg_latency_ms_24h": avg_latency_ms,
                },
                "sites": sites_payload,
                "recent_deliveries": [dict(r) for r in recent_rows],
            }
        )
    finally:
        conn.close()


@woo_order_webhook_bp.route("/api/settings/webhooks/toggle", methods=["POST"])
@login_required
def toggle_webhook_sync():
    """切换全局实时订单 Webhook 同步开关（仅全局同步管理员可操作）。"""
    if not _user_can_manage_all(current_user):
        return jsonify({"error": "无权修改全站实时同步开关"}), 403
    data = request.get_json(silent=True) or {}
    if "enabled" not in data:
        return jsonify({"error": "缺少 enabled 参数"}), 400
    conn = get_db_conn()
    try:
        if not webhook_tables_exist(conn):
            return jsonify({"error": "Webhook 数据表尚未迁移"}), 503
        new_state = set_webhook_sync_enabled(conn, bool(data.get("enabled")))
        return jsonify({"success": True, "enabled": new_state})
    finally:
        conn.close()


@woo_order_webhook_bp.route("/api/settings/webhooks/provision", methods=["POST"])
@login_required
def trigger_webhook_provision():
    """触发单站同步修复或全站一键部署巡检受管 Webhook。"""
    data = request.get_json(silent=True) or {}
    raw_site_id = data.get("site_id")
    rotate_secret = bool(data.get("rotate_secret", False))

    conn = get_db_conn()
    try:
        if not webhook_tables_exist(conn):
            return jsonify({"error": "Webhook 数据表尚未迁移"}), 503
        allowed_ids = _allowed_site_ids_for_user(conn, current_user)
    finally:
        conn.close()

    if raw_site_id is not None and str(raw_site_id).strip() != "":
        try:
            target_site_id = int(raw_site_id)
        except (TypeError, ValueError):
            return jsonify({"error": "无效的站点 ID"}), 400
        if allowed_ids is not None and target_site_id not in allowed_ids:
            return jsonify({"error": "无权操作该站点的 Webhook"}), 403
        res = provision_site_webhooks(
            target_site_id, force_verify=True, rotate_secret=rotate_secret
        )
        return jsonify({"success": bool(res.get("ok")), "mode": "single", "result": res}), (
            200 if res.get("ok") else 502
        )

    if allowed_ids is not None:
        return jsonify({"error": "仅超级管理员可执行全站批量部署"}), 403

    if db_backend.is_postgres_backend():
        try:
            from sync_tasks import reconcile_webhooks

            reconcile_webhooks.apply_async(kwargs={"force_verify": True})
            return jsonify(
                {
                    "success": True,
                    "mode": "async_all",
                    "message": "已提交全站 Webhook 部署与巡检任务，正在后台并发配置全部站点…",
                }
            ), 202
        except Exception as exc:
            current_app.logger.exception("投递全站 Webhook 部署任务失败: %s", exc)
            return jsonify({"error": "无法提交后台部署任务，请稍后重试"}), 503

    summary = provision_all_active_sites(force_verify=True)
    return jsonify({"success": True, "mode": "sync_all", "summary": summary})
