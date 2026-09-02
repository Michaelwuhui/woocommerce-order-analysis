"""Celery tasks for the durable WooCommerce synchronization pipeline."""

from __future__ import annotations

import hashlib
import json
import os
import random
import socket
from datetime import datetime
from typing import Any
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

import requests
from urllib3.util import connection as urllib3_connection

from celery_app import celery_app
from oid_utils import make_oid
from sync_service import (
    _event,
    _json,
    _refresh_run_completion,
    enqueue_fetch_page,
    enqueue_outbox,
    get_connection,
    mark_site_error,
    note_retry,
    publish_pending_outbox,
    recover_stale_work,
    start_sync,
)
from sync_utils import (
    _subtract_minutes_from_iso,
    run_post_commit_sync_actions,
    upsert_order_notes_in_transaction,
    upsert_orders_in_transaction,
)


CONNECT_TIMEOUT_SECONDS = float(os.getenv("WOO_SYNC_CONNECT_TIMEOUT", "5"))
READ_TIMEOUT_SECONDS = float(os.getenv("WOO_SYNC_READ_TIMEOUT", "30"))
MAX_FETCH_RETRIES = 3
MAX_WRITE_RETRIES = 3
RETRYABLE_HTTP = {429, 500, 502, 503, 504}
LOCAL_TIMEZONE = ZoneInfo("Asia/Hong_Kong")


def configure_ipv4_preference() -> bool:
    """Restrict urllib3 DNS selection to IPv4 in the dedicated fetch worker."""
    if os.getenv("WOO_SYNC_IPV4_ONLY", "0") != "1":
        return False
    urllib3_connection.allowed_gai_family = lambda: socket.AF_INET
    return urllib3_connection.allowed_gai_family() == socket.AF_INET


IPV4_PREFERENCE_ACTIVE = configure_ipv4_preference()


class TransientFetchError(RuntimeError):
    def __init__(self, message: str, retry_after: int | None = None):
        super().__init__(message)
        self.retry_after = retry_after


class AuthenticationFetchError(RuntimeError):
    pass


class PermanentFetchError(RuntimeError):
    pass


def _defer_busy_site(run_id: str, site_id: int, page: int) -> None:
    """Return a redundant same-site delivery to the durable outbox.

    Advisory-lock contention is coordination, not a WooCommerce failure, so it
    must not consume one of the three network retry attempts or increment the
    site's retry counter.
    """
    connection = get_connection()
    try:
        connection.execute(
            """
            UPDATE sync_task_outbox
            SET status='pending',available_at=CURRENT_TIMESTAMP + interval '15 seconds',
                updated_at=CURRENT_TIMESTAMP,last_error=NULL
            WHERE dedupe_key=?
              AND status IN ('published','publishing','error')
            """,
            (f"fetch:{run_id}:{site_id}:{page}",),
        )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _canonical_hash(orders: list[dict], notes: list[dict]) -> str:
    body = json.dumps(
        {"orders": orders, "notes": notes},
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(body).hexdigest()


def _retry_delay(retries: int, retry_after: int | None = None) -> int:
    if retry_after is not None:
        return min(120, max(1, int(retry_after)))
    return min(120, int((2 ** max(0, retries)) + random.uniform(0, 2)))


def _retry_after(response) -> int | None:
    value = response.headers.get("Retry-After")
    try:
        return int(value) if value is not None else None
    except (TypeError, ValueError):
        return None


def _http_json_list(session, url: str, *, auth, params=None):
    try:
        response = session.get(
            url,
            auth=auth,
            params=params,
            timeout=(CONNECT_TIMEOUT_SECONDS, READ_TIMEOUT_SECONDS),
            headers={
                "Accept": "application/json",
                "User-Agent": "WooCommerce API Client-Python/3.0.0",
            },
        )
    except (requests.Timeout, requests.ConnectionError) as exc:
        raise TransientFetchError(type(exc).__name__) from exc
    status = int(response.status_code)
    if status in {401, 403}:
        raise AuthenticationFetchError(f"WooCommerce authentication failed (HTTP {status})")
    if status in RETRYABLE_HTTP or 500 <= status <= 599:
        raise TransientFetchError(
            f"WooCommerce temporary failure (HTTP {status})",
            retry_after=_retry_after(response),
        )
    if status < 200 or status >= 300:
        raise PermanentFetchError(f"WooCommerce request failed (HTTP {status})")
    try:
        data = response.json()
    except ValueError as exc:
        raise TransientFetchError("WooCommerce returned invalid JSON") from exc
    if not isinstance(data, list):
        raise PermanentFetchError("WooCommerce returned a non-list response")
    return response, data


def _site_lock(connection, site_id: int) -> bool:
    row = connection.execute(
        "SELECT pg_try_advisory_lock(hashtextextended(?,0))",
        (f"woo-sync-site:{int(site_id)}",),
    ).fetchone()
    connection.commit()
    return bool(row and row[0])


def _site_unlock(connection, site_id: int) -> None:
    try:
        connection.execute(
            "SELECT pg_advisory_unlock(hashtextextended(?,0))",
            (f"woo-sync-site:{int(site_id)}",),
        )
        connection.commit()
    except Exception:
        connection.rollback()


def _cancel_site(connection, run_id: str, site_id: int, page: int) -> None:
    connection.execute(
        """
        UPDATE sync_page_dispatches
        SET status='cancelled',heartbeat_at=CURRENT_TIMESTAMP,
            updated_at=CURRENT_TIMESTAMP
        WHERE run_id=? AND site_id=? AND page=?
        """,
        (run_id, site_id, page),
    )
    connection.execute(
        """
        UPDATE sync_site_progress
        SET status='cancelled',finished_at=CURRENT_TIMESTAMP,
            heartbeat_at=CURRENT_TIMESTAMP,version=version+1
        WHERE run_id=? AND site_id=?
          AND status NOT IN ('success','error','auth_error','cancelled')
        """,
        (run_id, site_id),
    )
    _event(
        connection,
        run_id,
        "site_cancelled",
        "Site stopped before another page was fetched",
        site_id=site_id,
        details={"page": page},
    )
    _refresh_run_completion(connection, run_id)


def _claim_fetch(payload: dict[str, Any], task_id: str):
    run_id = str(payload["run_id"])
    site_id = int(payload["site_id"])
    page = int(payload["page"])
    connection = get_connection()
    try:
        row = connection.execute(
            """
            SELECT d.status,d.fetch_task_id,d.heartbeat_at,
                   r.mode,r.status AS run_status,r.requested_params,
                   r.cancellation_requested,
                   p.fetch_after,p.total_pages AS site_total_pages,
                   s.url,s.consumer_key,s.consumer_secret
            FROM sync_page_dispatches d
            JOIN sync_runs r ON r.run_id=d.run_id
            JOIN sync_site_progress p
              ON p.run_id=d.run_id AND p.site_id=d.site_id
            JOIN sites s ON s.id=d.site_id
            WHERE d.run_id=? AND d.site_id=? AND d.page=?
            FOR UPDATE OF d,p,r
            """,
            (run_id, site_id, page),
        ).fetchone()
        if not row:
            connection.rollback()
            return None
        if row["status"] in {"fetched", "writing", "completed", "cancelled"}:
            connection.rollback()
            return None
        if row["run_status"] in {"cancelled", "success", "error", "interrupted"}:
            connection.rollback()
            return None
        if bool(row["cancellation_requested"]):
            _cancel_site(connection, run_id, site_id, page)
            connection.commit()
            return None

        params = _json(row["requested_params"], {})
        fetch_after = row["fetch_after"]
        if page == 1 and row["mode"] != "deep" and not fetch_after:
            checkpoint = connection.execute(
                "SELECT MAX(date_modified) FROM orders WHERE source=?",
                (row["url"],),
            ).fetchone()[0]
            fetch_after = _subtract_minutes_from_iso(
                checkpoint,
                int(params.get("incremental_overlap_minutes", 10)),
            )
            connection.execute(
                """
                UPDATE sync_site_progress SET fetch_after=?
                WHERE run_id=? AND site_id=?
                """,
                (fetch_after, run_id, site_id),
            )
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='fetching',fetch_task_id=?,fetch_attempts=fetch_attempts+1,
                heartbeat_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP,
                error_message=NULL
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (task_id, run_id, site_id, page),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET status='fetching',current_page=?,started_at=COALESCE(started_at,CURRENT_TIMESTAMP),
                heartbeat_at=CURRENT_TIMESTAMP,error_message=NULL,version=version+1
            WHERE run_id=? AND site_id=?
            """,
            (page, run_id, site_id),
        )
        connection.execute(
            """
            UPDATE sync_runs
            SET status='running',current_site_id=?,heartbeat_at=CURRENT_TIMESTAMP,
                started_at=COALESCE(started_at,CURRENT_TIMESTAMP),version=version+1
            WHERE run_id=? AND status IN ('queued','running','recovering')
            """,
            (site_id, run_id),
        )
        connection.commit()
        return {
            "run_id": run_id,
            "site_id": site_id,
            "page": page,
            "mode": str(row["mode"]),
            "url": str(row["url"]).strip().rstrip("/"),
            "consumer_key": str(row["consumer_key"]),
            "consumer_secret": str(row["consumer_secret"]),
            "fetch_after": fetch_after,
            "params": params,
        }
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _fetch_page_data(claim: dict[str, Any]):
    per_page = int(claim["params"].get("per_page", 50))
    query = {
        "per_page": per_page,
        "page": int(claim["page"]),
        "orderby": "modified",
        "order": "asc",
    }
    if claim["mode"] != "deep" and claim.get("fetch_after"):
        query["modified_after"] = claim["fetch_after"]
    orders_url = urljoin(claim["url"] + "/", "wp-json/wc/v3/orders")
    session = requests.Session()
    try:
        auth = (claim["consumer_key"], claim["consumer_secret"])
        response, orders = _http_json_list(
            session, orders_url, auth=auth, params=query
        )
        for order in orders:
            order["source"] = claim["url"]

        notes: list[dict] = []
        notes_limit = (
            int(claim["params"].get("notes_per_page", 10))
            if claim["mode"] in {"quick", "auto"}
            else 0
        )
        candidates = [
            order for order in orders
            if order.get("id") is not None
            and order.get("status") in {"processing", "offline", "on-hold", "partial-shipped"}
        ][:notes_limit]
        for order in candidates:
            note_url = urljoin(
                claim["url"] + "/",
                f"wp-json/wc/v3/orders/{int(order['id'])}/notes",
            )
            _note_response, order_notes = _http_json_list(
                session, note_url, auth=auth
            )
            local_order_id = make_oid(claim["site_id"], order["id"])
            for note in order_notes:
                note["_local_order_id"] = local_order_id
                notes.append(note)

        header_pages = response.headers.get("X-WP-TotalPages")
        try:
            total_pages = max(int(claim["page"]), int(header_pages))
        except (TypeError, ValueError):
            total_pages = int(claim["page"]) + (1 if len(orders) >= per_page else 0)
        is_last_page = len(orders) < per_page or int(claim["page"]) >= total_pages
        return orders, notes, total_pages, is_last_page
    finally:
        close = getattr(session, "close", None)
        if close:
            close()


def _queue_page_write(
    claim: dict[str, Any],
    orders: list[dict],
    notes: list[dict],
    total_pages: int,
    is_last_page: bool,
) -> str:
    content_hash = _canonical_hash(orders, notes)
    payload = {
        "run_id": claim["run_id"],
        "site_id": claim["site_id"],
        "page": claim["page"],
        "orders": orders,
        "notes": notes,
        "content_hash": content_hash,
        "fetched_count": len(orders),
        "total_pages": int(total_pages),
        "is_last_page": bool(is_last_page),
    }
    connection = get_connection()
    try:
        dispatch = connection.execute(
            """
            SELECT status FROM sync_page_dispatches
            WHERE run_id=? AND site_id=? AND page=?
            FOR UPDATE
            """,
            (claim["run_id"], claim["site_id"], claim["page"]),
        ).fetchone()
        if not dispatch or dispatch["status"] in {
            "fetched", "writing", "completed", "cancelled"
        }:
            connection.rollback()
            return content_hash
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='fetched',content_hash=?,fetched_count=?,
                heartbeat_at=CURRENT_TIMESTAMP,updated_at=CURRENT_TIMESTAMP
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (
                content_hash,
                len(orders),
                claim["run_id"],
                claim["site_id"],
                claim["page"],
            ),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET status='writing',total_pages=GREATEST(total_pages,?),
                heartbeat_at=CURRENT_TIMESTAMP,last_content_hash=?,
                version=version+1
            WHERE run_id=? AND site_id=?
            """,
            (total_pages, content_hash, claim["run_id"], claim["site_id"]),
        )
        enqueue_outbox(
            connection,
            dedupe_key=(
                f"write:{claim['run_id']}:{claim['site_id']}:{claim['page']}"
            ),
            queue_name="sync_write",
            task_name="woo_sync.write_page",
            payload=payload,
        )
        _event(
            connection,
            claim["run_id"],
            "page_fetched",
            f"Fetched page {claim['page']} ({len(orders)} orders)",
            site_id=claim["site_id"],
            details={
                "page": claim["page"],
                "orders": len(orders),
                "notes": len(notes),
                "is_last_page": is_last_page,
            },
        )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()
    publish_pending_outbox(limit=10)
    return content_hash


@celery_app.task(
    bind=True,
    name="woo_sync.fetch_page",
    acks_late=True,
    reject_on_worker_lost=True,
    max_retries=MAX_FETCH_RETRIES,
    soft_time_limit=900,
    time_limit=960,
)
def fetch_page(self, payload: dict[str, Any]):
    run_id = str(payload.get("run_id", ""))
    site_id = int(payload.get("site_id", 0))
    page = int(payload.get("page", 0))
    if not run_id or site_id <= 0 or page <= 0:
        raise ValueError("invalid fetch payload")

    lock_connection = get_connection()
    locked = False
    try:
        locked = _site_lock(lock_connection, site_id)
        if not locked:
            _defer_busy_site(run_id, site_id, page)
            return {"site_busy": True, "deferred": True}
        claim = _claim_fetch(payload, str(self.request.id or ""))
        if not claim:
            return {"duplicate_or_cancelled": True}
        orders, notes, total_pages, is_last_page = _fetch_page_data(claim)
        content_hash = _queue_page_write(
            claim, orders, notes, total_pages, is_last_page
        )
        return {
            "run_id": run_id,
            "site_id": site_id,
            "page": page,
            "orders": len(orders),
            "content_hash": content_hash,
        }
    except AuthenticationFetchError as exc:
        mark_site_error(run_id, site_id, page, str(exc), auth_error=True)
        return {"auth_error": True}
    except PermanentFetchError as exc:
        mark_site_error(run_id, site_id, page, str(exc))
        return {"error": str(exc)}
    except TransientFetchError as exc:
        if int(self.request.retries or 0) >= MAX_FETCH_RETRIES:
            mark_site_error(run_id, site_id, page, str(exc))
            raise
        note_retry(run_id, site_id, page, str(exc))
        raise self.retry(
            exc=exc,
            countdown=_retry_delay(
                int(self.request.retries or 0), exc.retry_after
            ),
            max_retries=MAX_FETCH_RETRIES,
        )
    finally:
        if locked:
            _site_unlock(lock_connection, site_id)
        lock_connection.close()


def _page_payload(payload: dict[str, Any]):
    run_id = str(payload.get("run_id", ""))
    site_id = int(payload.get("site_id", 0))
    page = int(payload.get("page", 0))
    orders = payload.get("orders")
    notes = payload.get("notes")
    if (
        not run_id
        or site_id <= 0
        or page <= 0
        or not isinstance(orders, list)
        or len(orders) > 100
        or not isinstance(notes, list)
    ):
        raise ValueError("invalid write payload")
    actual_hash = _canonical_hash(orders, notes)
    if actual_hash != str(payload.get("content_hash", "")):
        raise ValueError("write payload hash mismatch")
    return run_id, site_id, page, orders, notes, actual_hash


def _claim_write(run_id: str, site_id: int, page: int, task_id: str) -> bool:
    connection = get_connection()
    try:
        dispatch = connection.execute(
            """
            SELECT status FROM sync_page_dispatches
            WHERE run_id=? AND site_id=? AND page=?
            FOR UPDATE
            """,
            (run_id, site_id, page),
        ).fetchone()
        if not dispatch:
            connection.rollback()
            return False
        receipt = connection.execute(
            """
            SELECT content_hash FROM sync_page_receipts
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (run_id, site_id, page),
        ).fetchone()
        if receipt:
            connection.rollback()
            return False
        if dispatch["status"] == "cancelled":
            connection.rollback()
            return False
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='writing',write_task_id=?,heartbeat_at=CURRENT_TIMESTAMP,
                updated_at=CURRENT_TIMESTAMP
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (task_id, run_id, site_id, page),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET status='writing',heartbeat_at=CURRENT_TIMESTAMP,version=version+1
            WHERE run_id=? AND site_id=?
            """,
            (run_id, site_id),
        )
        connection.execute(
            """
            UPDATE sync_runs
            SET heartbeat_at=CURRENT_TIMESTAMP,current_site_id=?,version=version+1
            WHERE run_id=?
            """,
            (site_id, run_id),
        )
        connection.commit()
        return True
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _write_page_transaction(payload: dict[str, Any]):
    run_id, site_id, page, orders, notes, content_hash = _page_payload(payload)
    connection = get_connection()
    result = None
    run_status = None
    try:
        dispatch = connection.execute(
            """
            SELECT content_hash,status FROM sync_page_dispatches
            WHERE run_id=? AND site_id=? AND page=?
            FOR UPDATE
            """,
            (run_id, site_id, page),
        ).fetchone()
        if not dispatch:
            raise ValueError("page dispatch does not exist")
        receipt = connection.execute(
            """
            SELECT content_hash,written_count,changed_count
            FROM sync_page_receipts
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (run_id, site_id, page),
        ).fetchone()
        if receipt:
            if str(receipt["content_hash"]) != content_hash:
                raise ValueError("conflicting content for an already-written page")
            connection.rollback()
            return {
                "duplicate": True,
                "written": int(receipt["written_count"]),
                "changed": int(receipt["changed_count"]),
                "planning_candidates": [],
            }
        if dispatch["content_hash"] and str(dispatch["content_hash"]) != content_hash:
            raise ValueError("dispatch content hash mismatch")

        result = upsert_orders_in_transaction(orders, connection)
        note_count = upsert_order_notes_in_transaction(notes, connection)
        planning_candidates = result.get("planning_candidates") or []
        post_commit_status = (
            "pending"
            if os.getenv("WOO_SYNC_POST_COMMIT_ACTIONS_ENABLED", "1") == "1"
            and planning_candidates
            else "skipped"
        )
        connection.execute(
            """
            INSERT INTO sync_page_receipts
                (run_id,site_id,page,content_hash,fetched_count,written_count,
                 changed_count,is_last_page,planning_candidates,post_commit_status)
            VALUES (?,?,?,?,?,?,?,?,?::jsonb,?)
            """,
            (
                run_id,
                site_id,
                page,
                content_hash,
                len(orders),
                result["written"],
                result["changed"],
                bool(payload.get("is_last_page")),
                json.dumps(
                    planning_candidates,
                    ensure_ascii=False,
                    separators=(",", ":"),
                ),
                post_commit_status,
            ),
        )
        if post_commit_status == "pending":
            enqueue_outbox(
                connection,
                dedupe_key=f"postcommit:{run_id}:{site_id}:{page}",
                queue_name="sync_write",
                task_name="woo_sync.post_commit_page",
                payload={"run_id": run_id, "site_id": site_id, "page": page},
            )
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='completed',heartbeat_at=CURRENT_TIMESTAMP,
                updated_at=CURRENT_TIMESTAMP,error_message=NULL
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (run_id, site_id, page),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET current_page=?,fetched_count=fetched_count+?,
                written_count=written_count+?,changed_count=changed_count+?,
                total_pages=GREATEST(total_pages,?),
                heartbeat_at=CURRENT_TIMESTAMP,error_message=NULL,
                version=version+1
            WHERE run_id=? AND site_id=?
            """,
            (
                page,
                len(orders),
                result["written"],
                result["changed"],
                int(payload.get("total_pages") or page),
                run_id,
                site_id,
            ),
        )
        run = connection.execute(
            """
            SELECT cancellation_requested,mode
            FROM sync_runs WHERE run_id=? FOR UPDATE
            """,
            (run_id,),
        ).fetchone()
        cancelled = bool(run["cancellation_requested"])
        is_last = bool(payload.get("is_last_page"))
        if cancelled:
            connection.execute(
                """
                UPDATE sync_site_progress
                SET status='cancelled',finished_at=CURRENT_TIMESTAMP,
                    heartbeat_at=CURRENT_TIMESTAMP,version=version+1
                WHERE run_id=? AND site_id=?
                """,
                (run_id, site_id),
            )
        elif is_last:
            connection.execute(
                """
                UPDATE sync_site_progress
                SET status='success',finished_at=CURRENT_TIMESTAMP,
                    heartbeat_at=CURRENT_TIMESTAMP,version=version+1
                WHERE run_id=? AND site_id=?
                """,
                (run_id, site_id),
            )
            connection.execute(
                """
                UPDATE sites
                SET last_sync=CURRENT_TIMESTAMP,api_status='ok',last_api_error=NULL
                WHERE id=?
                """,
                (site_id,),
            )
        else:
            connection.execute(
                """
                UPDATE sync_site_progress
                SET status='queued',heartbeat_at=CURRENT_TIMESTAMP,version=version+1
                WHERE run_id=? AND site_id=?
                """,
                (run_id, site_id),
            )
            enqueue_fetch_page(connection, run_id, site_id, page + 1)

        connection.execute(
            """
            UPDATE sync_runs
            SET completed_pages=(
                    SELECT COUNT(*) FROM sync_page_receipts WHERE run_id=?
                ),
                fetched_orders=(
                    SELECT COALESCE(SUM(fetched_count),0)
                    FROM sync_page_receipts WHERE run_id=?
                ),
                written_orders=(
                    SELECT COALESCE(SUM(written_count),0)
                    FROM sync_page_receipts WHERE run_id=?
                ),
                changed_orders=(
                    SELECT COALESCE(SUM(changed_count),0)
                    FROM sync_page_receipts WHERE run_id=?
                ),
                total_pages=(
                    SELECT COALESCE(SUM(GREATEST(total_pages,current_page)),0)
                    FROM sync_site_progress WHERE run_id=?
                ),
                heartbeat_at=CURRENT_TIMESTAMP,version=version+1
            WHERE run_id=?
            """,
            (run_id, run_id, run_id, run_id, run_id, run_id),
        )
        _event(
            connection,
            run_id,
            "page_committed",
            f"Committed page {page} ({result['written']} orders, {note_count} notes)",
            site_id=site_id,
            details={
                "page": page,
                "written": result["written"],
                "changed": result["changed"],
                "notes": note_count,
                "last": is_last,
                "cancelled": cancelled,
            },
        )
        run_status = _refresh_run_completion(connection, run_id)
        connection.commit()
        result["notes"] = note_count
        result["run_status"] = run_status
        return result
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _record_write_retry(run_id: str, site_id: int, page: int, message: str) -> None:
    connection = get_connection()
    try:
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='writing',error_message=?,heartbeat_at=CURRENT_TIMESTAMP,
                updated_at=CURRENT_TIMESTAMP
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (message[:1000], run_id, site_id, page),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET status='recovering',retry_count=retry_count+1,
                error_message=?,heartbeat_at=CURRENT_TIMESTAMP,version=version+1
            WHERE run_id=? AND site_id=?
            """,
            (message[:1000], run_id, site_id),
        )
        connection.execute(
            """
            UPDATE sync_runs
            SET status=CASE WHEN status='cancelling' THEN status ELSE 'recovering' END,
                heartbeat_at=CURRENT_TIMESTAMP,recovery_count=recovery_count+1,
                version=version+1
            WHERE run_id=?
            """,
            (run_id,),
        )
        connection.commit()
    finally:
        connection.close()


@celery_app.task(
    bind=True,
    name="woo_sync.write_page",
    acks_late=True,
    reject_on_worker_lost=True,
    max_retries=MAX_WRITE_RETRIES,
    soft_time_limit=240,
    time_limit=300,
)
def write_page(self, payload: dict[str, Any]):
    run_id, site_id, page, _orders, _notes, _hash = _page_payload(payload)
    if not _claim_write(run_id, site_id, page, str(self.request.id or "")):
        return {"duplicate_or_cancelled": True}
    try:
        result = _write_page_transaction(payload)
    except Exception as exc:
        message = type(exc).__name__ + ": " + str(exc)[:800]
        if int(self.request.retries or 0) >= MAX_WRITE_RETRIES:
            mark_site_error(run_id, site_id, page, "database write failed: " + message)
            raise
        _record_write_retry(run_id, site_id, page, message)
        raise self.retry(
            exc=exc,
            countdown=_retry_delay(int(self.request.retries or 0)),
            max_retries=MAX_WRITE_RETRIES,
        )

    publish_pending_outbox(limit=20)
    return {
        "run_id": run_id,
        "site_id": site_id,
        "page": page,
        "written": int(result.get("written") or 0),
        "changed": int(result.get("changed") or 0),
        "duplicate": bool(result.get("duplicate")),
    }


def _claim_post_commit(run_id: str, site_id: int, page: int):
    connection = get_connection()
    try:
        receipt = connection.execute(
            """
            SELECT planning_candidates,post_commit_status
            FROM sync_page_receipts
            WHERE run_id=? AND site_id=? AND page=?
            FOR UPDATE
            """,
            (run_id, site_id, page),
        ).fetchone()
        if not receipt or receipt["post_commit_status"] in {"completed", "skipped"}:
            connection.rollback()
            return None
        connection.execute(
            """
            UPDATE sync_page_receipts
            SET post_commit_status='processing',
                post_commit_attempts=post_commit_attempts+1,
                post_commit_heartbeat_at=CURRENT_TIMESTAMP,
                post_commit_error=NULL
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (run_id, site_id, page),
        )
        connection.commit()
        return _json(receipt["planning_candidates"], [])
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def _finish_post_commit(
    run_id: str, site_id: int, page: int, status: str, error: str | None = None
) -> None:
    connection = get_connection()
    try:
        connection.execute(
            """
            UPDATE sync_page_receipts
            SET post_commit_status=?,post_commit_error=?,
                post_commit_heartbeat_at=CURRENT_TIMESTAMP,
                post_commit_finished_at=CASE WHEN ?='completed'
                    THEN CURRENT_TIMESTAMP ELSE NULL END
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (status, error, status, run_id, site_id, page),
        )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


@celery_app.task(
    bind=True,
    name="woo_sync.post_commit_page",
    acks_late=True,
    reject_on_worker_lost=True,
    max_retries=MAX_WRITE_RETRIES,
    soft_time_limit=240,
    time_limit=300,
)
def post_commit_page(self, payload: dict[str, Any]):
    run_id = str(payload.get("run_id") or "")
    site_id = int(payload.get("site_id") or 0)
    page = int(payload.get("page") or 0)
    if not run_id or site_id <= 0 or page <= 0:
        raise ValueError("invalid post-commit payload")
    candidates = _claim_post_commit(run_id, site_id, page)
    if candidates is None:
        return {"duplicate_or_skipped": True}
    try:
        run_post_commit_sync_actions(candidates, strict=True)
    except Exception as exc:
        message = type(exc).__name__ + ": " + str(exc)[:800]
        _finish_post_commit(run_id, site_id, page, "error", message)
        if int(self.request.retries or 0) >= MAX_WRITE_RETRIES:
            raise
        raise self.retry(
            exc=exc,
            countdown=_retry_delay(int(self.request.retries or 0)),
            max_retries=MAX_WRITE_RETRIES,
        )
    _finish_post_commit(run_id, site_id, page, "completed")
    return {"run_id": run_id, "site_id": site_id, "page": page}


@celery_app.task(
    name="woo_sync.maintenance",
    acks_late=True,
    reject_on_worker_lost=True,
)
def maintenance():
    return recover_stale_work()


def _sync_settings(*keys: str) -> dict[str, str]:
    connection = get_connection()
    try:
        if not keys:
            return {}
        placeholders = ",".join("?" for _ in keys)
        rows = connection.execute(
            f"SELECT key,value FROM settings WHERE key IN ({placeholders})",
            tuple(keys),
        ).fetchall()
        return {str(row["key"]): str(row["value"] or "") for row in rows}
    finally:
        connection.close()


def _truthy(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _auto_due() -> tuple[bool, dict[str, Any]]:
    settings = _sync_settings("autosync_enabled", "autosync_interval")
    enabled = _truthy(settings.get("autosync_enabled"), False)
    try:
        interval = int(settings.get("autosync_interval", "3600"))
    except (TypeError, ValueError):
        interval = 3600
    interval = min(86400, max(300, interval))
    if not enabled:
        return False, {"reason": "disabled", "interval": interval}
    connection = get_connection()
    try:
        row = connection.execute(
            """
            SELECT COALESCE(
                MAX(created_at) <= CURRENT_TIMESTAMP - (? * interval '1 second'),
                true
            ) AS due
            FROM sync_runs
            WHERE mode='auto' AND created_by='celery-beat:auto'
            """,
            (interval,),
        ).fetchone()
        return bool(row and row["due"]), {"interval": interval}
    finally:
        connection.close()


def _deep_due() -> tuple[bool, dict[str, Any]]:
    settings = _sync_settings(
        "deep_sync_enabled", "deep_sync_hour", "deep_sync_minute"
    )
    enabled = _truthy(settings.get("deep_sync_enabled"), True)
    try:
        hour = min(23, max(0, int(settings.get("deep_sync_hour", "3"))))
        minute = min(59, max(0, int(settings.get("deep_sync_minute", "30"))))
    except (TypeError, ValueError):
        hour, minute = 3, 30
    now = datetime.now(LOCAL_TIMEZONE)
    scheduled = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if not enabled or now < scheduled:
        return False, {
            "reason": "disabled" if not enabled else "not_due",
            "hour": hour,
            "minute": minute,
        }
    connection = get_connection()
    try:
        row = connection.execute(
            """
            SELECT EXISTS (
                SELECT 1 FROM sync_runs
                WHERE mode='deep' AND created_by='celery-beat:deep'
                  AND timezone('Asia/Hong_Kong',created_at)::date=?::date
            ) AS already_started
            """,
            (now.date(),),
        ).fetchone()
        return not bool(row and row["already_started"]), {
            "hour": hour,
            "minute": minute,
        }
    finally:
        connection.close()


@celery_app.task(
    name="woo_sync.schedule_auto",
    acks_late=True,
    reject_on_worker_lost=True,
)
def schedule_auto():
    due, schedule = _auto_due()
    if not due:
        return {"created": False, "schedule": schedule}
    status, created = start_sync(
        mode="auto",
        created_by="celery-beat:auto",
        params={
            "per_page": 50,
            "incremental_overlap_minutes": 10,
            "notes_per_page": 10,
        },
    )
    return {"run_id": status["run_id"], "created": created, "schedule": schedule}


@celery_app.task(
    name="woo_sync.schedule_deep",
    acks_late=True,
    reject_on_worker_lost=True,
)
def schedule_deep():
    due, schedule = _deep_due()
    if not due:
        return {"created": False, "schedule": schedule}
    status, created = start_sync(
        mode="deep",
        created_by="celery-beat:deep",
        params={"per_page": 50, "notes_per_page": 0},
    )
    return {"run_id": status["run_id"], "created": created, "schedule": schedule}
