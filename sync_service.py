"""Durable orchestration for WooCommerce synchronization.

PostgreSQL is the source of truth. Redis/Celery only transports bounded
messages, and every message is recoverable from sync_task_outbox.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import uuid
from typing import Any, Iterable

import db_backend as db


DB_FILE = os.getenv("WOO_SQLITE_PATH", "woocommerce_orders.db")
ACTIVE_RUN_STATUSES = ("queued", "running", "recovering", "cancelling")
TERMINAL_RUN_STATUSES = ("cancelled", "success", "error", "interrupted")
TERMINAL_SITE_STATUSES = ("cancelled", "success", "error", "auth_error")
DEFAULT_PER_PAGE = 50
MIN_PER_PAGE = 50
MAX_PER_PAGE = 100
HEARTBEAT_STALE_SECONDS = int(os.getenv("WOO_SYNC_HEARTBEAT_STALE_SECONDS", "90"))
RECOVERY_STALE_SECONDS = int(os.getenv("WOO_SYNC_RECOVERY_STALE_SECONDS", "180"))


class SyncConfigurationError(RuntimeError):
    pass


def _require_postgres() -> None:
    if not db.is_postgres_backend():
        raise SyncConfigurationError(
            "Celery synchronization requires WOO_DB_BACKEND=postgres"
        )


def get_connection():
    _require_postgres()
    connection = db.connect(DB_FILE, timeout=10)
    connection.row_factory = db.Row
    return connection


def _json(value: Any, default: Any):
    if value is None:
        return default
    if isinstance(value, (dict, list)):
        return value
    try:
        return json.loads(value)
    except (TypeError, ValueError):
        return default


def _iso(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat()
    return str(value)


def _parse_time(value: Any) -> dt.datetime | None:
    if not value:
        return None
    if isinstance(value, dt.datetime):
        parsed = value
    else:
        try:
            parsed = dt.datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        except ValueError:
            return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=dt.timezone.utc)
    return parsed.astimezone(dt.timezone.utc)


def _bounded_per_page(value: Any) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        parsed = DEFAULT_PER_PAGE
    return min(MAX_PER_PAGE, max(MIN_PER_PAGE, parsed))


def _event(
    connection,
    run_id: str,
    event_type: str,
    message: str,
    *,
    site_id: int | None = None,
    level: str = "info",
    details: dict[str, Any] | None = None,
) -> None:
    connection.execute(
        """
        INSERT INTO sync_events
            (run_id,site_id,level,event_type,message,details)
        VALUES (?,?,?,?,?,?::jsonb)
        """,
        (
            run_id,
            site_id,
            level,
            event_type,
            str(message)[:2000],
            json.dumps(details or {}, ensure_ascii=False, separators=(",", ":")),
        ),
    )


def enqueue_outbox(
    connection,
    *,
    dedupe_key: str,
    queue_name: str,
    task_name: str,
    payload: dict[str, Any],
) -> None:
    if queue_name not in {"sync_fetch", "sync_write"}:
        raise ValueError("invalid sync queue")
    connection.execute(
        """
        INSERT INTO sync_task_outbox
            (dedupe_key,queue_name,task_name,payload,status,available_at,updated_at)
        VALUES (?,?,?,?::jsonb,'pending',CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
        ON CONFLICT(dedupe_key) DO UPDATE SET
            status=CASE
                WHEN sync_task_outbox.status IN ('published','publishing')
                    THEN sync_task_outbox.status
                ELSE 'pending'
            END,
            available_at=CASE
                WHEN sync_task_outbox.status IN ('published','publishing')
                    THEN sync_task_outbox.available_at
                ELSE CURRENT_TIMESTAMP
            END,
            updated_at=CURRENT_TIMESTAMP,
            last_error=NULL
        """,
        (
            dedupe_key,
            queue_name,
            task_name,
            json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        ),
    )


def enqueue_fetch_page(
    connection, run_id: str, site_id: int, page: int, *, reset=False
) -> None:
    connection.execute(
        """
        INSERT INTO sync_page_dispatches
            (run_id,site_id,page,status,available_at,updated_at)
        VALUES (?,?,?,'queued',CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
        ON CONFLICT(run_id,site_id,page) DO UPDATE SET
            status=CASE
                WHEN sync_page_dispatches.status IN ('completed','cancelled')
                    THEN sync_page_dispatches.status
                WHEN ? THEN 'retry'
                ELSE sync_page_dispatches.status
            END,
            available_at=CURRENT_TIMESTAMP,
            updated_at=CURRENT_TIMESTAMP
        """,
        (run_id, site_id, page, bool(reset)),
    )
    enqueue_outbox(
        connection,
        dedupe_key=f"fetch:{run_id}:{site_id}:{page}",
        queue_name="sync_fetch",
        task_name="woo_sync.fetch_page",
        payload={"run_id": run_id, "site_id": int(site_id), "page": int(page)},
    )


def _normalize_site_ids(site_ids: Iterable[Any] | None) -> list[int] | None:
    if site_ids is None:
        return None
    normalized: list[int] = []
    for value in site_ids:
        parsed = int(value)
        if parsed not in normalized:
            normalized.append(parsed)
    return normalized


def _load_sites(connection, site_ids: list[int] | None):
    if site_ids is None:
        return connection.execute(
            "SELECT id,url FROM sites ORDER BY id"
        ).fetchall()
    if not site_ids:
        return []
    placeholders = ",".join("?" for _ in site_ids)
    return connection.execute(
        f"SELECT id,url FROM sites WHERE id IN ({placeholders}) ORDER BY id",
        tuple(site_ids),
    ).fetchall()


def active_run(connection=None) -> dict[str, Any] | None:
    own = connection is None
    connection = connection or get_connection()
    try:
        row = connection.execute(
            """
            SELECT run_id,mode,status,created_by,created_at,started_at,
                   heartbeat_at,total_sites,completed_sites,total_pages,
                   completed_pages,fetched_orders,written_orders,changed_orders,
                   cancellation_requested
            FROM sync_runs
            WHERE status IN ('queued','running','recovering','cancelling')
            ORDER BY created_at
            LIMIT 1
            """
        ).fetchone()
        return dict(row) if row else None
    finally:
        if own:
            connection.close()


def start_sync(
    *,
    mode: str,
    created_by: str | None,
    site_ids: Iterable[Any] | None = None,
    params: dict[str, Any] | None = None,
    publish: bool = True,
) -> tuple[dict[str, Any], bool]:
    """Create one globally-exclusive run or return the already-active run."""

    _require_postgres()
    if mode not in {"quick", "auto", "deep"}:
        raise ValueError("invalid synchronization mode")
    normalized_ids = _normalize_site_ids(site_ids)
    clean_params = dict(params or {})
    clean_params["per_page"] = _bounded_per_page(clean_params.get("per_page"))
    clean_params["incremental_overlap_minutes"] = max(
        0, min(1440, int(clean_params.get("incremental_overlap_minutes", 10)))
    )
    clean_params["notes_per_page"] = max(
        0, min(10, int(clean_params.get("notes_per_page", 10)))
    )
    run_id = str(uuid.uuid4())
    connection = get_connection()
    created = False
    try:
        sites = _load_sites(connection, normalized_ids)
        if normalized_ids is not None and len(sites) != len(normalized_ids):
            found = {int(row["id"]) for row in sites}
            missing = sorted(set(normalized_ids) - found)
            raise ValueError("unknown site ids: " + ",".join(map(str, missing)))
        if not sites:
            raise ValueError("no sites configured for synchronization")
        connection.execute(
            """
            INSERT INTO sync_runs
                (run_id,mode,status,created_by,requested_params,total_sites,
                 started_at,heartbeat_at)
            VALUES (?,?,'queued',?,?::jsonb,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)
            """,
            (
                run_id,
                mode,
                str(created_by or "")[:255],
                json.dumps(clean_params, ensure_ascii=False, separators=(",", ":")),
                len(sites),
            ),
        )
        for site in sites:
            site_id = int(site["id"])
            connection.execute(
                """
                INSERT INTO sync_site_progress
                    (run_id,site_id,status,current_page,heartbeat_at)
                VALUES (?,?,'queued',0,CURRENT_TIMESTAMP)
                """,
                (run_id, site_id),
            )
            enqueue_fetch_page(connection, run_id, site_id, 1)
        _event(
            connection,
            run_id,
            "run_created",
            f"{mode} synchronization queued for {len(sites)} site(s)",
            details={"total_sites": len(sites), "mode": mode},
        )
        connection.commit()
        created = True
    except db.IntegrityError:
        connection.rollback()
        existing = active_run(connection)
        if not existing:
            raise
        return get_run_status(str(existing["run_id"]), connection=connection), False
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()

    if publish:
        publish_pending_outbox(limit=max(10, len(sites)))
    return get_run_status(run_id), created


def _refresh_run_completion(connection, run_id: str) -> str:
    summary = connection.execute(
        """
        SELECT COUNT(*) AS total,
               COUNT(*) FILTER (
                   WHERE status IN ('cancelled','success','error','auth_error')
               ) AS terminal,
               COUNT(*) FILTER (WHERE status='success') AS succeeded,
               COUNT(*) FILTER (WHERE status IN ('error','auth_error')) AS failed
        FROM sync_site_progress
        WHERE run_id=?
        """,
        (run_id,),
    ).fetchone()
    run = connection.execute(
        "SELECT status,cancellation_requested FROM sync_runs WHERE run_id=? FOR UPDATE",
        (run_id,),
    ).fetchone()
    if not run:
        raise KeyError(run_id)
    total = int(summary["total"] or 0)
    terminal = int(summary["terminal"] or 0)
    connection.execute(
        """
        UPDATE sync_runs
        SET completed_sites=?,heartbeat_at=CURRENT_TIMESTAMP,version=version+1
        WHERE run_id=?
        """,
        (terminal, run_id),
    )
    if total and terminal == total:
        if bool(run["cancellation_requested"]):
            status = "cancelled"
        elif int(summary["failed"] or 0):
            status = "error"
        else:
            status = "success"
        connection.execute(
            """
            UPDATE sync_runs
            SET status=?,finished_at=CURRENT_TIMESTAMP,heartbeat_at=CURRENT_TIMESTAMP,
                current_site_id=NULL,version=version+1
            WHERE run_id=?
            """,
            (status, run_id),
        )
        return status
    if run["status"] == "queued":
        connection.execute(
            "UPDATE sync_runs SET status='running',started_at=COALESCE(started_at,CURRENT_TIMESTAMP) WHERE run_id=?",
            (run_id,),
        )
        return "running"
    return str(run["status"])


def mark_site_error(
    run_id: str,
    site_id: int,
    page: int,
    message: str,
    *,
    auth_error: bool = False,
) -> None:
    connection = get_connection()
    try:
        status = "auth_error" if auth_error else "error"
        safe_message = str(message)[:2000]
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status=?,error_message=?,heartbeat_at=CURRENT_TIMESTAMP,
                updated_at=CURRENT_TIMESTAMP
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (status, safe_message, run_id, site_id, page),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET status=?,error_message=?,finished_at=CURRENT_TIMESTAMP,
                heartbeat_at=CURRENT_TIMESTAMP,version=version+1
            WHERE run_id=? AND site_id=?
            """,
            (status, safe_message, run_id, site_id),
        )
        _event(
            connection,
            run_id,
            "site_error",
            safe_message,
            site_id=site_id,
            level="error",
            details={"page": page, "auth_error": auth_error},
        )
        _refresh_run_completion(connection, run_id)
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def note_retry(run_id: str, site_id: int, page: int, message: str) -> None:
    connection = get_connection()
    try:
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='retry',error_message=?,heartbeat_at=CURRENT_TIMESTAMP,
                updated_at=CURRENT_TIMESTAMP
            WHERE run_id=? AND site_id=? AND page=?
            """,
            (str(message)[:1000], run_id, site_id, page),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET retry_count=retry_count+1,status='recovering',
                heartbeat_at=CURRENT_TIMESTAMP,error_message=?,
                version=version+1
            WHERE run_id=? AND site_id=?
            """,
            (str(message)[:1000], run_id, site_id),
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
        _event(
            connection,
            run_id,
            "page_retry",
            str(message)[:1000],
            site_id=site_id,
            level="warning",
            details={"page": page},
        )
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def cancel_sync(run_id: str, requested_by: str | None = None) -> dict[str, Any]:
    connection = get_connection()
    try:
        run = connection.execute(
            "SELECT status FROM sync_runs WHERE run_id=? FOR UPDATE", (run_id,)
        ).fetchone()
        if not run:
            raise KeyError(run_id)
        if run["status"] in TERMINAL_RUN_STATUSES:
            connection.rollback()
            return get_run_status(run_id)
        connection.execute(
            """
            UPDATE sync_runs
            SET cancellation_requested=true,status='cancelling',
                heartbeat_at=CURRENT_TIMESTAMP,version=version+1
            WHERE run_id=?
            """,
            (run_id,),
        )
        connection.execute(
            """
            UPDATE sync_site_progress
            SET status='cancelled',finished_at=CURRENT_TIMESTAMP,
                heartbeat_at=CURRENT_TIMESTAMP,error_message=NULL,version=version+1
            WHERE run_id=? AND status='queued'
            """,
            (run_id,),
        )
        connection.execute(
            """
            UPDATE sync_page_dispatches
            SET status='cancelled',updated_at=CURRENT_TIMESTAMP,
                heartbeat_at=CURRENT_TIMESTAMP
            WHERE run_id=? AND status IN ('queued','retry')
            """,
            (run_id,),
        )
        connection.execute(
            """
            UPDATE sync_task_outbox
            SET status='cancelled',updated_at=CURRENT_TIMESTAMP
            WHERE payload->>'run_id'=? AND status IN ('pending','error')
            """,
            (run_id,),
        )
        _event(
            connection,
            run_id,
            "cancellation_requested",
            "Synchronization cancellation requested",
            details={"requested_by": str(requested_by or "")[:255]},
        )
        _refresh_run_completion(connection, run_id)
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()
    return get_run_status(run_id)


def get_run_status(run_id: str, *, connection=None) -> dict[str, Any]:
    try:
        normalized = str(uuid.UUID(str(run_id)))
    except (ValueError, TypeError, AttributeError) as exc:
        raise KeyError(run_id) from exc
    own = connection is None
    connection = connection or get_connection()
    try:
        run = connection.execute(
            """
            SELECT run_id,mode,status,created_by,requested_params,created_at,
                   started_at,heartbeat_at,finished_at,total_sites,completed_sites,
                   total_pages,completed_pages,fetched_orders,written_orders,
                   changed_orders,current_site_id,error_message,
                   cancellation_requested,recovery_count,version
            FROM sync_runs WHERE run_id=?
            """,
            (normalized,),
        ).fetchone()
        if not run:
            raise KeyError(normalized)
        sites = connection.execute(
            """
            SELECT p.site_id,s.url,p.status,p.current_page,p.fetched_count,
                   p.written_count,p.changed_count,p.retry_count,p.heartbeat_at,
                   p.started_at,p.finished_at,p.error_message
            FROM sync_site_progress p
            JOIN sites s ON s.id=p.site_id
            WHERE p.run_id=?
            ORDER BY p.site_id
            """,
            (normalized,),
        ).fetchall()
        events = connection.execute(
            """
            SELECT site_id,level,event_type,message,created_at
            FROM sync_events WHERE run_id=?
            ORDER BY event_id DESC LIMIT 30
            """,
            (normalized,),
        ).fetchall()
        item = dict(run)
        item["run_id"] = str(item["run_id"])
        item["sync_id"] = item["run_id"]
        item["requested_params"] = _json(item.get("requested_params"), {})
        item["cancellation_requested"] = bool(item["cancellation_requested"])
        for key in ("created_at", "started_at", "heartbeat_at", "finished_at"):
            item[key] = _iso(item.get(key))
        site_items = []
        current = None
        for row in sites:
            value = dict(row)
            for key in ("heartbeat_at", "started_at", "finished_at"):
                value[key] = _iso(value.get(key))
            value["site_id"] = int(value["site_id"])
            site_items.append(value)
            if current is None and value["status"] in {
                "fetching", "writing", "recovering"
            }:
                current = value
        heartbeat = _parse_time(run["heartbeat_at"])
        stale_seconds = None
        if heartbeat:
            stale_seconds = max(
                0.0,
                (dt.datetime.now(dt.timezone.utc) - heartbeat).total_seconds(),
            )
        item["stale_seconds"] = round(stale_seconds, 1) if stale_seconds is not None else None
        item["interruption_state"] = (
            "recovering"
            if item["status"] in ACTIVE_RUN_STATUSES
            and stale_seconds is not None
            and stale_seconds > HEARTBEAT_STALE_SECONDS
            else None
        )
        item["sites"] = site_items
        item["current_site"] = current
        item["current_page"] = int(current["current_page"]) if current else 0
        item["retry_count"] = sum(int(row["retry_count"] or 0) for row in site_items)
        item["logs"] = [
            f"[{_iso(row['created_at'])}] {row['message']}" for row in reversed(events)
        ]
        item["message"] = _status_message(item, current)
        return item
    finally:
        if own:
            connection.close()


def _status_message(run: dict[str, Any], current: dict[str, Any] | None) -> str:
    labels = {"quick": "快速同步", "auto": "自动同步", "deep": "深度同步"}
    prefix = labels.get(str(run.get("mode")), str(run.get("mode")))
    status = str(run.get("status"))
    if status == "success":
        return f"{prefix}已完成"
    if status == "cancelled":
        return f"{prefix}已取消"
    if status == "error":
        return run.get("error_message") or f"{prefix}完成，但有站点失败"
    if run.get("interruption_state") == "recovering":
        return "任务已中断/正在恢复"
    if status == "cancelling":
        return "正在安全取消；已进入写事务的分页会正常完成"
    if current:
        return f"{prefix}：站点 {current['site_id']}，第 {current['current_page']} 页"
    return f"{prefix}已排队"


def _task_id(dedupe_key: str) -> str:
    digest = hashlib.sha256(dedupe_key.encode("utf-8")).hexdigest()
    return "woo-sync-" + digest[:40]


def publish_pending_outbox(limit: int = 100) -> int:
    """Publish durable outbox rows at least once.

    A crash after broker publish but before the PostgreSQL status update may
    publish a duplicate. Fetch claims and page receipts make that harmless.
    """

    _require_postgres()
    published = 0
    for _ in range(max(0, int(limit))):
        connection = get_connection()
        row = None
        try:
            row = connection.execute(
                """
                SELECT outbox_id,dedupe_key,queue_name,task_name,payload
                FROM sync_task_outbox
                WHERE status IN ('pending','error')
                  AND available_at<=CURRENT_TIMESTAMP
                ORDER BY outbox_id
                FOR UPDATE SKIP LOCKED
                LIMIT 1
                """
            ).fetchone()
            if not row:
                connection.rollback()
                connection.close()
                break
            task_id = _task_id(str(row["dedupe_key"]))
            connection.execute(
                """
                UPDATE sync_task_outbox
                SET status='publishing',celery_task_id=?,attempts=attempts+1,
                    updated_at=CURRENT_TIMESTAMP,last_error=NULL
                WHERE outbox_id=?
                """,
                (task_id, row["outbox_id"]),
            )
            connection.commit()
        except Exception:
            connection.rollback()
            connection.close()
            raise
        connection.close()

        try:
            from celery_app import celery_app

            celery_app.send_task(
                str(row["task_name"]),
                args=[_json(row["payload"], {})],
                queue=str(row["queue_name"]),
                task_id=task_id,
                serializer="json",
                delivery_mode=2,
            )
        except Exception as exc:
            connection = get_connection()
            try:
                connection.execute(
                    """
                    UPDATE sync_task_outbox
                    SET status='error',last_error=?,
                        available_at=CURRENT_TIMESTAMP + interval '30 seconds',
                        updated_at=CURRENT_TIMESTAMP
                    WHERE outbox_id=? AND status='publishing'
                    """,
                    (type(exc).__name__ + ": " + str(exc)[:500], row["outbox_id"]),
                )
                connection.commit()
            finally:
                connection.close()
            break

        connection = get_connection()
        try:
            connection.execute(
                """
                UPDATE sync_task_outbox
                SET status='published',published_at=CURRENT_TIMESTAMP,
                    updated_at=CURRENT_TIMESTAMP,last_error=NULL
                WHERE outbox_id=? AND status='publishing'
                """,
                (row["outbox_id"],),
            )
            connection.commit()
            published += 1
        finally:
            connection.close()
    return published


def recover_stale_work() -> dict[str, int]:
    """Requeue stale broker hand-offs and interrupted fetch/write tasks."""

    connection = get_connection()
    counts = {"outbox": 0, "dispatches": 0, "post_commit": 0, "runs": 0}
    try:
        cursor = connection.execute(
            """
            UPDATE sync_task_outbox
            SET status='pending',available_at=CURRENT_TIMESTAMP,
                updated_at=CURRENT_TIMESTAMP,
                last_error=COALESCE(last_error,'publisher interrupted')
            WHERE status='publishing'
              AND updated_at<CURRENT_TIMESTAMP - interval '2 minutes'
            """
        )
        counts["outbox"] += max(0, int(cursor.rowcount or 0))

        stale_post_commit = connection.execute(
            """
            SELECT run_id,site_id,page,post_commit_status
            FROM sync_page_receipts
            WHERE post_commit_status IN ('pending','processing','error')
              AND COALESCE(post_commit_heartbeat_at,committed_at)
                    < CURRENT_TIMESTAMP - (? * interval '1 second')
            FOR UPDATE SKIP LOCKED
            """,
            (RECOVERY_STALE_SECONDS,),
        ).fetchall()
        for receipt in stale_post_commit:
            run_id = str(receipt["run_id"])
            site_id = int(receipt["site_id"])
            page = int(receipt["page"])
            enqueue_outbox(
                connection,
                dedupe_key=f"postcommit:{run_id}:{site_id}:{page}",
                queue_name="sync_write",
                task_name="woo_sync.post_commit_page",
                payload={"run_id": run_id, "site_id": site_id, "page": page},
            )
            connection.execute(
                """
                UPDATE sync_task_outbox
                SET status='pending',available_at=CURRENT_TIMESTAMP,
                    updated_at=CURRENT_TIMESTAMP,
                    last_error=COALESCE(last_error,'post-commit delivery stale')
                WHERE dedupe_key=?
                  AND status IN ('published','publishing','error')
                """,
                (f"postcommit:{run_id}:{site_id}:{page}",),
            )
            connection.execute(
                """
                UPDATE sync_page_receipts
                SET post_commit_status=CASE
                        WHEN post_commit_status='processing' THEN 'error'
                        ELSE post_commit_status
                    END,
                    post_commit_error=CASE
                        WHEN post_commit_status='processing'
                            THEN COALESCE(post_commit_error,'worker interrupted')
                        ELSE post_commit_error
                    END,
                    post_commit_heartbeat_at=CURRENT_TIMESTAMP
                WHERE run_id=? AND site_id=? AND page=?
                """,
                (run_id, site_id, page),
            )
            counts["post_commit"] += 1

        stale = connection.execute(
            """
            SELECT run_id,site_id,page,status
            FROM sync_page_dispatches
            WHERE (
                    status IN ('fetching','fetched','writing')
                    AND heartbeat_at<CURRENT_TIMESTAMP - (? * interval '1 second')
                  )
               OR (
                    status IN ('queued','retry')
                    AND updated_at<CURRENT_TIMESTAMP - (? * interval '1 second')
                  )
            FOR UPDATE SKIP LOCKED
            """,
            (RECOVERY_STALE_SECONDS, RECOVERY_STALE_SECONDS),
        ).fetchall()
        for row in stale:
            run_id = str(row["run_id"])
            site_id = int(row["site_id"])
            page = int(row["page"])
            run = connection.execute(
                "SELECT cancellation_requested FROM sync_runs WHERE run_id=?",
                (run_id,),
            ).fetchone()
            if not run:
                continue
            if bool(run["cancellation_requested"]):
                connection.execute(
                    """
                    UPDATE sync_page_dispatches
                    SET status='cancelled',updated_at=CURRENT_TIMESTAMP
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
                    """,
                    (run_id, site_id),
                )
                _refresh_run_completion(connection, run_id)
                continue
            if row["status"] in {"queued", "retry", "fetching"}:
                enqueue_fetch_page(connection, run_id, site_id, page, reset=True)
                # enqueue_outbox normally preserves an already-published row.
                # Here PostgreSQL proves that the worker lease is stale, so a
                # fresh at-least-once delivery is intentional and safe.
                connection.execute(
                    """
                    UPDATE sync_task_outbox
                    SET status='pending',available_at=CURRENT_TIMESTAMP,
                        updated_at=CURRENT_TIMESTAMP
                    WHERE dedupe_key=?
                      AND status IN ('published','publishing','error')
                    """,
                    (f"fetch:{run_id}:{site_id}:{page}",),
                )
            else:
                connection.execute(
                    """
                    UPDATE sync_task_outbox
                    SET status='pending',available_at=CURRENT_TIMESTAMP,
                        updated_at=CURRENT_TIMESTAMP
                    WHERE dedupe_key=? AND status IN ('published','publishing','error')
                    """,
                    (f"write:{run_id}:{site_id}:{page}",),
                )
                connection.execute(
                    """
                    UPDATE sync_page_dispatches
                    SET status='fetched',updated_at=CURRENT_TIMESTAMP
                    WHERE run_id=? AND site_id=? AND page=?
                    """,
                    (run_id, site_id, page),
                )
            connection.execute(
                """
                UPDATE sync_site_progress
                SET status='recovering',retry_count=retry_count+1,
                    heartbeat_at=CURRENT_TIMESTAMP,version=version+1
                WHERE run_id=? AND site_id=?
                """,
                (run_id, site_id),
            )
            connection.execute(
                """
                UPDATE sync_runs
                SET status='recovering',heartbeat_at=CURRENT_TIMESTAMP,
                    recovery_count=recovery_count+1,version=version+1
                WHERE run_id=? AND status IN ('queued','running','recovering')
                """,
                (run_id,),
            )
            counts["dispatches"] += 1

        cursor = connection.execute(
            """
            UPDATE sync_runs
            SET status='recovering',recovery_count=recovery_count+1,version=version+1
            WHERE status IN ('queued','running')
              AND heartbeat_at<CURRENT_TIMESTAMP - (? * interval '1 second')
            """,
            (RECOVERY_STALE_SECONDS,),
        )
        counts["runs"] = max(0, int(cursor.rowcount or 0))
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()
    counts["published"] = publish_pending_outbox(limit=100)
    return counts
