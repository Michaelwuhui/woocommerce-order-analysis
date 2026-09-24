"""PostgreSQL clean sync: verified remote absence, archive, then delete.

The job is recorded in sync_runs and delivered through the existing outbox.
It runs on the serial writer queue so it cannot outlive a Gunicorn worker or
race the ordinary order-sync writer. A retry rescans the site; a committed
archive/delete and the terminal site status are one transaction.
"""

from __future__ import annotations

import time
from datetime import datetime
from zoneinfo import ZoneInfo

from woocommerce import API

from celery_app import celery_app
from sync_service import (
    _event,
    _refresh_run_completion,
    get_connection,
    start_sync,
)


TIMEZONE = ZoneInfo("Asia/Hong_Kong")
MAX_PAGES = 1000
MAX_CANDIDATES = 20
ARCHIVED_ORDER_CHILDREN = {"order_notes", "order_note_sync_state", "order_notes_archive"}


class CleanSyncError(RuntimeError):
    pass


class CleanSyncCancelled(RuntimeError):
    pass


def _get(api, path: str, params=None):
    last_error = None
    for attempt in range(3):
        try:
            response = api.get(path, params=params)
            if response.status_code == 200 or response.status_code == 404:
                return response
            last_error = f"HTTP {response.status_code}: {path}"
            if response.status_code not in {429, 500, 502, 503, 504, 525}:
                break
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {str(exc)[:160]}"
        if attempt < 2:
            time.sleep(2 * (attempt + 1))
    raise CleanSyncError(last_error or f"读取远端失败: {path}")


def _id_page(api, page: int):
    response = _get(
        api,
        "orders",
        {
            "per_page": 100,
            "page": page,
            "_fields": "id",
            "status": "any",
            "orderby": "id",
            "order": "asc",
        },
    )
    if response.status_code != 200:
        raise CleanSyncError(f"订单列表第 {page} 页返回 HTTP {response.status_code}")
    try:
        rows = response.json()
        total = int(response.headers["X-WP-Total"])
        pages = int(response.headers["X-WP-TotalPages"])
    except (ValueError, KeyError, TypeError) as exc:
        raise CleanSyncError("WooCommerce 分页或 JSON 证据不完整") from exc
    if not isinstance(rows, list) or not all(
        isinstance(row, dict) and type(row.get("id")) is int for row in rows
    ):
        raise CleanSyncError(f"订单列表第 {page} 页格式异常")
    ids = [row["id"] for row in rows]
    if not (0 <= pages <= MAX_PAGES and 0 <= total <= max(1, pages) * 100):
        raise CleanSyncError("WooCommerce 分页计数超出安全范围")
    if (pages == 0 and (total or ids)) or (pages > 0 and page > pages):
        raise CleanSyncError("WooCommerce 分页信息互相矛盾")
    return ids, total, pages


def scan_remote_ids(api, on_page=lambda *_: None):
    first, total, pages = _id_page(api, 1)
    remote = list(first)
    on_page(1, pages, len(remote))
    for page in range(2, pages + 1):
        ids, page_total, page_count = _id_page(api, page)
        if (page_total, page_count) != (total, pages):
            raise CleanSyncError("WooCommerce 分页在扫描过程中发生变化")
        remote.extend(ids)
        on_page(page, pages, len(remote))
    repeated, repeated_total, repeated_pages = _id_page(api, 1)
    if (repeated, repeated_total, repeated_pages) != (first, total, pages):
        raise CleanSyncError("WooCommerce 首页在扫描过程中发生变化")
    if len(remote) != total or len(set(remote)) != total or remote != sorted(remote):
        raise CleanSyncError("WooCommerce 订单列表不完整、重复或排序不稳定")
    return set(remote), total, pages


def confirm_remote_deleted(api, woo_id: int) -> bool:
    detail = _get(api, f"orders/{woo_id}")
    if detail.status_code == 200:
        return False
    try:
        body = detail.json()
    except ValueError as exc:
        raise CleanSyncError(f"订单 {woo_id} 的 404 缺少结构化证据") from exc
    if not isinstance(body, dict) or body.get("code") != "woocommerce_rest_shop_order_invalid_id":
        raise CleanSyncError(f"订单 {woo_id} 的 404 不是已删除订单证据")
    trash = _get(
        api,
        "orders",
        {"status": "trash", "include": woo_id, "per_page": 100},
    )
    if trash.status_code != 200:
        raise CleanSyncError(f"订单 {woo_id} 的垃圾箱状态无法核实")
    try:
        rows = trash.json()
    except ValueError as exc:
        raise CleanSyncError("垃圾箱列表 JSON 无效") from exc
    if not isinstance(rows, list) or not all(isinstance(row, dict) for row in rows):
        raise CleanSyncError("垃圾箱列表格式异常")
    return not any(row.get("id") == woo_id for row in rows)


def _heartbeat(connection, run_id: str, site_id: int, page: int, fetched: int):
    cancelled = connection.execute(
        "SELECT cancellation_requested FROM sync_runs WHERE run_id=?",
        (run_id,),
    ).fetchone()
    if not cancelled or bool(cancelled["cancellation_requested"]):
        raise CleanSyncCancelled()
    connection.execute(
        """UPDATE sync_site_progress SET current_page=?,fetched_count=?,
           heartbeat_at=CURRENT_TIMESTAMP,version=version+1
           WHERE run_id=? AND site_id=?""",
        (page, fetched, run_id, site_id),
    )
    connection.execute(
        "UPDATE sync_runs SET heartbeat_at=CURRENT_TIMESTAMP,current_site_id=? WHERE run_id=?",
        (site_id, run_id),
    )
    connection.commit()


def _finish_site(connection, run_id: str, site_id: int, status: str, message: str):
    connection.execute(
        """UPDATE sync_site_progress SET status=?,error_message=?,
           finished_at=CURRENT_TIMESTAMP,heartbeat_at=CURRENT_TIMESTAMP,
           version=version+1 WHERE run_id=? AND site_id=?
           AND status NOT IN ('success','error','auth_error','cancelled')""",
        (status, message[:2000] if status == "error" else None, run_id, site_id),
    )
    _event(
        connection,
        run_id,
        "clean_site_" + status,
        message[:2000],
        site_id=site_id,
        level="error" if status == "error" else "info",
    )
    _refresh_run_completion(connection, run_id)
    connection.commit()


def _quote(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def _archive_columns(connection, source_table="orders", archive_table="orders_archive"):
    rows = connection.execute(
        """SELECT table_name,column_name FROM information_schema.columns
           WHERE table_schema='public' AND table_name IN (?,?)
           ORDER BY ordinal_position""",
        (source_table, archive_table),
    ).fetchall()
    source_columns = [row["column_name"] for row in rows if row["table_name"] == source_table]
    archive_columns = {row["column_name"] for row in rows if row["table_name"] == archive_table}
    if not source_columns or not set(source_columns) <= archive_columns or not {
        "archived_at", "archive_reason"
    } <= archive_columns:
        raise CleanSyncError(f"{archive_table} 归档表结构不完整，已停止清理")
    return ", ".join(_quote(name) for name in source_columns)


def _reference_tables(connection):
    rows = connection.execute(
        """SELECT table_name FROM information_schema.columns
           WHERE table_schema='public' AND column_name='order_id'"""
    ).fetchall()
    return sorted({row["table_name"] for row in rows} - ARCHIVED_ORDER_CHILDREN)


def _apply_candidates(connection, run_id, site_id, source, candidates, remote_count, pages, skipped):
    """Commit this site's archive, deletion and success status atomically."""
    try:
        connection.execute("SET LOCAL lock_timeout = '5s'")
        run = connection.execute(
            "SELECT mode,cancellation_requested FROM sync_runs WHERE run_id=? FOR UPDATE",
            (run_id,),
        ).fetchone()
        if not run or run["mode"] != "clean" or bool(run["cancellation_requested"]):
            raise CleanSyncCancelled()
        names = _archive_columns(connection)
        note_names = _archive_columns(connection, "order_notes", "order_notes_archive")
        reference_tables = _reference_tables(connection)
        removed = 0
        blocked = 0
        for original in candidates:
            order_id = original["id"]
            current = connection.execute(
                "SELECT id,source,woo_id,status FROM orders WHERE id=? FOR UPDATE",
                (order_id,),
            ).fetchone()
            if not current or any(current[key] != original[key] for key in ("id", "source", "woo_id", "status")):
                raise CleanSyncError(f"订单 {order_id} 在核验后发生变化")
            if current["source"] != source:
                raise CleanSyncError(f"订单 {order_id} 的站点已变化")
            existing = connection.execute(
                "SELECT 1 FROM orders_archive WHERE id=?", (order_id,)
            ).fetchone()
            if existing:
                raise CleanSyncError(f"订单 {order_id} 已有归档记录")
            references = [
                table for table in reference_tables
                if connection.execute(
                    f"SELECT 1 FROM {_quote(table)} WHERE order_id=? LIMIT 1",
                    (order_id,),
                ).fetchone()
            ]
            if references:
                blocked += 1
                _event(
                    connection, run_id, "clean_order_blocked",
                    f"订单 {order_id} 有业务关联，保留原记录",
                    site_id=site_id, level="warning",
                    details={"order_id": order_id, "tables": references},
                )
                continue
            note_count = connection.execute(
                "SELECT COUNT(*) FROM order_notes WHERE order_id=?", (order_id,)
            ).fetchone()[0]
            archived_notes = connection.execute(
                f"INSERT INTO order_notes_archive ({note_names}, archived_at, archive_reason) "
                f"SELECT {note_names}, CURRENT_TIMESTAMP, ? FROM order_notes WHERE order_id=?",
                ("orphaned_remote_deleted", order_id),
            )
            if archived_notes.rowcount != note_count:
                raise CleanSyncError(f"订单 {order_id} 的备注归档不完整")
            deleted_notes = connection.execute(
                "DELETE FROM order_notes WHERE order_id=?", (order_id,)
            )
            if deleted_notes.rowcount != note_count:
                raise CleanSyncError(f"订单 {order_id} 的备注移出不完整")
            connection.execute(
                "DELETE FROM order_note_sync_state WHERE order_id=?", (order_id,)
            )
            inserted = connection.execute(
                f"INSERT INTO orders_archive ({names}, archived_at, archive_reason) "
                f"SELECT {names}, CURRENT_TIMESTAMP, ? FROM orders WHERE id=?",
                ("orphaned_remote_deleted", order_id),
            )
            if inserted.rowcount != 1:
                raise CleanSyncError(f"订单 {order_id} 归档失败")
            deleted = connection.execute(
                "DELETE FROM orders WHERE id=? AND source=? AND woo_id=?",
                (order_id, source, original["woo_id"]),
            )
            if deleted.rowcount != 1:
                raise CleanSyncError(f"订单 {order_id} 移出失败")
            removed += 1
        connection.execute(
            """UPDATE sync_site_progress
               SET status='success',current_page=?,total_pages=?,fetched_count=?,
                   written_count=?,changed_count=?,error_message=NULL,
                   finished_at=CURRENT_TIMESTAMP,heartbeat_at=CURRENT_TIMESTAMP,
                   version=version+1 WHERE run_id=? AND site_id=?""",
            (pages, pages, remote_count, removed, removed, run_id, site_id),
        )
        connection.execute(
            """UPDATE sync_runs
               SET total_pages=total_pages+?,completed_pages=completed_pages+?,
                   fetched_orders=fetched_orders+?,written_orders=written_orders+?,
                   changed_orders=changed_orders+?,heartbeat_at=CURRENT_TIMESTAMP,
                   current_site_id=?,version=version+1 WHERE run_id=?""",
            (pages, pages, remote_count, removed, removed, site_id, run_id),
        )
        _event(
            connection, run_id, "clean_site_completed",
            f"远端 {remote_count} 单，候选 {len(candidates) + skipped} 单，归档移出 {removed} 单，保留 {skipped + blocked} 单",
            site_id=site_id,
            details={
                "remote_count": remote_count, "candidate_count": len(candidates) + skipped,
                "removed": removed, "retained": skipped + blocked,
            },
        )
        _refresh_run_completion(connection, run_id)
        connection.commit()
        return {"removed": removed, "retained": skipped + blocked}
    except Exception:
        connection.rollback()
        raise


def run_clean_site(payload: dict):
    run_id, site_id = str(payload["run_id"]), int(payload["site_id"])
    connection = get_connection()
    locked = False
    try:
        row = connection.execute(
            """SELECT r.mode,r.status AS run_status,r.cancellation_requested,
                      p.status AS site_status,s.url,s.consumer_key,s.consumer_secret
               FROM sync_site_progress p JOIN sync_runs r ON r.run_id=p.run_id
               JOIN sites s ON s.id=p.site_id
               WHERE p.run_id=? AND p.site_id=? FOR UPDATE OF r,p""",
            (run_id, site_id),
        ).fetchone()
        if not row or row["mode"] != "clean" or row["site_status"] in {
            "success", "error", "auth_error", "cancelled"
        }:
            connection.rollback()
            return {"skipped": True}
        if bool(row["cancellation_requested"]):
            _finish_site(connection, run_id, site_id, "cancelled", "清理同步已取消")
            return {"cancelled": True}
        locked = bool(connection.execute(
            "SELECT pg_try_advisory_lock(hashtextextended(?,0))",
            (f"woo-sync-site:{site_id}",),
        ).fetchone()[0])
        if not locked:
            connection.execute(
                """UPDATE sync_task_outbox SET status='pending',
                   available_at=CURRENT_TIMESTAMP + interval '15 seconds',
                   updated_at=CURRENT_TIMESTAMP
                   WHERE dedupe_key=?""",
                (f"clean:{run_id}:{site_id}",),
            )
            connection.commit()
            return {"busy": True}
        connection.execute(
            """UPDATE sync_site_progress SET status='fetching',
               started_at=COALESCE(started_at,CURRENT_TIMESTAMP),
               heartbeat_at=CURRENT_TIMESTAMP,version=version+1
               WHERE run_id=? AND site_id=?""",
            (run_id, site_id),
        )
        connection.execute(
            """UPDATE sync_runs SET status='running',current_site_id=?,
               started_at=COALESCE(started_at,CURRENT_TIMESTAMP),
               heartbeat_at=CURRENT_TIMESTAMP,version=version+1
               WHERE run_id=? AND status IN ('queued','running','recovering')""",
            (site_id, run_id),
        )
        _event(connection, run_id, "clean_site_started", "核对远端订单列表", site_id=site_id)
        connection.commit()
        source = str(row["url"]).strip()
        api = API(
            url=source,
            consumer_key=row["consumer_key"],
            consumer_secret=row["consumer_secret"],
            version="wc/v3",
            timeout=30,
            user_agent="WooCommerce API Client-Python/3.0.0",
        )
        local_rows = [
            dict(value) for value in connection.execute(
                "SELECT id,source,woo_id,status FROM orders WHERE source=?",
                (source,),
            ).fetchall()
        ]
        connection.commit()
        remote_ids, remote_count, pages = scan_remote_ids(
            api,
            lambda page, _total, fetched: _heartbeat(
                connection, run_id, site_id, page, fetched
            ),
        )
        candidates = [
            order for order in local_rows
            if order["woo_id"] is not None and int(order["woo_id"]) not in remote_ids
        ]
        if len(candidates) > MAX_CANDIDATES or (
            len(candidates) >= 5
            and len(candidates) * 10 >= max(1, len(local_rows))
        ):
            raise CleanSyncError(
                f"站点候选 {len(candidates)}/{len(local_rows)} 过多，疑似远端回滚，未删除任何订单"
            )
        confirmed = []
        skipped = 0
        for order in candidates:
            _heartbeat(connection, run_id, site_id, pages, remote_count)
            if confirm_remote_deleted(api, int(order["woo_id"])):
                confirmed.append(order)
            else:
                skipped += 1
        connection.execute(
            """UPDATE sync_site_progress SET status='writing',
               heartbeat_at=CURRENT_TIMESTAMP,version=version+1
               WHERE run_id=? AND site_id=?""",
            (run_id, site_id),
        )
        connection.commit()
        return _apply_candidates(
            connection, run_id, site_id, source, confirmed, remote_count, pages, skipped
        )
    except CleanSyncCancelled:
        connection.rollback()
        _finish_site(connection, run_id, site_id, "cancelled", "清理同步已取消")
        return {"cancelled": True}
    except Exception as exc:
        connection.rollback()
        _finish_site(
            connection, run_id, site_id, "error",
            f"清理同步失败，未移出本站订单：{type(exc).__name__}: {str(exc)[:1500]}",
        )
        return {"error": str(exc)[:500]}
    finally:
        if locked:
            try:
                connection.execute(
                    "SELECT pg_advisory_unlock(hashtextextended(?,0))",
                    (f"woo-sync-site:{site_id}",),
                )
                connection.commit()
            except Exception:
                connection.rollback()
        connection.close()


@celery_app.task(name="woo_sync.clean_site", acks_late=True, reject_on_worker_lost=True)
def clean_site(payload: dict):
    return run_clean_site(payload)


def _clean_due():
    connection = get_connection()
    try:
        rows = connection.execute(
            """SELECT key,value FROM settings WHERE key IN
               ('clean_sync_enabled','clean_sync_day','clean_sync_hour','clean_sync_minute')"""
        ).fetchall()
        settings = {row["key"]: row["value"] for row in rows}
        enabled = str(settings.get("clean_sync_enabled", "false")).strip().lower() == "true"
        try:
            day = int(settings.get("clean_sync_day", "0"))
            hour = int(settings.get("clean_sync_hour", "4"))
            minute = int(settings.get("clean_sync_minute", "0"))
        except (TypeError, ValueError) as exc:
            raise CleanSyncError("清理同步定时配置无效，未启动任务") from exc
        if not (0 <= day <= 6 and 0 <= hour <= 23 and 0 <= minute <= 59):
            raise CleanSyncError("清理同步定时配置越界，未启动任务")
        now = datetime.now(TIMEZONE)
        schedule = {"enabled": enabled, "day": day, "hour": hour, "minute": minute}
        if not enabled or (now.weekday() + 1) % 7 != day or now < now.replace(
            hour=hour, minute=minute, second=0, microsecond=0
        ):
            return False, schedule
        already = connection.execute(
            """SELECT EXISTS(SELECT 1 FROM sync_runs
               WHERE mode='clean' AND created_by='celery-beat:clean'
               AND timezone('Asia/Hong_Kong',created_at)::date=?::date)""",
            (now.date(),),
        ).fetchone()
        return not bool(already[0]), schedule
    finally:
        connection.close()


@celery_app.task(name="woo_sync.schedule_clean", acks_late=True, reject_on_worker_lost=True)
def schedule_clean():
    due, schedule = _clean_due()
    if not due:
        return {"created": False, "schedule": schedule}
    status, created = start_sync(
        mode="clean", created_by="celery-beat:clean",
        params={"per_page": 100, "notes_per_page": 0},
    )
    return {"run_id": status["run_id"], "created": created, "schedule": schedule}
