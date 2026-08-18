"""Durable SQLite queue primitives for product clone jobs.

Product cloning can take several minutes because WooCommerce downloads images
and creates variations one by one.  Keeping that work outside the web request
prevents Gunicorn's request timeout from turning a successful/partial clone into
an HTML 502/504 response.
"""

from __future__ import annotations

import json
import uuid


TERMINAL_STATUSES = frozenset({"succeeded", "partial_failed", "failed", "interrupted"})


def init_product_clone_jobs(conn) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS product_clone_jobs (
            id TEXT PRIMARY KEY,
            source_site_id INTEGER NOT NULL,
            target_site_id INTEGER NOT NULL,
            product_ids_json TEXT NOT NULL,
            options_json TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'queued',
            total_count INTEGER NOT NULL DEFAULT 0,
            completed_count INTEGER NOT NULL DEFAULT 0,
            current_product_id INTEGER,
            success_count INTEGER NOT NULL DEFAULT 0,
            failed_count INTEGER NOT NULL DEFAULT 0,
            results_json TEXT NOT NULL DEFAULT '{"success":[],"failed":[]}',
            target_url TEXT NOT NULL DEFAULT '',
            created_by_id TEXT NOT NULL DEFAULT '',
            created_by_name TEXT NOT NULL DEFAULT '',
            worker_id TEXT,
            last_error TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            started_at TEXT,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            completed_at TEXT
        )
        """
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_product_clone_jobs_status_created "
        "ON product_clone_jobs(status, created_at)"
    )
    conn.commit()


def enqueue_clone_job(
    conn,
    *,
    source_site_id: int,
    target_site_id: int,
    product_ids: list[int],
    options: dict,
    target_url: str,
    created_by_id: str,
    created_by_name: str,
) -> dict:
    job_id = str(uuid.uuid4())
    conn.execute(
        """
        INSERT INTO product_clone_jobs (
            id, source_site_id, target_site_id, product_ids_json, options_json,
            total_count, target_url, created_by_id, created_by_name
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            job_id,
            source_site_id,
            target_site_id,
            json.dumps(product_ids, ensure_ascii=False),
            json.dumps(options, ensure_ascii=False),
            len(product_ids),
            target_url,
            str(created_by_id or ""),
            created_by_name or "",
        ),
    )
    conn.commit()
    return get_clone_job(conn, job_id)


def recover_interrupted_jobs(conn) -> int:
    """Fail previously-running jobs instead of retrying non-idempotent writes."""
    cur = conn.execute(
        """
        UPDATE product_clone_jobs
        SET status='interrupted',
            last_error='克隆 worker 曾中断；请检查目标站草稿后重新提交，系统不会自动重试以免重复创建。',
            current_product_id=NULL,
            completed_at=CURRENT_TIMESTAMP,
            updated_at=CURRENT_TIMESTAMP
        WHERE status='running'
        """
    )
    conn.commit()
    return int(cur.rowcount or 0)


def claim_clone_job(conn, worker_id: str) -> dict | None:
    conn.execute("BEGIN IMMEDIATE")
    row = conn.execute(
        "SELECT id FROM product_clone_jobs WHERE status='queued' ORDER BY created_at, id LIMIT 1"
    ).fetchone()
    if not row:
        conn.commit()
        return None
    job_id = row["id"] if hasattr(row, "keys") else row[0]
    cur = conn.execute(
        """
        UPDATE product_clone_jobs
        SET status='running', worker_id=?, started_at=CURRENT_TIMESTAMP,
            updated_at=CURRENT_TIMESTAMP, last_error=NULL
        WHERE id=? AND status='queued'
        """,
        (worker_id, job_id),
    )
    if cur.rowcount != 1:
        conn.rollback()
        return None
    conn.commit()
    return get_clone_job(conn, job_id)


def set_current_product(conn, job_id: str, product_id: int) -> None:
    conn.execute(
        """UPDATE product_clone_jobs
           SET current_product_id=?, updated_at=CURRENT_TIMESTAMP
           WHERE id=? AND status='running'""",
        (product_id, job_id),
    )
    conn.commit()


def save_clone_progress(conn, job_id: str, results: dict) -> None:
    success_count = len(results.get("success") or [])
    failed_count = len(results.get("failed") or [])
    conn.execute(
        """
        UPDATE product_clone_jobs
        SET completed_count=?, success_count=?, failed_count=?, results_json=?,
            current_product_id=NULL, updated_at=CURRENT_TIMESTAMP
        WHERE id=? AND status='running'
        """,
        (
            success_count + failed_count,
            success_count,
            failed_count,
            json.dumps(results, ensure_ascii=False),
            job_id,
        ),
    )
    conn.commit()


def finish_clone_job(conn, job_id: str, results: dict) -> None:
    success_count = len(results.get("success") or [])
    failed_count = len(results.get("failed") or [])
    status = "succeeded" if failed_count == 0 else "partial_failed"
    if success_count == 0 and failed_count:
        status = "failed"
    conn.execute(
        """
        UPDATE product_clone_jobs
        SET status=?, completed_count=?, success_count=?, failed_count=?,
            results_json=?, current_product_id=NULL, completed_at=CURRENT_TIMESTAMP,
            updated_at=CURRENT_TIMESTAMP
        WHERE id=?
        """,
        (
            status,
            success_count + failed_count,
            success_count,
            failed_count,
            json.dumps(results, ensure_ascii=False),
            job_id,
        ),
    )
    conn.commit()


def fail_clone_job(conn, job_id: str, error: str) -> None:
    conn.execute(
        """
        UPDATE product_clone_jobs
        SET status='failed', last_error=?, current_product_id=NULL,
            completed_at=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP
        WHERE id=?
        """,
        ((error or "未知错误")[:2000], job_id),
    )
    conn.commit()


def get_clone_job(conn, job_id: str) -> dict | None:
    row = conn.execute("SELECT * FROM product_clone_jobs WHERE id=?", (job_id,)).fetchone()
    return serialize_clone_job(row) if row else None


def serialize_clone_job(row) -> dict:
    raw = dict(row)
    for key, default in (
        ("product_ids_json", []),
        ("options_json", {}),
        ("results_json", {"success": [], "failed": []}),
    ):
        try:
            raw[key.removesuffix("_json")] = json.loads(raw.get(key) or "")
        except (TypeError, ValueError, json.JSONDecodeError):
            raw[key.removesuffix("_json")] = default
        raw.pop(key, None)
    raw["terminal"] = raw.get("status") in TERMINAL_STATUSES
    return raw
