"""SQLite-backed runtime status shared by all Gunicorn workers."""

from __future__ import annotations

import json
import secrets
import time


_MIN_BROWSER_SAFE_STATUS_ID = 1_000_000_000_000_000
_MAX_BROWSER_SAFE_STATUS_ID = 9_000_000_000_000_000


def new_sync_runtime_status_id() -> int:
    """Return a per-run ID that remains exact in JavaScript and SQLite."""
    return _MIN_BROWSER_SAFE_STATUS_ID + secrets.randbelow(
        _MAX_BROWSER_SAFE_STATUS_ID - _MIN_BROWSER_SAFE_STATUS_ID
    )


def init_sync_runtime_status(conn) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS sync_runtime_status (
            status_id INTEGER PRIMARY KEY,
            status TEXT NOT NULL,
            message TEXT NOT NULL DEFAULT '',
            logs_json TEXT NOT NULL DEFAULT '[]',
            progress REAL NOT NULL DEFAULT 0,
            updated_at_epoch REAL NOT NULL,
            started_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            completed_at TEXT
        )
        """
    )
    columns = {
        row[1] for row in conn.execute("PRAGMA table_info(sync_runtime_status)")
    }
    if "progress" not in columns:
        try:
            conn.execute(
                "ALTER TABLE sync_runtime_status "
                "ADD COLUMN progress REAL NOT NULL DEFAULT 0"
            )
        except Exception as exc:
            if "duplicate column" not in str(exc).lower():
                raise
    conn.commit()


def save_sync_runtime_status(conn, status_id: int, entry: dict) -> None:
    status = str(entry.get("status") or "unknown")
    updated_at = float(entry.get("updated_at") or time.time())
    completed_at_sql = (
        "CURRENT_TIMESTAMP"
        if status in {"success", "error"}
        else "NULL"
    )
    conn.execute(
        f"""
        INSERT INTO sync_runtime_status (
            status_id, status, message, logs_json, progress,
            updated_at_epoch, completed_at
        ) VALUES (?, ?, ?, ?, ?, ?, {completed_at_sql})
        ON CONFLICT(status_id) DO UPDATE SET
            status=excluded.status,
            message=excluded.message,
            logs_json=excluded.logs_json,
            progress=excluded.progress,
            updated_at_epoch=excluded.updated_at_epoch,
            started_at=CASE
                WHEN excluded.status='running'
                     AND sync_runtime_status.status!='running'
                THEN CURRENT_TIMESTAMP
                ELSE sync_runtime_status.started_at
            END,
            completed_at={completed_at_sql}
        """,
        (
            int(status_id),
            status,
            str(entry.get("message") or ""),
            json.dumps(entry.get("logs") or [], ensure_ascii=False),
            float(entry.get("progress") or 0),
            updated_at,
        ),
    )
    conn.commit()


def load_sync_runtime_status(conn, status_id: int, *, now: float | None = None) -> dict | None:
    row = conn.execute(
        "SELECT * FROM sync_runtime_status WHERE status_id=?", (int(status_id),)
    ).fetchone()
    if not row:
        return None
    raw = dict(row)
    try:
        logs = json.loads(raw.get("logs_json") or "[]")
    except (TypeError, ValueError, json.JSONDecodeError):
        logs = []
    updated_at = float(raw.get("updated_at_epoch") or 0)
    return {
        "status": raw.get("status") or "unknown",
        "message": raw.get("message") or "",
        "logs": logs if isinstance(logs, list) else [],
        "progress": float(raw.get("progress") or 0),
        "updated_at": updated_at,
        "stale_seconds": round(max(0.0, (now or time.time()) - updated_at), 1),
    }
