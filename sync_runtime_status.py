"""SQLite-backed runtime status shared by all Gunicorn workers."""

from __future__ import annotations

import json
import time


def init_sync_runtime_status(conn) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS sync_runtime_status (
            status_id INTEGER PRIMARY KEY,
            status TEXT NOT NULL,
            message TEXT NOT NULL DEFAULT '',
            logs_json TEXT NOT NULL DEFAULT '[]',
            updated_at_epoch REAL NOT NULL,
            started_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            completed_at TEXT
        )
        """
    )
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
            status_id, status, message, logs_json, updated_at_epoch, completed_at
        ) VALUES (?, ?, ?, ?, ?, {completed_at_sql})
        ON CONFLICT(status_id) DO UPDATE SET
            status=excluded.status,
            message=excluded.message,
            logs_json=excluded.logs_json,
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
        "updated_at": updated_at,
        "stale_seconds": round(max(0.0, (now or time.time()) - updated_at), 1),
    }
