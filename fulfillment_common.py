"""Shared infrastructure for the multi-warehouse fulfillment domain.

The fulfillment modules deliberately do not import ``app.py``.  They are used
both from Flask request handlers and from the standalone durable worker.
"""

from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime, timezone
from typing import Any, Iterable


DB_FILE = os.environ.get(
    "OMS_DB_FILE", os.environ.get("INV_DB_FILE", "woocommerce_orders.db")
)


def utcnow() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def get_conn(db_file: str | None = None) -> sqlite3.Connection:
    """Return a consistently configured SQLite connection.

    A 30 second busy timeout is required because the web app currently runs
    with four Gunicorn workers while the fulfillment worker writes to the same
    WAL database.
    """

    conn = sqlite3.connect(db_file or DB_FILE, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 30000")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def json_load(value: Any, default: Any = None) -> Any:
    if value is None or value == "":
        return default
    if isinstance(value, (dict, list)):
        return value
    try:
        return json.loads(value)
    except (TypeError, ValueError):
        return default


def json_dump(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def chunks(items: Iterable[Any], size: int):
    batch = []
    for item in items:
        batch.append(item)
        if len(batch) >= size:
            yield batch
            batch = []
    if batch:
        yield batch


_SENSITIVE_KEYS = {
    "salt",
    "authorization",
    "consumer_key",
    "consumer_secret",
    "password",
    "api_key",
    "token",
}

_PII_KEYS = {
    "tel",
    "phone",
    "email",
    "consignee",
    "billing",
    "shipping",
    "detail",
    "address",
    "address_1",
    "address_2",
}


def redact(value: Any, *, redact_pii: bool = True) -> Any:
    """Return a JSON-safe audit copy with credentials and PII removed."""

    if isinstance(value, dict):
        out = {}
        for key, item in value.items():
            normalized = str(key).lower()
            if normalized in _SENSITIVE_KEYS:
                out[key] = "[REDACTED]"
            elif redact_pii and normalized in _PII_KEYS:
                out[key] = "[PII]"
            else:
                out[key] = redact(item, redact_pii=redact_pii)
        return out
    if isinstance(value, list):
        return [redact(item, redact_pii=redact_pii) for item in value]
    return value


def table_exists(conn: sqlite3.Connection, name: str) -> bool:
    return bool(
        conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
        ).fetchone()
    )
