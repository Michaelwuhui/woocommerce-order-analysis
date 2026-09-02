"""Minimal, service-to-service read-only order API for the mail center.

The endpoint deliberately does not expose billing, shipping, line items, notes,
amounts, or WooCommerce credentials. It opens SQLite in read-only/query-only
mode and authenticates with a dedicated token file that must not contain a
WooCommerce consumer key or secret.
"""

from __future__ import annotations

import hashlib
import hmac
import os
import re
import db_backend as sqlite3
import stat
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, urlsplit

from flask import Blueprint, jsonify, request


mail_center_readonly_bp = Blueprint("mail_center_readonly", __name__)
_ORDER_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,63}$")


def _token_file() -> Path:
    configured = os.getenv("MAIL_CENTER_ORDER_API_TOKEN_FILE", "").strip()
    if not configured:
        credentials_dir = os.getenv("CREDENTIALS_DIRECTORY", "").strip()
        if credentials_dir:
            configured = str(Path(credentials_dir) / "mail-center-order-api-token")
    if not configured:
        raise RuntimeError("mail_center_order_api_token_not_configured")
    requested_path = Path(configured)
    if requested_path.is_symlink():
        raise RuntimeError("mail_center_order_api_token_symlink_forbidden")
    path = requested_path.resolve(strict=True)
    info = path.stat()
    if not stat.S_ISREG(info.st_mode):
        raise RuntimeError("mail_center_order_api_token_invalid_file")
    if os.name != "nt":
        if stat.S_IMODE(info.st_mode) & 0o077:
            raise RuntimeError("mail_center_order_api_token_permissions_too_broad")
        if info.st_uid != os.geteuid():
            raise RuntimeError("mail_center_order_api_token_owner_invalid")
    return path


def _expected_token() -> str:
    token = _token_file().read_text(encoding="utf-8").strip()
    if not 32 <= len(token) <= 256 or any(ch.isspace() for ch in token):
        raise RuntimeError("mail_center_order_api_token_invalid")
    lowered = token.lower()
    if lowered.startswith(("ck_", "cs_")) or "consumer_key" in lowered or "consumer_secret" in lowered:
        raise RuntimeError("mail_center_order_api_rejects_woocommerce_credentials")
    return token


def _authorized() -> bool:
    value = request.headers.get("Authorization", "")
    scheme, separator, supplied = value.partition(" ")
    if separator != " " or scheme.lower() != "bearer" or not supplied:
        return False
    try:
        expected = _expected_token()
    except (OSError, RuntimeError, UnicodeError):
        return False
    return hmac.compare_digest(supplied.encode("utf-8"), expected.encode("utf-8"))


def _database_path() -> Path:
    configured = os.getenv("MAIL_CENTER_ORDER_DB_PATH", "").strip()
    return Path(configured) if configured else Path(__file__).resolve().with_name("woocommerce_orders.db")


def _connect_readonly() -> sqlite3.Connection:
    if sqlite3.is_postgres_backend():
        conn = sqlite3.connect(
            os.getenv("WOO_SQLITE_PATH", "woocommerce_orders.db"),
            timeout=2.0,
        )
        conn.row_factory = sqlite3.Row
        return conn
    path = _database_path().resolve(strict=True)
    encoded_path = quote(path.as_posix(), safe="/:")
    conn = sqlite3.connect(f"file:{encoded_path}?mode=ro", uri=True, timeout=2.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA query_only=ON")
    conn.execute("PRAGMA busy_timeout=2000")
    return conn


def _allowed_sites() -> set[str]:
    return {
        value.strip().lower()
        for value in os.getenv("MAIL_CENTER_ORDER_ALLOWED_SITES", "").split(",")
        if value.strip()
    }


def _site_key(source: object) -> str:
    value = str(source or "").strip()
    parsed = urlsplit(value if "://" in value else f"https://{value}")
    return (parsed.hostname or "").lower()


def _public_order_url(order_id: str) -> str:
    base = os.getenv("MAIL_CENTER_ORDER_PUBLIC_BASE_URL", "").strip().rstrip("/")
    parsed = urlsplit(base)
    if (
        parsed.scheme != "https"
        or not parsed.hostname
        or parsed.username
        or parsed.password
        or parsed.query
        or parsed.fragment
        or parsed.path not in {"", "/"}
    ):
        raise RuntimeError("mail_center_order_public_base_url_invalid")
    return f"{base}/orders?search={quote(order_id, safe='')}"


@mail_center_readonly_bp.get("/internal/customer-service/orders/<order_id>")
def get_mail_center_order(order_id: str):
    if not _authorized():
        return jsonify({"error": "forbidden"}), 403
    if not _ORDER_ID_RE.fullmatch(order_id):
        return jsonify({"error": "invalid_order_id"}), 400
    try:
        with closing(_connect_readonly()) as conn:
            order = conn.execute(
                """
                SELECT id, number, source, status, date_modified_gmt, updated_at
                FROM orders
                WHERE id = ? OR number = ?
                ORDER BY CASE WHEN id = ? THEN 0 ELSE 1 END
                LIMIT 1
                """,
                (order_id, order_id, order_id),
            ).fetchone()
    except (OSError, sqlite3.Error):
        return jsonify({"error": "order_store_unavailable"}), 503
    if not order:
        return jsonify({"error": "order_not_found"}), 404

    site_key = _site_key(order["source"])
    allowed_sites = _allowed_sites()
    if not allowed_sites or site_key not in allowed_sites:
        return jsonify({"error": "site_forbidden"}), 403

    source_timestamp = order["date_modified_gmt"] or order["updated_at"] or "unknown"
    source_version = hashlib.sha256(
        f"{order['id']}|{order['status']}|{source_timestamp}".encode("utf-8")
    ).hexdigest()
    fetched_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    return jsonify(
        {
            "orderId": str(order["number"] or order["id"]),
            "siteKey": site_key,
            "status": str(order["status"] or "unknown"),
            "internalUrl": _public_order_url(str(order["number"] or order["id"])),
            "sourceVersion": source_version,
            "updatedAt": str(source_timestamp),
            "fetchedAt": fetched_at,
        }
    )
