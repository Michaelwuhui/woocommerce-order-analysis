"""SKU helpers for explicit clone-as-new product operations."""

from __future__ import annotations

import hashlib
import re
import uuid
from datetime import datetime


MAX_WC_SKU_LENGTH = 100


def normalize_clone_suffix(value: str) -> str:
    """Return a short, WooCommerce-safe uppercase suffix."""
    cleaned = re.sub(r"[^A-Z0-9]+", "-", (value or "").upper()).strip("-")
    return cleaned[:32].rstrip("-")


def make_clone_suffix(now: datetime | None = None, token: str | None = None) -> str:
    """Build a unique suffix that is stable once persisted in a clone job."""
    now = now or datetime.now()
    token = normalize_clone_suffix(token or uuid.uuid4().hex[:6])[:6]
    return f"NEW-{now:%Y%m%d}-{token}"


def build_clone_sku(source_sku: str, suffix: str, *, fallback: str) -> str:
    """Append ``suffix`` while keeping the final SKU unique and <= 100 chars."""
    suffix = normalize_clone_suffix(suffix)
    if not suffix:
        raise ValueError("clone SKU suffix is required")

    base = (source_sku or fallback or "CLONE").strip()
    available = MAX_WC_SKU_LENGTH - len(suffix) - 1
    if available < 10:
        raise ValueError("clone SKU suffix is too long")
    if len(base) > available:
        digest = hashlib.sha1(base.encode("utf-8")).hexdigest()[:8].upper()
        base = f"{base[:available - len(digest) - 1]}-{digest}"
    return f"{base}-{suffix}"
