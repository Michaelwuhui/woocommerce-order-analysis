"""Pure helpers for order-detail multi-parcel presentation and lookup."""

from __future__ import annotations

import json
from typing import Any, Iterable


PENDING_SYNC_STATUS = "pending_sync"


def _dict(value: Any) -> dict:
    if isinstance(value, dict):
        return value
    try:
        return dict(value)
    except (TypeError, ValueError):
        return {}


def _jsonish(value: Any, default: Any) -> Any:
    if value in (None, ""):
        return default
    if isinstance(value, (dict, list)):
        return value
    if isinstance(value, str):
        try:
            return json.loads(value)
        except (TypeError, ValueError):
            return default
    return default


def _line_indexes(line_items: Iterable[dict]) -> tuple[dict[str, dict], dict[str, dict]]:
    by_item: dict[str, dict] = {}
    by_product: dict[str, dict] = {}
    for raw_line in line_items or []:
        line = _dict(raw_line)
        if line.get("id") is not None:
            by_item[str(line["id"])] = line
        for key in ("variation_id", "product_id"):
            value = str(line.get(key) or "").strip()
            if value and value != "0":
                by_product.setdefault(value, line)
    return by_item, by_product


def build_shipping_log_parcels(shipping_logs: Iterable[Any], line_items: Iterable[dict]) -> list[dict]:
    """Return every local parcel with human-readable item/quantity details."""
    by_item, by_product = _line_indexes(line_items)
    parcels: list[dict] = []
    for raw_row in shipping_logs or []:
        row = _dict(raw_row)
        number = str(row.get("tracking_number") or "").strip()
        if not number:
            continue
        items = []
        for raw_product in _jsonish(row.get("items_json"), []) or []:
            product = _dict(raw_product)
            item_id = str(product.get("item_id") or "").strip()
            product_id = str(product.get("product") or "").strip()
            line = by_item.get(item_id) or by_product.get(product_id) or {}
            try:
                qty = int(product.get("qty") or 0)
            except (TypeError, ValueError):
                qty = 0
            items.append({
                "item_id": item_id,
                "product": product_id,
                "name": str(line.get("name") or f"商品行 {item_id or product_id or '-'}"),
                "sku": str(line.get("sku") or ""),
                "qty": qty,
            })
        parcels.append({
            "id": row.get("id"),
            "tracking_number": number,
            "carrier_slug": str(row.get("carrier_slug") or ""),
            "shipped_at": row.get("shipped_at"),
            "status": str(row.get("status") or "shipped"),
            "is_partial": bool(row.get("is_partial")),
            "is_reship": bool(row.get("is_reship")),
            "reship_reason": str(row.get("reship_reason") or ""),
            "items": items,
        })
    return parcels


def partition_shipping_logs(shipping_logs: Iterable[Any]) -> tuple[list[dict], list[dict]]:
    """Separate confirmed parcels from attempts waiting for remote sync."""
    confirmed: list[dict] = []
    pending: list[dict] = []
    for raw_row in shipping_logs or []:
        row = _dict(raw_row)
        if str(row.get("status") or "shipped") == PENDING_SYNC_STATUS:
            pending.append(row)
        else:
            confirmed.append(row)
    return confirmed, pending


def is_pending_shipping_candidate(order: Any, shipping_logs: Iterable[Any]) -> bool:
    """Return whether an order still belongs in the normal shipping queue.

    A confirmed non-partial parcel means the order already had its final
    shipment.  Status drift in WooCommerce must not turn that parcel into a
    fake "continue shipping" action.  Pending-sync attempts stay visible so
    operators can retry them, while delivered/returned outcomes leave this
    queue and are handled by their dedicated workflows.
    """
    row = _dict(order)
    if str(row.get("status") or "") not in {
        "processing", "offline", "partial-shipped",
    }:
        return False
    if any(bool(row.get(key)) for key in (
        "is_undelivered", "is_problem_return", "delivery_confirmed",
    )):
        return False

    confirmed, pending = partition_shipping_logs(shipping_logs)
    if pending:
        return True
    return not any(not bool(parcel.get("is_partial")) for parcel in confirmed)


def find_duplicate_tracking(conn: Any, order_id: Any, carrier_slug: str, tracking_number: str) -> dict | None:
    """Return another order already holding the same carrier/tracking pair."""
    row = conn.execute(
        """SELECT sl.order_id, o.number
           FROM shipping_logs sl
           LEFT JOIN orders o ON o.id=sl.order_id
           WHERE sl.order_id<>?
             AND lower(trim(COALESCE(sl.carrier_slug, '')))=lower(trim(?))
             AND lower(trim(sl.tracking_number))=lower(trim(?))
           ORDER BY sl.id DESC LIMIT 1""",
        (order_id, carrier_slug, tracking_number),
    ).fetchone()
    return _dict(row) if row else None


def detect_tracking_format_rows(rows: Iterable[Any]) -> str:
    """Choose the dominant tracking format, using the newest tracked row to break ties."""
    counts = {"ast": 0, "villatheme": 0, "custom_lineitem": 0}
    newest = None
    for raw_row in rows or []:
        row = _dict(raw_row)
        meta_data = str(row.get("meta_data") or "")
        line_items = str(row.get("line_items") or "")
        kinds = []
        if "_wc_shipment_tracking_items" in meta_data:
            kinds.append("ast")
        if "_vi_wot_order_item_tracking_data" in line_items:
            kinds.append("villatheme")
        elif '"key":"tracking_number"' in line_items or '"key": "tracking_number"' in line_items:
            kinds.append("custom_lineitem")
        if kinds and newest is None:
            newest = kinds[0]
        for kind in kinds:
            counts[kind] += 1

    highest = max(counts.values())
    if not highest:
        return "unknown"
    winners = [kind for kind, count in counts.items() if count == highest]
    if newest in winners:
        return newest
    return winners[0]


def extract_tracking_candidates(
    meta_data: Any,
    line_items: Any,
    shipping_lines: Any,
    shipping_logs: Iterable[Any],
) -> list[dict]:
    """Collect and deduplicate all tracking numbers that belong to an order."""
    candidates: list[dict] = []

    def add(number: Any, provider: Any = "") -> None:
        number_text = str(number or "").strip()
        provider_text = str(provider or "").strip()
        if not number_text:
            return
        existing = next(
            (candidate for candidate in candidates
             if candidate["tracking_number"].casefold() == number_text.casefold()),
            None,
        )
        if existing:
            if provider_text and not existing["provider"]:
                existing["provider"] = provider_text
            return
        candidates.append({"tracking_number": number_text, "provider": provider_text})

    order_meta = _jsonish(meta_data, []) or []
    order_provider = ""
    for raw_meta in order_meta:
        meta = _dict(raw_meta)
        if meta.get("key") == "_tracking_provider":
            order_provider = str(meta.get("value") or "")
        if meta.get("key") == "_wc_shipment_tracking_items":
            for raw_tracking in _jsonish(meta.get("value"), []) or []:
                tracking = _dict(raw_tracking)
                add(tracking.get("tracking_number"), tracking.get("tracking_provider"))

    for raw_line in _jsonish(line_items, []) or []:
        line = _dict(raw_line)
        for raw_meta in line.get("meta_data") or []:
            meta = _dict(raw_meta)
            if meta.get("key") == "_vi_wot_order_item_tracking_data":
                for raw_tracking in _jsonish(meta.get("value"), []) or []:
                    tracking = _dict(raw_tracking)
                    add(
                        tracking.get("tracking_number"),
                        tracking.get("carrier_slug") or tracking.get("carrier_name"),
                    )
            elif meta.get("key") == "tracking_number":
                add(meta.get("value"))

    for raw_line in _jsonish(shipping_lines, []) or []:
        line = _dict(raw_line)
        for raw_meta in line.get("meta_data") or []:
            meta = _dict(raw_meta)
            if meta.get("key") == "tracking_number":
                add(meta.get("value"))

    for raw_meta in order_meta:
        meta = _dict(raw_meta)
        if meta.get("key") == "_tracking_number":
            add(meta.get("value"), order_provider)

    for raw_row in shipping_logs or []:
        row = _dict(raw_row)
        add(row.get("tracking_number"), row.get("carrier_slug"))
    return candidates
