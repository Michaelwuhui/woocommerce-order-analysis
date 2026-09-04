"""Small, deterministic product drill-down aggregates for the analysis page.

The product page already classifies every order line while building its sales
ranking. Keeping source prices and recent order samples during that pass is far
cheaper than repeating the full product-name parser whenever a user opens the
mapping modal.
"""

from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from urllib.parse import urlparse


RECENT_ORDER_LIMIT = 10
_CENT = Decimal("0.01")


def _row_value(row, key, default=None):
    try:
        value = row[key]
    except (KeyError, IndexError, TypeError):
        try:
            value = row.get(key, default)
        except AttributeError:
            return default
    return default if value is None else value


def _positive_decimal(value):
    if value in (None, ""):
        return None
    try:
        parsed = Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        return None
    if not parsed.is_finite() or parsed < 0:
        return None
    return parsed


def line_item_unit_price(item):
    """Return the order's observed unit price without binary-float maths."""

    explicit = _positive_decimal(item.get("price"))
    if explicit is not None:
        return explicit
    total = _positive_decimal(item.get("total"))
    quantity = _positive_decimal(item.get("quantity"))
    if total is None or quantity is None or quantity <= 0:
        return None
    return total / quantity


def _money_text(value):
    if value is None:
        return None
    return format(value.quantize(_CENT, rounding=ROUND_HALF_UP), ".2f")


def source_display_name(source):
    text = str(source or "").strip()
    parsed = urlparse(text)
    host = parsed.hostname or text.replace("https://", "").replace("http://", "").split("/")[0]
    return host[4:] if host.lower().startswith("www.") else host


def record_product_drilldown(product, order, item, quantity):
    """Add one classified order-line observation to a product aggregate."""

    source = str(_row_value(order, "source", "") or "").strip()
    if not source:
        return
    currency = str(_row_value(order, "currency", "") or "N/A").strip() or "N/A"
    order_id = str(_row_value(order, "id", "") or "")
    order_number = str(_row_value(order, "number", "") or order_id)
    order_date = str(_row_value(order, "date_created", "") or "")
    sort_key = (order_date, order_id)
    price = line_item_unit_price(item)

    source_rows = product.setdefault("_source_price_rows", {})
    row_key = (source, currency)
    row = source_rows.setdefault(
        row_key,
        {
            "source": source,
            "currency": currency,
            "latest_price": None,
            "latest_sort_key": ("", ""),
            "latest_date": "",
            "min_price": None,
            "max_price": None,
            "quantity": 0,
            "order_ids": set(),
        },
    )
    try:
        row["quantity"] += int(quantity or 0)
    except (TypeError, ValueError):
        pass
    if order_id:
        row["order_ids"].add(order_id)
    if price is not None:
        row["min_price"] = price if row["min_price"] is None else min(row["min_price"], price)
        row["max_price"] = price if row["max_price"] is None else max(row["max_price"], price)
        if sort_key >= row["latest_sort_key"]:
            row["latest_price"] = price
            row["latest_sort_key"] = sort_key
            row["latest_date"] = order_date

    samples = product.setdefault("_recent_order_rows", {})
    if order_id and order_id not in samples:
        samples[order_id] = {
            "order_number": order_number,
            "source": source,
            "date": order_date[:10],
            "_sort_key": sort_key,
        }
        if len(samples) > RECENT_ORDER_LIMIT:
            oldest_id = min(samples, key=lambda key: samples[key]["_sort_key"])
            samples.pop(oldest_id, None)


def finalize_product_drilldown(product, site_managers, recent_limit=RECENT_ORDER_LIMIT):
    """Convert private Decimal/set state into compact JSON-safe modal data."""

    source_rows = product.pop("_source_price_rows", {})
    finalized_sources = []
    for row in source_rows.values():
        source = row["source"]
        finalized_sources.append(
            {
                "source": source,
                "site": source_display_name(source),
                "manager": str(site_managers.get(source, "") or ""),
                "currency": row["currency"],
                "latest_price": _money_text(row["latest_price"]),
                "min_price": _money_text(row["min_price"]),
                "max_price": _money_text(row["max_price"]),
                "latest_date": str(row["latest_date"] or "")[:10],
                "order_count": len(row["order_ids"]),
                "quantity": int(row["quantity"]),
            }
        )
    product["source_prices"] = sorted(
        finalized_sources,
        key=lambda row: (row["site"].casefold(), row["currency"].casefold()),
    )

    samples = product.pop("_recent_order_rows", {})
    recent = sorted(samples.values(), key=lambda row: row["_sort_key"], reverse=True)
    product["recent_orders"] = [
        {
            "order_number": row["order_number"],
            "source": source_display_name(row["source"]),
            "manager": str(site_managers.get(row["source"], "") or ""),
            "date": row["date"],
        }
        for row in recent[: max(0, int(recent_limit))]
    ]
    return product
