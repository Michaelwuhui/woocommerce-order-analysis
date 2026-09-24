"""Exchange-rate helpers for the sales board.

An explicitly entered settlement rate represents the CNY actually received
after conversion losses. Otherwise the monthly receipt-weighted rate is used,
followed by the existing fallback and system rates.
"""

from collections import defaultdict
import math
import re
import db_backend as sqlite3


def parse_sales_board_rate_updates(data):
    """Validate a monthly rate save before any database row is changed.

    ``rate`` keeps its established fallback meaning. ``settlement_rate`` is
    optional so older clients cannot accidentally erase actual-settlement rates.
    Empty or zero values explicitly clear the corresponding saved rate.
    """
    if not isinstance(data, dict):
        raise ValueError("请求数据格式错误")
    month = data.get("month")
    if not isinstance(month, str) or not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", month):
        raise ValueError("月份格式应为 YYYY-MM")
    rates = data.get("rates", [])
    if not isinstance(rates, list):
        raise ValueError("汇率列表格式错误")

    parsed = []
    seen = set()
    for entry in rates:
        if not isinstance(entry, dict):
            raise ValueError("汇率条目格式错误")
        currency = entry.get("currency")
        if not isinstance(currency, str) or not re.fullmatch(r"[A-Za-z]{3}", currency):
            raise ValueError("币种必须是三个英文字母")
        currency = currency.upper()
        if currency == "CNY" or currency in seen:
            raise ValueError("币种重复或不能设置人民币汇率")
        seen.add(currency)

        update = {"currency": currency}
        for field in ("rate", "settlement_rate"):
            if field not in entry:
                continue
            raw = entry[field]
            if isinstance(raw, bool):
                raise ValueError(f"{currency} 的汇率必须是正数")
            if raw in (None, "", 0, "0"):
                update[field] = None
                continue
            try:
                value = float(raw)
            except (TypeError, ValueError):
                raise ValueError(f"{currency} 的汇率必须是正数") from None
            if not math.isfinite(value) or value <= 0:
                raise ValueError(f"{currency} 的汇率必须是有限正数")
            update[field] = value
        parsed.append(update)
    return month, parsed


def load_monthly_receipt_rates(conn, year_month):
    """Return monthly receipt-weighted CNY rates keyed by partner currency.

    Each valid receipt contributes its native amount and recorded CNY amount.
    When ``amount_cny`` is absent, the receipt's own exchange rate is used to
    derive it. Invalid or zero-value rows are ignored.
    """
    try:
        rows = conn.execute(
            """
            SELECT UPPER(COALESCE(p.currency, '')) AS currency,
                   p.name AS partner_name,
                   pr.amount_pln AS native_amount,
                   pr.exchange_rate_cny,
                   pr.amount_cny
            FROM partner_receipts pr
            JOIN partners p ON p.id = pr.partner_id
            WHERE substr(pr.receipt_date, 1, 7) = ?
            ORDER BY pr.receipt_date, pr.id
            """,
            (year_month,),
        ).fetchall()
    except sqlite3.OperationalError:
        return {}

    totals = defaultdict(
        lambda: {
            "native_amount": 0.0,
            "cny_amount": 0.0,
            "receipt_count": 0,
            "partner_names": [],
        }
    )
    for row in rows:
        currency = (row["currency"] or "").upper()
        try:
            native_amount = float(row["native_amount"] or 0)
            amount_cny = float(row["amount_cny"] or 0)
            receipt_rate = float(row["exchange_rate_cny"] or 0)
        except (TypeError, ValueError):
            continue
        if not currency or currency == "CNY" or native_amount <= 0:
            continue
        if amount_cny <= 0 and receipt_rate > 0:
            amount_cny = native_amount * receipt_rate
        if amount_cny <= 0:
            continue

        item = totals[currency]
        item["native_amount"] += native_amount
        item["cny_amount"] += amount_cny
        item["receipt_count"] += 1
        partner_name = row["partner_name"]
        if partner_name and partner_name not in item["partner_names"]:
            item["partner_names"].append(partner_name)

    result = {}
    for currency, item in totals.items():
        native_amount = item["native_amount"]
        if native_amount <= 0:
            continue
        result[currency] = {
            **item,
            "rate": item["cny_amount"] / native_amount,
        }
    return result


def resolve_sales_board_rate(
    currency,
    receipt_rates=None,
    custom_overrides=None,
    system_rate=0,
    settlement_rates=None,
):
    """Resolve one sales-board rate using the documented priority order."""
    currency = (currency or "").upper()
    if currency == "CNY":
        return 1.0, "system"

    receipt_rates = receipt_rates or {}
    custom_overrides = custom_overrides or {}
    settlement_rates = settlement_rates or {}
    try:
        settlement_rate = float(settlement_rates.get(currency) or 0)
    except (TypeError, ValueError):
        settlement_rate = 0
    if math.isfinite(settlement_rate) and settlement_rate > 0:
        return settlement_rate, "settlement"

    receipt = receipt_rates.get(currency) or {}
    try:
        receipt_rate = float(receipt.get("rate") or 0)
    except (TypeError, ValueError):
        receipt_rate = 0
    if receipt_rate > 0:
        return receipt_rate, "receipt"

    try:
        custom_rate = float(custom_overrides.get(currency) or 0)
    except (TypeError, ValueError):
        custom_rate = 0
    if custom_rate > 0:
        return custom_rate, "override"

    try:
        return float(system_rate or 0), "system"
    except (TypeError, ValueError):
        return 0.0, "system"
