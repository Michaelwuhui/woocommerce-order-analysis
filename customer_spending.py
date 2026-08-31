"""Customer spending helpers that do not depend on Flask application state."""

from datetime import datetime


def customer_spending_cny_by_email(rows, rate_getter):
    """Fold monthly native-currency spending rows into CNY totals by email."""
    totals = {}
    for row in rows:
        email = (row['email'] or '').strip().lower()
        if not email:
            continue

        amount = float(row['total_spent'] or 0)
        currency = (row['currency'] or 'PLN').strip().upper()
        year_month = row['year_month'] or datetime.now().strftime('%Y-%m')
        item = totals.setdefault(email, {
            'total_spent_cny': 0.0,
            'missing_exchange_rates': set(),
        })
        rate, _ = rate_getter(currency, year_month)
        if rate is None:
            if amount:
                item['missing_exchange_rates'].add(f'{currency}@{year_month}')
            continue
        item['total_spent_cny'] += amount * float(rate)

    return totals
