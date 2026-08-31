from customer_spending import customer_spending_cny_by_email


def test_customer_spending_uses_each_month_and_currency_rate():
    rows = [
        {'email': ' A@example.com ', 'currency': 'PLN', 'year_month': '2026-07', 'total_spent': 100},
        {'email': 'a@example.com', 'currency': 'AUD', 'year_month': '2026-08', 'total_spent': 100},
        {'email': 'b@example.com', 'currency': 'PLN', 'year_month': '2026-08', 'total_spent': 150},
    ]
    rates = {
        ('PLN', '2026-07'): 2.0,
        ('AUD', '2026-08'): 4.0,
        ('PLN', '2026-08'): 1.9,
    }

    totals = customer_spending_cny_by_email(
        rows, lambda currency, month: (rates.get((currency, month)), month)
    )

    assert totals['a@example.com']['total_spent_cny'] == 600.0
    assert totals['b@example.com']['total_spent_cny'] == 285.0


def test_customer_spending_marks_missing_rate_instead_of_treating_native_amount_as_cny():
    rows = [
        {'email': 'a@example.com', 'currency': 'XYZ', 'year_month': '2026-08', 'total_spent': 500},
    ]

    totals = customer_spending_cny_by_email(
        rows, lambda currency, month: (None, None)
    )

    assert totals['a@example.com']['total_spent_cny'] == 0
    assert totals['a@example.com']['missing_exchange_rates'] == {'XYZ@2026-08'}
