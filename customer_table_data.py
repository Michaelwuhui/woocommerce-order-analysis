"""Presentation data for the deferred customer table, independent of Flask."""

CUSTOMER_TABLE_FIELDS = (
    'name', 'email', 'phone', 'identity_emails', 'identity_email_count',
    'identity_matched_by', 'source', 'site_count', 'site_list', 'tier',
    'successful_orders', 'total_spent_cny', 'missing_exchange_rates', 'total_loss',
    'undelivered_orders', 'problem_return_orders', 'total_orders', 'refusal_rate',
    'shipping_loss_total', 'product_loss_total', 'currency', 'last_order_date',
    'actions',
)


def customer_table_context(customers):
    """Keep the full authorized result searchable without rendering all its rows.

    The caller already sorted the complete identities by CNY spending. Charts
    use that full result, independently of the table's current page or search.
    """
    return {
        'customers_payload': [
            {field: customer.get(field) for field in CUSTOMER_TABLE_FIELDS}
            for customer in customers
        ],
        'top_customers': [
            {field: customer.get(field) for field in ('name', 'email', 'total_spent_cny')}
            for customer in customers[:10]
        ],
    }
