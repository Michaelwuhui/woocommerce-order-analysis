from customer_table_data import CUSTOMER_TABLE_FIELDS, customer_table_context


def test_table_retains_entire_authorized_result_and_searchable_identity_fields():
    customers = [
        {'name': f'Customer {i}', 'email': f'primary{i}@example.test',
         'identity_emails': [f'primary{i}@example.test', f'secondary{i}@example.test'],
         'phone': f'+485001{i:05}', 'total_spent_cny': 1000 - i,
         'missing_exchange_rates': ['XYZ@2026-09'] if i == 49 else [],
         'total_loss': 25 if i == 49 else 0, 'site_count': 2,
         'identity_email_count': 2, 'cluster_key': 'internal aggregation key'}
        for i in range(50)
    ]
    context = customer_table_context(customers)
    assert len(context['customers_payload']) == 50
    last = context['customers_payload'][-1]
    assert last['identity_emails'] == customers[-1]['identity_emails']
    assert last['phone'] == customers[-1]['phone']
    assert last['missing_exchange_rates'] == ['XYZ@2026-09']
    assert last['total_loss'] == 25
    assert last['identity_email_count'] == 2
    assert set(last) == set(CUSTOMER_TABLE_FIELDS)
    assert 'cluster_key' not in last
    assert 'cluster_key' in customers[-1]


def test_top_ten_uses_complete_cny_order_without_table_filter_or_page():
    customers = [{'name': f'#{i}', 'email': f'{i}@example.test',
                  'total_spent_cny': 10000 - i, 'total_spent': i * 100000}
                 for i in range(30)]
    context = customer_table_context(customers)
    assert [c['name'] for c in context['top_customers']] == [f'#{i}' for i in range(10)]
    assert context['top_customers'][0]['total_spent_cny'] == 10000
    assert set(context['top_customers'][0]) == {'name', 'email', 'total_spent_cny'}
    assert len(context['customers_payload']) == 30


def test_empty_scope_has_no_payload_or_chart_customers():
    assert customer_table_context([]) == {'customers_payload': [], 'top_customers': []}
