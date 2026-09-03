import ast
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class _RowsConnection:
    def __init__(self, rows):
        self.rows = rows
        self.query = ''
        self.params = []
        self.closed = False

    def execute(self, query, params):
        self.query = query
        self.params = params
        return self

    def fetchall(self):
        return self.rows

    def close(self):
        self.closed = True


def _load_report_aggregator(rows, rates):
    source = (ROOT / 'app.py').read_text(encoding='utf-8')
    tree = ast.parse(source)
    function = next(
        node for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name == 'get_orders_for_report'
    )
    connection = _RowsConnection(rows)
    namespace = {
        'get_db_connection': lambda: connection,
        '_revenue_status_cond': lambda: '1 = 1',
        'get_cny_rate': lambda currency, month: (rates.get((currency, month)), month),
        'datetime': datetime,
    }
    exec(compile(ast.Module(body=[function], type_ignores=[]), 'app.py', 'exec'), namespace)
    return namespace['get_orders_for_report'], connection


def test_report_gmv_uses_each_currency_month_rate():
    rows = [
        {
            'source': 'shop.example', 'period': '2026-08', 'rate_month': '2026-08', 'currency': 'PLN',
            'order_count': 2, 'total_amount': 100, 'net_amount': 90, 'success_net': 80,
        },
        {
            'source': 'shop.example', 'period': '2026-08', 'rate_month': '2026-08', 'currency': 'AUD',
            'order_count': 1, 'total_amount': 50, 'net_amount': 40, 'success_net': 30,
        },
    ]
    aggregate, connection = _load_report_aggregator(
        rows,
        {('PLN', '2026-08'): 2.0, ('AUD', '2026-08'): 4.0},
    )

    result = aggregate('2026-08-01', '2026-08-31')
    period = result['shop.example']['periods']['2026-08']

    assert period['order_count'] == 3
    assert period['total_amount'] == 150.0  # legacy native-currency field only
    assert period['gmv_cny'] == 400.0
    assert period['net_cny'] == 280.0
    assert period['missing_exchange_rates'] == []
    assert 'GROUP BY source, period, rate_month, currency' in connection.query
    assert connection.closed


def test_report_missing_rate_never_treats_native_amount_as_cny():
    rows = [{
        'source': 'shop.example', 'period': '2026-08', 'rate_month': '2026-08', 'currency': 'XYZ',
        'order_count': 1, 'total_amount': 500, 'net_amount': 450, 'success_net': 400,
    }]
    aggregate, _ = _load_report_aggregator(rows, {})

    period = aggregate('2026-08-01', '2026-08-31')['shop.example']['periods']['2026-08']

    assert period['gmv_cny'] is None
    assert period['net_cny'] is None
    assert period['missing_exchange_rates'] == ['XYZ@2026-08']


def test_week_spanning_months_uses_each_orders_calendar_month_rate():
    rows = [
        {
            'source': 'shop.example', 'period': '2026-08-31', 'rate_month': '2026-08', 'currency': 'PLN',
            'order_count': 1, 'total_amount': 100, 'net_amount': 90, 'success_net': 80,
        },
        {
            'source': 'shop.example', 'period': '2026-09-01', 'rate_month': '2026-09', 'currency': 'PLN',
            'order_count': 1, 'total_amount': 100, 'net_amount': 90, 'success_net': 80,
        },
    ]
    aggregate, connection = _load_report_aggregator(
        rows,
        {('PLN', '2026-08'): 2.0, ('PLN', '2026-09'): 3.0},
    )

    period = aggregate('2026-08-31', '2026-09-06', granularity='week')['shop.example']['periods']['2026-W36']

    assert period['gmv_cny'] == 500.0
    assert period['net_cny'] == 400.0
    assert "%Y-W%W" not in connection.query
    assert "strftime('%Y-%m-%d', date_created)" in connection.query


def test_iso_week_uses_iso_year_at_new_year_boundary():
    rows = [{
        'source': 'shop.example', 'period': '2027-01-01', 'rate_month': '2027-01', 'currency': 'PLN',
        'order_count': 1, 'total_amount': 10, 'net_amount': 9, 'success_net': 8,
    }]
    aggregate, _ = _load_report_aggregator(rows, {('PLN', '2027-01'): 2.0})

    periods = aggregate(
        '2027-01-01', '2027-01-01', granularity='week'
    )['shop.example']['periods']

    assert list(periods) == ['2026-W53']


def test_report_ui_uses_cny_gmv_for_amount_and_aov():
    template = (ROOT / 'templates' / 'report.html').read_text(encoding='utf-8')
    app_source = (ROOT / 'app.py').read_text(encoding='utf-8')

    assert "key: 'gmv_cny', label: 'GMV(CNY)'" in template
    assert "key: 'total_amount', label: '订单金额'" not in template
    assert 'totalCNY += siteData[month].gmv_cny || 0;' in template
    assert 'grand.gmv_cny / grand.order_count' in template
    assert "const REPORT_DATA_CACHE_VERSION = 'v7';" in template
    assert "report_cache_version = 'v7'" in app_source
