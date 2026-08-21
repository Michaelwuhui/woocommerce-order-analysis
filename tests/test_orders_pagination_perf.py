"""Regression tests for bounded order-list rendering and request-local FX caching."""

import os
import sqlite3
import sys

ROOT = "/www/wwwroot/woo-analysis"
sys.path.insert(0, ROOT)
os.chdir(ROOT)

import app as app_module


def _admin_client():
    conn = sqlite3.connect(os.path.join(ROOT, "woocommerce_orders.db"))
    row = conn.execute("SELECT id FROM users WHERE username='admin' LIMIT 1").fetchone()
    conn.close()
    assert row, "admin user not found"
    client = app_module.app.test_client()
    with client.session_transaction() as session:
        session["_user_id"] = str(row[0])
        session["_fresh"] = True
    return client


def test_orders_all_view_is_server_paginated():
    client = _admin_client()
    response = client.get("/orders?quick_date=one_year&month=all&per_page=50&page=1")
    assert response.status_code == 200
    html = response.get_data(as_text=True)
    assert html.count("showOrderDetail('") <= 50
    assert 'data-order-pagination="true"' in html
    assert "第 1 /" in html
    assert len(response.data) < 1_500_000


def test_orders_second_page_has_different_rows():
    client = _admin_client()
    first = client.get("/orders?quick_date=one_year&month=all&per_page=20&page=1").get_data(as_text=True)
    second = client.get("/orders?quick_date=one_year&month=all&per_page=20&page=2").get_data(as_text=True)
    import re
    first_ids = re.findall(r"showOrderDetail\('([^']+)'\)", first)
    second_ids = re.findall(r"showOrderDetail\('([^']+)'\)", second)
    assert len(first_ids) == 20
    assert len(second_ids) == 20
    assert set(first_ids).isdisjoint(second_ids)


def test_orders_month_totals_do_not_change_between_pages():
    import re

    client = _admin_client()
    base = "/orders?quick_date=one_year&month=2026-08&per_page=20&page={}"
    first = client.get(base.format(1)).get_data(as_text=True)
    second = client.get(base.format(2)).get_data(as_text=True)

    def totals_row(html):
        match = re.search(
            r'<span class="text-muted">整月合计</span>(.*?)</tr>', html, re.S
        )
        assert match, "monthly totals row not rendered"
        return re.sub(r"\s+", " ", match.group(1)).strip()

    assert totals_row(first) == totals_row(second)

    count_match = re.search(
        r'整月合计</span>\s*<span[^>]*>([\d,]+) 单</span>', first
    )
    assert count_match
    assert int(count_match.group(1).replace(",", "")) > 20


def test_get_cny_rate_reuses_lookup_within_request():
    original = app_module.get_db_connection
    calls = {"count": 0}

    def counted_connection():
        calls["count"] += 1
        return original()

    app_module.get_db_connection = counted_connection
    try:
        with app_module.app.test_request_context("/orders"):
            first = app_module.get_cny_rate("PLN", "2026-08")
            second = app_module.get_cny_rate("PLN", "2026-08")
            assert first == second
            assert calls["count"] == 1
    finally:
        app_module.get_db_connection = original
