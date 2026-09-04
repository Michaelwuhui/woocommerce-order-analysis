from decimal import Decimal
from pathlib import Path

from product_analysis_drilldown import (
    finalize_product_drilldown,
    line_item_unit_price,
    record_product_drilldown,
    source_display_name,
)


def _order(order_id, number, source, currency, created):
    return {
        "id": order_id,
        "number": number,
        "source": source,
        "currency": currency,
        "date_created": created,
    }


def test_line_item_price_prefers_observed_price_and_has_decimal_fallback():
    assert line_item_unit_price({"price": "69.90", "total": "100", "quantity": 2}) == Decimal("69.90")
    assert line_item_unit_price({"total": "100.00", "quantity": 4}) == Decimal("25.00")
    assert line_item_unit_price({"price": "not-a-price", "total": "9.99", "quantity": 0}) is None


def test_finalize_groups_prices_per_site_and_keeps_latest_and_range():
    product = {}
    site = "https://www.shop-one.example"
    record_product_drilldown(
        product,
        _order("1-10", "10", site, "PLN", "2026-09-01T10:00:00"),
        {"price": "54.90"},
        2,
    )
    record_product_drilldown(
        product,
        _order("1-11", "11", site, "PLN", "2026-09-02T10:00:00"),
        {"price": "59.90"},
        1,
    )

    result = finalize_product_drilldown(product, {site: "测试负责人"})

    assert result["source_prices"] == [
        {
            "source": site,
            "site": "shop-one.example",
            "manager": "测试负责人",
            "currency": "PLN",
            "latest_price": "59.90",
            "min_price": "54.90",
            "max_price": "59.90",
            "latest_date": "2026-09-02",
            "order_count": 2,
            "quantity": 3,
        }
    ]
    assert [row["order_number"] for row in result["recent_orders"]] == ["11", "10"]
    assert not any(key.startswith("_") for key in result)


def test_finalize_keeps_site_currency_separate_and_deduplicates_order_count():
    product = {}
    site = "https://shop.example/path"
    order = _order("2-7", "A-7", site, "CZK", "2026-09-03T12:00:00")
    record_product_drilldown(product, order, {"price": 399}, 1)
    record_product_drilldown(product, order, {"price": 399}, 2)
    record_product_drilldown(
        product,
        _order("3-8", "B-8", "https://other.example", "EUR", "2026-09-04T12:00:00"),
        {"total": "20.00", "quantity": 2},
        2,
    )

    result = finalize_product_drilldown(product, {}, recent_limit=1)

    assert [(row["site"], row["currency"]) for row in result["source_prices"]] == [
        ("other.example", "EUR"),
        ("shop.example", "CZK"),
    ]
    shop = next(row for row in result["source_prices"] if row["site"] == "shop.example")
    assert shop["order_count"] == 1
    assert shop["quantity"] == 3
    assert result["recent_orders"] == [
        {
            "order_number": "B-8",
            "source": "other.example",
            "manager": "",
            "date": "2026-09-04",
        }
    ]
    assert source_display_name(site) == "shop.example"


def test_recent_order_samples_are_bounded_before_finalize():
    product = {}
    for index in range(25):
        record_product_drilldown(
            product,
            _order(
                f"1-{index}",
                str(index),
                "https://shop.example",
                "PLN",
                f"2026-09-{index + 1:02d}T12:00:00",
            ),
            {"price": "10.00"},
            1,
        )

    assert len(product["_recent_order_rows"]) == 10
    result = finalize_product_drilldown(product, {})
    assert len(result["recent_orders"]) == 10
    assert result["recent_orders"][0]["order_number"] == "24"


def test_product_modal_uses_preloaded_drilldown_instead_of_slow_samples_api():
    template = (Path(__file__).resolve().parents[1] / "templates" / "products.html").read_text()

    assert "fetch('/api/products/samples?'" not in template
    assert "onclick=\"openProductMapping({{ loop.index0 }})\"" in template
    assert "product.source_prices || []" in template
    assert "各来源网站成交单价" in template
