from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_SOURCE = (ROOT / "app.py").read_text(encoding="utf-8")
SETTINGS_TEMPLATE = (ROOT / "templates" / "settings.html").read_text(encoding="utf-8")


def test_unknown_products_api_returns_stable_source_list():
    assert "SELECT line_items, source FROM orders" in APP_SOURCE
    assert "unknown_products[key]['sources'].add(order['source'])" in APP_SOURCE
    assert "product['sources'] = sorted(product['sources'])" in APP_SOURCE


def test_settings_unknown_products_table_renders_all_source_sites():
    assert "<th>来源站点</th>" in SETTINGS_TEMPLATE
    assert "[...new Set(product.sources || [])].sort().map(source =>" in SETTINGS_TEMPLATE
    assert "new URL(source).hostname.replace(/^www\\./, '')" in SETTINGS_TEMPLATE
    assert "${escapeHtml(label)}" in SETTINGS_TEMPLATE
    assert 'colspan="5"' in SETTINGS_TEMPLATE
