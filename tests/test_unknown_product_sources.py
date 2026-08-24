import ast
import json
from pathlib import Path
from types import SimpleNamespace


ROOT = Path(__file__).resolve().parents[1]
APP_SOURCE = (ROOT / "app.py").read_text(encoding="utf-8")
SETTINGS_TEMPLATE = (ROOT / "templates" / "settings.html").read_text(encoding="utf-8")


def test_unknown_products_api_returns_stable_source_list():
    assert "SELECT line_items, source FROM orders" in APP_SOURCE
    assert "unknown_products[key]['sources'].add(order['source'])" in APP_SOURCE
    assert "product['sources'] = sorted(product['sources'])" in APP_SOURCE


def test_unknown_products_respects_manual_mapping_and_meta_puffs():
    tree = ast.parse(APP_SOURCE)
    endpoint = next(
        node for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name == "get_unknown_products"
    )
    endpoint.decorator_list = []

    orders = [{
        "source": "https://merrymipolska.pl",
        "line_items": json.dumps([
            {"name": "Merry Mi Blade 30000 — 21 smaków", "quantity": 1},
            {"name": "Mystery Product", "quantity": 2},
        ]),
    }]
    brands = [{"id": 26, "name": "Merrymi", "aliases": "[]"}]
    mappings = [{
        "raw_name": "Merry Mi Blade 30000 — 21 smaków",
        "source": None,
        "puff_count": 30000,
        "brand_name": "Merrymi",
    }]

    class FakeResult(list):
        def fetchall(self):
            return self

    class FakeConnection:
        def execute(self, query, _params=()):
            if "SELECT line_items, source FROM orders" in query:
                return FakeResult(orders)
            if "SELECT id, name, aliases FROM brands" in query:
                return FakeResult(brands)
            if "FROM product_mappings pm" in query:
                return FakeResult(mappings)
            raise AssertionError(query)

        def close(self):
            pass

    def full_product_name(item):
        puffs = 40000 if item["name"] == "Mystery Product" else None
        return item["name"], "", puffs

    namespace = {
        "get_db_connection": lambda: FakeConnection(),
        "request": SimpleNamespace(args={"days": "90", "limit": "50"}),
        "_active_status_cond": lambda: "1 = 1",
        "parse_json_field": json.loads,
        "get_full_product_name": full_product_name,
        "normalize_raw_name": lambda value: value,
        "parse_product_name": lambda *_args: {
            "brand": None, "puffs": None, "flavor": None,
        },
        "json": json,
        "jsonify": lambda value: value,
    }
    exec(compile(ast.Module(body=[endpoint], type_ignores=[]), "app.py", "exec"), namespace)

    result = namespace["get_unknown_products"]()

    assert [product["name"] for product in result] == ["Mystery Product"]
    assert result[0]["puffs"] == 40000
    assert result[0]["sources"] == ["https://merrymipolska.pl"]


def test_settings_unknown_products_table_renders_all_source_sites():
    assert "<th>来源站点</th>" in SETTINGS_TEMPLATE
    assert "[...new Set(product.sources || [])].sort().map(source =>" in SETTINGS_TEMPLATE
    assert "new URL(source).hostname.replace(/^www\\./, '')" in SETTINGS_TEMPLATE
    assert "${escapeHtml(label)}" in SETTINGS_TEMPLATE
    assert 'colspan="5"' in SETTINGS_TEMPLATE
