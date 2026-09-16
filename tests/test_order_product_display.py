"""Order presentation regressions; load functions without production startup."""

import ast
import copy
import html
import json
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
TREE = ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))
HELPERS = {
    'extract_flavor_from_meta', 'extract_puffs_from_meta',
    'get_full_product_name', 'get_display_product_name',
}
NAMESPACE = {'html': html, 'parse_json_field': json.loads}
exec(compile(ast.Module(body=[
    node for node in TREE.body
    if isinstance(node, ast.FunctionDef) and node.name in HELPERS
], type_ignores=[]), 'app.py', 'exec'), NAMESPACE)
display_name = NAMESPACE['get_display_product_name']


@pytest.fixture
def merry_item():
    # Sanitized product data from order #1878; no customer details.
    return {
        'id': 16, 'product_id': 45, 'variation_id': 61,
        'sku': 'MB30K-GREEN-APPLE',
        'name': 'Merry Mi Blade 30000 — 21 smaków',
        'quantity': 1, 'total': '98.00',
        'meta_data': [{
            'key': 'pa_smak', 'value': 'green-apple',
            'display_key': 'Smak', 'display_value': 'Green Apple',
        }],
    }


def test_order_1878_uses_selected_flavor_without_changing_source(merry_item):
    original = copy.deepcopy(merry_item)
    assert display_name(merry_item) == 'Merry Mi Blade 30000 — 21 smaków - Green Apple'
    assert merry_item == original


@pytest.mark.parametrize('key', ['pa_smak', 'smak', 'pa_smaki', 'smaki',
                                'pa_flavour', 'flavour', 'pa_flavor', 'flavor',
                                'pa_taste', 'taste', 'pa_variant', 'variant', ' PA_SMAK '])
def test_existing_flavor_attribute_formats(merry_item, key):
    merry_item['meta_data'][0]['key'] = key
    assert display_name(merry_item).endswith(' - Green Apple')


def test_flavor_in_product_name_is_not_repeated(merry_item):
    merry_item['name'] = 'Merry Mi Blade 30000 - GREEN APPLE'
    assert display_name(merry_item) == merry_item['name']


def test_encoded_flavor_in_name_is_not_repeated(merry_item):
    merry_item['name'] = 'Merry Mi Blade 30000 - Apple &amp; Pear'
    merry_item['meta_data'][0]['display_value'] = 'Apple & Pear'
    assert display_name(merry_item) == 'Merry Mi Blade 30000 - Apple & Pear'


def test_uses_raw_value_when_display_value_is_missing(merry_item):
    del merry_item['meta_data'][0]['display_value']
    assert display_name(merry_item).endswith(' - green-apple')


@pytest.mark.parametrize('metadata', [None, {}, [], [None, {'key': None}],
                                    [{'key': '_reduced_stock', 'value': '1'}]])
def test_missing_flavor_preserves_product_title(merry_item, metadata):
    merry_item['meta_data'] = metadata
    assert display_name(merry_item) == merry_item['name']


def test_empty_or_structured_values_do_not_mask_valid_flavor(merry_item):
    merry_item['meta_data'][:0] = [
        {'key': 'flavor', 'value': ''},
        {'key': 'flavour', 'display_value': {'invalid': 'value'}},
        {'key': None, 'value': 'ignored'},
    ]
    assert display_name(merry_item).endswith(' - Green Apple')


@pytest.mark.parametrize('endpoint', ['orders', 'get_pending_orders', 'get_pending_outcome_orders',
                                     'print_shipping_label', 'process_shipped_order'])
def test_order_presentations_use_full_name_and_preserve_amounts(merry_item, endpoint):
    """Execute each real product-list expression, catching bare-name regressions."""
    function = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == endpoint)
    product_list = next(
        n for n in ast.walk(function) if isinstance(n, ast.ListComp)
        and isinstance(n.elt, ast.Dict)
        and any(isinstance(k, ast.Constant) and k.value == 'name' for k in n.elt.keys)
    )
    namespace = dict(NAMESPACE, items=[merry_item], line_items=[merry_item],
                     order={'line_items': json.dumps([merry_item])})
    product, = eval(compile(ast.Expression(product_list), 'app.py', 'eval'), namespace)
    assert product['name'] == 'Merry Mi Blade 30000 — 21 smaków - Green Apple'
    assert product.get('quantity', product.get('qty')) == 1
    if 'total' in product:
        assert product['total'] == 98.0


def test_details_add_display_name_and_preserve_original_line_item(merry_item):
    function = next(n for n in TREE.body if isinstance(n, ast.FunctionDef) and n.name == 'get_order_details')
    # Run the actual response enrichment loop, without network or DB access.
    loop = next(n for n in function.body if isinstance(n, ast.For) and any(
        isinstance(child, ast.Constant) and child.value == 'display_name'
        for child in ast.walk(n)
    ))
    order = {'line_items': [copy.deepcopy(merry_item)]}
    namespace = dict(NAMESPACE, order_dict=order)
    exec(compile(ast.Module(body=[loop], type_ignores=[]), 'app.py', 'exec'), namespace)
    item, = order['line_items']
    assert item.pop('display_name') == 'Merry Mi Blade 30000 — 21 smaków - Green Apple'
    assert item == merry_item


def test_detail_template_escapes_the_display_name():
    template = (ROOT / 'templates' / 'base.html').read_text(encoding='utf-8')
    assert "trackingEsc(item.display_name || item.name || '-')" in template


def test_shipping_templates_escape_decoded_product_names():
    template = (ROOT / 'templates' / 'shipping.html').read_text(encoding='utf-8')
    assert '${p.name}' not in template
    assert '${shipEscape(p.name)}' in template
