import json
import os
import shutil
import subprocess
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILTER_JS = os.path.join(ROOT, 'static', 'js', 'shipping_product_filter_compact1.js')
SHIPPING_TEMPLATE = os.path.join(ROOT, 'templates', 'shipping.html')


class ShippingProductFilterTests(unittest.TestCase):
    def test_shipping_page_exposes_searchable_product_controls(self):
        with open(SHIPPING_TEMPLATE, encoding='utf-8') as handle:
            template = handle.read()

        for expected in (
            'id="pendingProductFilter"',
            'id="pendingProductMenu"',
            'id="applyPendingProductBtn"',
            'id="clearPendingProductBtn"',
            'id="pendingProductStats"',
            "filename='js/shipping_product_filter_compact1.js'",
        ):
            self.assertIn(expected, template)
        self.assertNotIn('<datalist', template)
        self.assertNotIn('list="pendingProductOptions"', template)

    @unittest.skipUnless(shutil.which('node'), 'Node.js is required for the browser helper test')
    def test_selected_product_uses_exact_match_and_typed_search_is_fuzzy(self):
        orders = [
            {
                'id': 'one',
                'products': [
                    {'name': 'Merry Mi Blade - Dragon Fruit Ice', 'quantity': 2},
                    {'name': 'Another Product', 'quantity': 1},
                ],
            },
            {
                'id': 'two',
                'products': [
                    {'name': 'merry mi blade - dragon fruit ice', 'quantity': 1},
                    {'name': 'Merry Mi Blade - Dragon Fruit Ice Plus', 'quantity': 3},
                ],
            },
            {
                'id': 'three',
                'products': [{'name': 'Other Vape', 'quantity': 4}],
            },
        ]
        script = f"""
const assert = require('node:assert/strict');
const filter = require({json.dumps(FILTER_JS)});
const orders = {json.dumps(orders)};

const options = filter.productOptions(orders);
const exactOption = options.find(item => filter.normalize(item.name) === 'merry mi blade - dragon fruit ice');
assert.equal(exactOption.units, 3);
assert.equal(exactOption.orderCount, 2);
assert.equal(filter.searchOptions(options, 'm', 8).length, 0);
assert.equal(filter.searchOptions(options, 'Merry Mi Blade - Dragon Fruit Ice', 8)[0].name, exactOption.name);
assert.equal(filter.searchOptions(options, 'merry mi blade', 1).length, 1);

const all = filter.summarize(orders, '');
assert.equal(all.orderCount, 3);
assert.equal(all.units, 11);

const exact = filter.summarize(orders, 'Merry Mi Blade - Dragon Fruit Ice');
assert.equal(exact.exactMatch, true);
assert.equal(exact.orderCount, 2);
assert.equal(exact.units, 3);

const fuzzy = filter.summarize(orders, 'merry mi blade');
assert.equal(fuzzy.exactMatch, false);
assert.equal(fuzzy.orderCount, 2);
assert.equal(fuzzy.units, 6);

const missing = filter.summarize(orders, 'not available');
assert.equal(missing.orderCount, 0);
assert.equal(missing.units, 0);
"""
        completed = subprocess.run(
            [shutil.which('node'), '-e', script],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(0, completed.returncode, completed.stderr)


if __name__ == '__main__':
    unittest.main()
