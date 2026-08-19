import json
import sqlite3
import unittest

from inv_mapping_service import (
    apply_safe_mappings,
    mapping_detail,
    mapping_overview,
    scan_site_catalog,
    serving_sites,
    warehouse_rows,
)
from inv_migrations import up_019


class FakeResponse:
    def __init__(self, payload, status=200):
        self._payload = payload
        self.status_code = status
        self.headers = {}
        self.text = json.dumps(payload)

    def json(self):
        return self._payload


class FakeWooSession:
    def get(self, url, **kwargs):
        if url.endswith('/variations'):
            return FakeResponse([
                {
                    'id': 501,
                    'sku': 'SKU-BLUE',
                    'status': 'publish',
                    'manage_stock': True,
                    'stock_quantity': 7,
                    'attributes': [{'name': 'Flavor', 'option': 'Blueberry Raspberry'}],
                },
                {
                    'id': 502,
                    'sku': '',
                    'status': 'publish',
                    'manage_stock': True,
                    'stock_quantity': 6,
                    'attributes': [{'name': 'Flavor', 'option': 'Cola Ice'}],
                },
                {
                    'id': 503,
                    'sku': '',
                    'status': 'publish',
                    'manage_stock': True,
                    'stock_quantity': 5,
                    'attributes': [{'name': 'Flavor', 'option': 'Mystery'}],
                },
            ])
        return FakeResponse([
            {
                'id': 100,
                'name': 'Fumot Leopard 40000 Puffs Disposable',
                'sku': '',
                'type': 'variable',
                'status': 'publish',
                'permalink': 'https://cz.test/product/fumot-leopard',
            },
            {
                'id': 999,
                'name': 'Unrelated Product',
                'sku': '',
                'type': 'simple',
                'status': 'publish',
            },
        ])


class WarehouseMappingTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(':memory:')
        self.db.row_factory = sqlite3.Row
        self.db.execute('PRAGMA foreign_keys=ON')
        self.db.executescript(
            '''
            CREATE TABLE warehouses (
              id INTEGER PRIMARY KEY, name TEXT, code TEXT, country TEXT, is_active INTEGER DEFAULT 1
            );
            CREATE TABLE sites (
              id INTEGER PRIMARY KEY, url TEXT, country TEXT, manager TEXT,
              consumer_key TEXT, consumer_secret TEXT
            );
            CREATE TABLE inv_market_warehouses (
              market_code TEXT, warehouse_id INTEGER, priority INTEGER,
              is_active INTEGER DEFAULT 1
            );
            CREATE TABLE inv_skus (
              id INTEGER PRIMARY KEY, sku_code TEXT, name TEXT, barcode TEXT, flavor TEXT,
              notes TEXT, is_active INTEGER DEFAULT 1
            );
            CREATE TABLE oms_sku_warehouses (
              sku_id INTEGER, warehouse_id INTEGER, is_enabled INTEGER DEFAULT 1
            );
            CREATE TABLE oms_warehouse_integrations (
              warehouse_id INTEGER PRIMARY KEY, inventory_authority TEXT
            );
            CREATE TABLE inv_stock (
              warehouse_id INTEGER, sku_id INTEGER, on_hand INTEGER, reserved INTEGER
            );
            CREATE TABLE inv_site_sku_map (
              id INTEGER PRIMARY KEY AUTOINCREMENT, site_id INTEGER, wc_product_id INTEGER,
              wc_variation_id INTEGER DEFAULT 0, wc_sku TEXT, raw_name TEXT, sku_id INTEGER,
              qty_per_item INTEGER DEFAULT 1, is_active INTEGER DEFAULT 1,
              UNIQUE(site_id,wc_product_id,wc_variation_id)
            );
            '''
        )
        up_019(self.db)
        self.db.executemany(
            'INSERT INTO warehouses (id,name,code,country,is_active) VALUES (?,?,?,?,1)',
            [(1, 'Managed warehouse', 'WH1', 'PL'), (2, 'Empty partner', 'WH2', 'PL')],
        )
        self.db.executemany(
            'INSERT INTO sites VALUES (?,?,?,?,?,?)',
            [
                (10, 'https://cz.test', 'CZ', 'Owner', 'ck', 'cs'),
                (11, 'https://au.test', 'AU', 'Other', 'ck', 'cs'),
            ],
        )
        self.db.execute('INSERT INTO inv_market_warehouses VALUES (\'CZ\',1,1,1)')
        note = json.dumps({'managed_family': 'fumot-leopard-40000-puffs'})
        self.db.executemany(
            'INSERT INTO inv_skus VALUES (?,?,?,?,?,?,1)',
            [
                (1, 'SKU-BLUE', 'Fumot Leopard 40000 Puffs - Blueberry Raspberry', None,
                 'Blueberry Raspberry', note),
                (2, 'SKU-COLA', 'Fumot Leopard 40000 Puffs - Cola Ice', None, 'Cola Ice', note),
            ],
        )
        self.db.executemany(
            'INSERT INTO oms_sku_warehouses VALUES (?,1,1)', [(1,), (2,)]
        )
        self.db.executemany(
            'INSERT INTO inv_stock VALUES (1,?,?,0)', [(1, 10), (2, 10)]
        )
        self.db.commit()

    def tearDown(self):
        self.db.close()

    def test_warehouse_and_site_scope_are_warehouse_first(self):
        rows = warehouse_rows(self.db)
        self.assertEqual([1], [row['id'] for row in rows])
        self.assertEqual([10], [row['id'] for row in serving_sites(self.db, 1)])

    def test_scan_requires_review_for_fuzzy_names_and_auto_maps_only_exact(self):
        result = scan_site_catalog(self.db, 10, 1, session=FakeWooSession())
        self.assertEqual(3, result['total_products'])

        overview = mapping_overview(self.db, 1)[0]
        self.assertEqual('action_needed', overview['readiness'])
        self.assertEqual(1, overview['exact_candidate_count'])
        self.assertEqual(1, overview['review_candidate_count'])
        self.assertEqual(1, overview['unresolved_count'])

        applied = apply_safe_mappings(self.db, 10, 1, operator_id=9, operator_name='tester')
        self.assertEqual(1, applied['created'])
        maps = self.db.execute('SELECT * FROM inv_site_sku_map').fetchall()
        self.assertEqual(1, len(maps))
        self.assertEqual(1, maps[0]['sku_id'])
        self.assertEqual(501, maps[0]['wc_variation_id'])
        self.assertEqual(1, self.db.execute('SELECT COUNT(*) FROM inv_mapping_audit').fetchone()[0])

        detail = mapping_detail(self.db, 1, 10)
        statuses = {row['id']: row['status'] for row in detail['skus']}
        self.assertEqual('mapped', statuses[1])
        self.assertEqual('candidate', statuses[2])


if __name__ == '__main__':
    unittest.main()
