import json
import sqlite3
import unittest
from unittest.mock import patch

from inv_push import compute_site_stock


class InventoryPushScopeTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(':memory:')
        self.db.row_factory = sqlite3.Row
        self.db.executescript(
            '''
            CREATE TABLE sites (id INTEGER PRIMARY KEY,url TEXT,country TEXT);
            CREATE TABLE warehouses (id INTEGER PRIMARY KEY,name TEXT,code TEXT,country TEXT,is_active INTEGER);
            CREATE TABLE oms_warehouse_integrations (
              warehouse_id INTEGER PRIMARY KEY,inventory_authority TEXT,config_json TEXT
            );
            CREATE TABLE inv_skus (id INTEGER PRIMARY KEY,sku_code TEXT,name TEXT);
            CREATE TABLE inv_site_sku_map (
              site_id INTEGER,wc_product_id INTEGER,wc_variation_id INTEGER,
              sku_id INTEGER,qty_per_item INTEGER,is_active INTEGER
            );
            CREATE TABLE oms_sku_warehouses (
              warehouse_id INTEGER,sku_id INTEGER,is_enabled INTEGER
            );
            CREATE TABLE inv_stock (warehouse_id INTEGER,sku_id INTEGER,on_hand INTEGER,reserved INTEGER);
            CREATE TABLE oms_external_stock (warehouse_id INTEGER,sku_id INTEGER,available_quantity INTEGER);
            '''
        )
        self.db.execute("INSERT INTO sites VALUES (1,'https://cz.test','CZ')")
        self.db.executemany(
            'INSERT INTO warehouses VALUES (?,?,?,?,1)',
            [(1, 'Finite warehouse', 'FINITE', 'PL'), (2, 'Manual partner', 'MANUAL', 'PL')],
        )
        self.db.executemany(
            'INSERT INTO oms_warehouse_integrations VALUES (?,?,?)',
            [
                (1, 'local', json.dumps({'stock_policy': 'finite_local_ledger'})),
                (2, 'manual_partner', json.dumps({'requires_quantity_inventory': False})),
            ],
        )
        self.db.execute("INSERT INTO inv_skus VALUES (1,'SKU1','Product')")
        self.db.execute('INSERT INTO inv_site_sku_map VALUES (1,100,501,1,1,1)')
        self.db.executemany('INSERT INTO oms_sku_warehouses VALUES (?,1,1)', [(1,), (2,)])
        self.db.execute('INSERT INTO inv_stock VALUES (1,1,7,2)')
        self.db.commit()

    def tearDown(self):
        self.db.close()

    @patch('inv_push.inv_allocator.candidate_warehouses')
    def test_manual_partner_is_excluded_from_woo_stock(self, candidates):
        candidates.return_value = [{'warehouse_id': 1}, {'warehouse_id': 2}]
        row = compute_site_stock(self.db, 1)[0]
        self.assertEqual(5, row['available_sku'])
        self.assertEqual(5, row['publishable'])
        self.assertEqual([1], row['serving_warehouses'])
        self.assertEqual(501, row['wc_variation_id'])


if __name__ == '__main__':
    unittest.main()
