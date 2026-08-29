import sqlite3
import unittest
from unittest.mock import patch

import inv_migrations
import inv_push


class InventoryAutoPushTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self.db.executescript(
            """
            PRAGMA foreign_keys=ON;
            CREATE TABLE settings (key TEXT PRIMARY KEY,value TEXT);
            CREATE TABLE sites (
              id INTEGER PRIMARY KEY,url TEXT,country TEXT,manager TEXT,
              consumer_key TEXT,consumer_secret TEXT,product_master_id INTEGER
            );
            CREATE TABLE warehouses (
              id INTEGER PRIMARY KEY,name TEXT,code TEXT,country TEXT,is_active INTEGER
            );
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
            CREATE TABLE inv_stock (
              warehouse_id INTEGER,sku_id INTEGER,on_hand INTEGER,reserved INTEGER
            );
            CREATE TABLE oms_external_stock (
              warehouse_id INTEGER,sku_id INTEGER,available_quantity INTEGER,
              source_updated_at TEXT,synced_at TEXT
            );
            """
        )
        self.db.executemany(
            "INSERT INTO sites VALUES (?,?,?,?,?,?,NULL)",
            [
                (1, "https://one.test", "PL", "One", "ck", "cs"),
                (2, "https://two.test", "PL", "Two", "ck", "cs"),
            ],
        )
        self.db.execute("INSERT INTO warehouses VALUES (1,'Finite','FINITE','PL',1)")
        self.db.execute(
            "INSERT INTO oms_warehouse_integrations VALUES (1,'local','{}')"
        )
        self.db.execute("INSERT INTO inv_skus VALUES (1,'SKU1','Product')")
        self.db.executemany(
            "INSERT INTO inv_site_sku_map VALUES (?,?,?,?,?,1)",
            [(1, 101, 0, 1, 1), (2, 201, 0, 1, 1)],
        )
        self.db.execute("INSERT INTO oms_sku_warehouses VALUES (1,1,1)")
        self.db.execute("INSERT INTO inv_stock VALUES (1,1,9,0)")
        inv_migrations.up_004(self.db)
        inv_migrations.up_021(self.db)

    def tearDown(self):
        self.db.close()

    @patch("inv_push.inv_allocator.candidate_warehouses")
    def test_shared_stock_is_weighted_without_overallocation(self, candidates):
        candidates.return_value = [{"warehouse_id": 1}]
        inv_push.update_site_sync_config(
            self.db, 1, {"mode": "observe", "allocation_weight": 1}, (1, "admin")
        )
        inv_push.update_site_sync_config(
            self.db, 2, {"mode": "observe", "allocation_weight": 2}, (1, "admin")
        )
        one = inv_push.compute_site_stock(self.db, 1, use_sync_strategy=True)[0]
        two = inv_push.compute_site_stock(self.db, 2, use_sync_strategy=True)[0]
        self.assertEqual(3, one["allocated_sku"])
        self.assertEqual(6, two["allocated_sku"])
        self.assertEqual(9, one["allocated_sku"] + two["allocated_sku"])
        self.assertEqual(2, one["allocation_participants"])

    @patch("inv_push.inv_allocator.candidate_warehouses")
    def test_live_readiness_requires_shared_sites_to_move_together(self, candidates):
        candidates.return_value = [{"warehouse_id": 1}]
        inv_push.update_site_sync_config(
            self.db, 1, {"mode": "live"}, (1, "admin")
        )
        readiness = inv_push.site_sync_readiness(self.db, 1, for_live=True)
        self.assertFalse(readiness["ready"])
        self.assertEqual([2], readiness["shared_not_live_site_ids"])
        inv_push.update_site_sync_config(
            self.db, 2, {"mode": "live"}, (1, "admin")
        )
        self.assertTrue(inv_push.site_sync_readiness(self.db, 1, for_live=True)["ready"])

    @patch("inv_push.inv_allocator.candidate_warehouses")
    def test_mirror_strategy_is_blocked_for_shared_live_stock(self, candidates):
        candidates.return_value = [{"warehouse_id": 1}]
        inv_push.update_site_sync_config(
            self.db, 1, {"mode": "live", "allocation_strategy": "mirror"}, (1, "admin")
        )
        readiness = inv_push.site_sync_readiness(self.db, 1, for_live=True)
        self.assertFalse(readiness["ready"])
        self.assertIn("镜像库存不能用于", "；".join(readiness["errors"]))

    def test_global_and_site_switches_are_fail_closed(self):
        inv_push.update_site_sync_config(
            self.db, 1, {"mode": "observe"}, (1, "admin")
        )
        self.assertFalse(inv_push.global_sync_enabled(self.db))
        self.assertEqual([], inv_push.scheduler_site_ids(self.db))
        self.db.execute(
            "UPDATE settings SET value='1' WHERE key='inv_auto_push_global_enabled'"
        )
        self.db.commit()
        self.assertEqual([1], inv_push.scheduler_site_ids(self.db))

    @patch("inv_push._put_stock")
    @patch("inv_push._get_stock_state")
    def test_write_is_read_back_and_verified(self, get_state, put_stock):
        get_state.side_effect = [
            ({"manage_stock": True, "stock_quantity": 4}, None),
            ({"manage_stock": True, "stock_quantity": 7}, None),
        ]
        put_stock.return_value = (True, None)
        status, previous, remote, error = inv_push._sync_one_stock(
            "https://one.test",
            "ck",
            "cs",
            {"wc_product_id": 101, "wc_variation_id": 0, "publishable": 7},
        )
        self.assertEqual(("ok", 4, 7, None), (status, previous, remote, error))
        put_stock.assert_called_once()
        self.assertEqual(2, get_state.call_count)

    @patch("inv_push._put_stock")
    @patch("inv_push._get_stock_state")
    def test_unchanged_stock_skips_put(self, get_state, put_stock):
        get_state.return_value = ({"manage_stock": True, "stock_quantity": 7}, None)
        result = inv_push._sync_one_stock(
            "https://one.test",
            "ck",
            "cs",
            {"wc_product_id": 101, "wc_variation_id": 0, "publishable": 7},
        )
        self.assertEqual("unchanged", result[0])
        put_stock.assert_not_called()

    def test_mass_zero_guard_blocks_suspicious_publish(self):
        for product_id in range(1, 5):
            self.db.execute(
                """INSERT INTO inv_push_logs
                   (site_id,wc_product_id,wc_variation_id,pushed_qty,status)
                   VALUES (1,?,0,10,'ok')""",
                (product_id,),
            )
        self.db.commit()
        items = [
            {"wc_product_id": product_id, "wc_variation_id": 0, "publishable": 0}
            for product_id in range(1, 5)
        ]
        self.assertIn("安全拦截", inv_push.detect_mass_drop(self.db, 1, items))

    @patch("inv_push.push_site")
    def test_repeated_failures_pause_the_site(self, push_site):
        push_site.return_value = {
            "site_id": 1,
            "total": 1,
            "ok": 0,
            "unchanged": 0,
            "error": 1,
            "fatal": "simulated",
        }
        inv_push.update_site_sync_config(
            self.db,
            1,
            {"mode": "observe", "failure_threshold": 2},
            (1, "admin"),
        )
        self.db.execute(
            "UPDATE settings SET value='1' WHERE key='inv_auto_push_global_enabled'"
        )
        self.db.commit()
        first = inv_push.execute_site_sync(self.db, 1)
        self.assertEqual("error", first["status"])
        self.db.execute(
            "UPDATE inv_site_sync_config SET next_run_at=NULL WHERE site_id=1"
        )
        self.db.commit()
        inv_push.execute_site_sync(self.db, 1)
        config = inv_push.get_site_sync_config(self.db, 1)
        self.assertEqual("paused", config["mode"])
        self.assertEqual(2, config["consecutive_failures"])


class InventoryAutoPushMigrationTests(unittest.TestCase):
    def test_roundtrip_preserves_preexisting_setting(self):
        db = sqlite3.connect(":memory:")
        db.row_factory = sqlite3.Row
        db.executescript(
            """
            PRAGMA foreign_keys=ON;
            CREATE TABLE sites (id INTEGER PRIMARY KEY);
            CREATE TABLE settings (key TEXT PRIMARY KEY,value TEXT);
            INSERT INTO settings VALUES ('inv_auto_push_max_drop_percent','75');
            """
        )
        inv_migrations.up_021(db)
        inv_migrations.up_021(db)
        self.assertEqual(
            "0",
            db.execute(
                "SELECT value FROM settings WHERE key='inv_auto_push_global_enabled'"
            ).fetchone()["value"],
        )
        inv_migrations.down_021(db)
        self.assertEqual(
            "75",
            db.execute(
                "SELECT value FROM settings WHERE key='inv_auto_push_max_drop_percent'"
            ).fetchone()["value"],
        )
        self.assertIsNone(
            db.execute(
                "SELECT 1 FROM settings WHERE key='inv_auto_push_global_enabled'"
            ).fetchone()
        )
        db.close()


if __name__ == "__main__":
    unittest.main()
