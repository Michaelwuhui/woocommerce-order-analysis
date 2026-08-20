import json
import re
import sqlite3
import unittest
from decimal import Decimal

from fulfillment_service import (
    add_tracking_event,
    build_poland_wms_payload,
    build_wms_payload,
    completion_guard,
    create_shipment,
    managed_product_family,
    order_contains_managed_product,
    plan_order,
    recompute_order_status,
    transition_fulfillment,
)
from inv_migrations import (
    down_009,
    down_010,
    down_011,
    down_012,
    down_017,
    down_018,
    down_020,
    up_001,
    up_005,
    up_006,
    up_007,
    up_008,
    up_009,
    up_010,
    up_011,
    up_012,
    up_017,
    up_018,
    up_020,
)
from inv_common import record_movement, replenishment_metrics
from inv_mapping_service import replan_shortage_orders_for_mappings
from fulfillment_woocommerce import _all_order_shipments, _ast_items, _customer_note_body
from shipment_customer_messages import basic_shipment_note


class FulfillmentDomainTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.executescript(
            """
            CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT, name TEXT);
            CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT);
            CREATE TABLE sites (id INTEGER PRIMARY KEY, url TEXT UNIQUE, country TEXT);
            CREATE TABLE warehouses (
              id INTEGER PRIMARY KEY, code TEXT, name TEXT, country TEXT,
              is_active INTEGER DEFAULT 1
            );
            CREATE TABLE orders (
              id TEXT PRIMARY KEY, number TEXT, status TEXT, source TEXT,
              line_items TEXT, billing TEXT, shipping TEXT, total REAL,
              shipping_total REAL, currency TEXT, payment_method TEXT, date_created TEXT
            );
            """
        )
        self.db.executemany(
            "INSERT INTO warehouses (id,code,name,country,is_active) VALUES (?,?,?,?,1)",
            [(1, "PL", "波兰仓", "PL"), (2, "HU", "匈牙利仓", "HU")],
        )
        self.db.executemany(
            "INSERT INTO sites (id,url,country) VALUES (?,?,?)",
            [(1, "https://pl.test", "PL"), (2, "https://hu.test", "HU"), (3, "https://cz.test", "CZ")],
        )
        up_001(self.db)
        up_006(self.db)
        up_007(self.db)
        up_008(self.db)
        self.db.execute(
            "INSERT INTO inv_skus (id,sku_code,name,barcode,is_active) VALUES (1,'SKU1','Test product','BAR1',1)"
        )
        for site_id in (1, 2, 3):
            self.db.execute(
                """INSERT INTO inv_site_sku_map
                   (site_id,wc_product_id,wc_variation_id,wc_sku,raw_name,sku_id,qty_per_item,is_active)
                   VALUES (?,101,0,'SKU1','Test product',1,1,1)""",
                (site_id,),
            )
        self.db.executemany(
            """INSERT INTO oms_sku_warehouses
               (sku_id,warehouse_id,is_primary,is_enabled,wms_product_name_zh,wms_product_name_en,product_type)
               VALUES (1,?,?,1,'测试产品','Test product','P')""",
            [(1, 1), (2, 0)],
        )
        self.set_stock(pl=20, hu=20)

    def tearDown(self):
        self.db.close()

    def set_stock(self, *, pl, hu):
        self.db.execute(
            """INSERT INTO inv_stock (warehouse_id,sku_id,on_hand,reserved,updated_at)
               VALUES (1,1,?,0,CURRENT_TIMESTAMP)
               ON CONFLICT(warehouse_id,sku_id) DO UPDATE SET on_hand=excluded.on_hand,reserved=0""",
            (pl,),
        )
        self.db.execute(
            """INSERT INTO oms_external_stock
               (warehouse_id,sku_barcode,sku_id,quantity,lock_quantity,available_quantity)
               VALUES (2,'BAR1',1,?,0,?)
               ON CONFLICT(warehouse_id,sku_barcode) DO UPDATE SET
                 quantity=excluded.quantity,available_quantity=excluded.available_quantity""",
            (hu, hu),
        )
        self.db.commit()

    def add_order(
        self,
        oid,
        market,
        qty=1,
        *,
        shipping_total=0,
        currency="EUR",
        line_total=None,
        line_tax=0,
        order_total=None,
        payment_method="cod",
        product_id=101,
        sku="SKU1",
        name="Test product",
    ):
        site = {"PL": "https://pl.test", "HU": "https://hu.test", "CZ": "https://cz.test"}[market]
        if line_total is None:
            line_total = qty * 10
        if order_total is None:
            order_total = Decimal(str(line_total)) + Decimal(str(line_tax)) + Decimal(str(shipping_total))
        item = {
            "id": 501,
            "product_id": product_id,
            "variation_id": 0,
            "sku": sku,
            "name": name,
            "quantity": qty,
            "total": str(line_total),
            "total_tax": str(line_tax),
        }
        address = {"first_name": "Test", "last_name": "Buyer", "email": "buyer@example.test", "phone": "123456", "address_1": "Main 1", "address_2": "", "city": "Budapest", "state": "", "postcode": "1000", "country": market}
        self.db.execute(
            """INSERT INTO orders
               (id,number,status,source,line_items,billing,shipping,total,
                shipping_total,currency,payment_method,date_created)
               VALUES (?,?, 'processing', ?,?,?,?,?,?,?,?,'2026-07-18T10:00:00')""",
            (
                oid,
                oid,
                site,
                json.dumps([item]),
                json.dumps(address),
                json.dumps(address),
                float(order_total),
                float(shipping_total),
                currency,
                payment_method,
            ),
        )
        self.db.commit()

    def allocations(self, order_id):
        return [dict(r) for r in self.db.execute(
            """SELECT f.warehouse_id, SUM(fi.allocated_qty) qty
               FROM oms_fulfillments f JOIN oms_fulfillment_items fi ON fi.fulfillment_id=f.id
               WHERE f.order_id=? AND f.status!='superseded' GROUP BY f.warehouse_id ORDER BY f.warehouse_id""",
            (order_id,),
        )]

    def test_site_preference_and_idempotent_plan(self):
        self.add_order("pl-1", "PL", 2)
        first = plan_order(self.db, "pl-1")
        second = plan_order(self.db, "pl-1")
        self.assertEqual([{"warehouse_id": 1, "qty": 2}], self.allocations("pl-1"))
        self.assertEqual("planned", first["action"])
        self.assertEqual("noop", second["action"])
        self.assertEqual(1, second["revision"])

        self.add_order("hu-1", "HU", 2)
        plan_order(self.db, "hu-1")
        self.assertEqual([{"warehouse_id": 2, "qty": 2}], self.allocations("hu-1"))

    def test_started_unchanged_order_plan_is_locked_noop_not_manual_failure(self):
        self.add_order("locked-noop", "PL", 1)
        planned = plan_order(self.db, "locked-noop")
        create_shipment(self.db, planned["fulfillment_ids"][0], "LOCKED-001")

        repeated = plan_order(self.db, "locked-noop")

        state = self.db.execute(
            "SELECT manual_review FROM oms_order_fulfillment_state WHERE order_id='locked-noop'"
        ).fetchone()
        self.assertEqual("locked_noop", repeated["action"])
        self.assertEqual(0, state["manual_review"])

    def test_split_order_and_czech_cost_routing(self):
        self.set_stock(pl=1, hu=10)
        self.add_order("split-1", "PL", 3)
        plan_order(self.db, "split-1")
        self.assertEqual([{"warehouse_id": 1, "qty": 1}, {"warehouse_id": 2, "qty": 2}], self.allocations("split-1"))

        self.db.executemany(
            "INSERT INTO oms_shipping_costs (market_code,warehouse_id,service_code,amount,currency,is_active) VALUES ('CZ',?,'default',?,'EUR',1)",
            [(1, 4.5), (2, 7.0)],
        )
        self.add_order("cz-1", "CZ", 1)
        plan_order(self.db, "cz-1")
        self.assertEqual([{"warehouse_id": 1, "qty": 1}], self.allocations("cz-1"))

    def test_shortage_is_prominent_manual_state(self):
        self.set_stock(pl=0, hu=0)
        self.add_order("short-1", "PL", 2)
        result = plan_order(self.db, "short-1")
        state = self.db.execute("SELECT * FROM oms_order_fulfillment_state WHERE order_id='short-1'").fetchone()
        self.assertEqual("stock_shortage", result["aggregate_status"])
        self.assertEqual(1, state["has_shortage"])
        self.assertEqual(1, state["manual_review"])
        self.assertFalse(completion_guard(self.db, "short-1")[0])

    def test_reserved_family_isolated_before_sku_mapping(self):
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES ('oms_managed_product_isolation_enabled','1')"
        )
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)",
            (
                "oms_managed_product_families",
                json.dumps([
                    "fumot-eco-4in1-80k", "fumot-leopard-40k",
                    "fumot-randm-tornado-9000", "fumot-randm-tornado-15000",
                ]),
            ),
        )
        self.add_order(
            "managed-no-map", "HU", product_id=999, sku="",
            name="Fumot RandM Tornado 9000 - Kiwi Passion Fruit Guava",
        )

        result = plan_order(self.db, "managed-no-map")
        item = self.db.execute(
            "SELECT sku_id FROM oms_order_items WHERE order_id='managed-no-map'"
        ).fetchone()

        self.assertEqual("stock_shortage", result["aggregate_status"])
        self.assertTrue(result["shortages"])
        self.assertIsNone(item["sku_id"])

    def test_reserved_family_never_falls_back_to_legacy_mapped_sku(self):
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES ('oms_managed_product_isolation_enabled','1')"
        )
        self.db.execute("DELETE FROM oms_sku_warehouses WHERE warehouse_id=2")
        self.db.execute(
            "UPDATE oms_warehouse_integrations SET inventory_authority='manual_partner' WHERE warehouse_id=1"
        )
        self.add_order(
            "managed-legacy-map", "HU", product_id=101, sku="SKU1",
            name="Fumot RandM Tornado 9000 Strawberry",
        )

        result = plan_order(self.db, "managed-legacy-map")

        self.assertEqual("stock_shortage", result["aggregate_status"])
        self.assertTrue(result["shortages"])
        self.assertEqual([], self.allocations("managed-legacy-map"))

    def test_new_mapping_automatically_replans_matching_shortage_order(self):
        self.add_order("mapping-replan", "PL", product_id=999, sku="", name="New product")
        first = plan_order(self.db, "mapping-replan")
        self.assertEqual("stock_shortage", first["aggregate_status"])
        self.db.execute(
            """INSERT INTO inv_site_sku_map
               (site_id,wc_product_id,wc_variation_id,wc_sku,raw_name,sku_id,qty_per_item,is_active)
               VALUES (1,999,0,NULL,'New product',1,1,1)"""
        )

        result = replan_shortage_orders_for_mappings(
            self.db, 1, [{"wc_product_id": 999, "wc_variation_id": 0}],
            operator_id=9, operator_name="tester",
        )

        state = self.db.execute(
            "SELECT * FROM oms_order_fulfillment_state WHERE order_id='mapping-replan'"
        ).fetchone()
        self.assertEqual({"matched": 1, "updated": 1, "failed": []}, result)
        self.assertEqual(0, state["has_shortage"])
        self.assertEqual([{"warehouse_id": 1, "qty": 1}], self.allocations("mapping-replan"))

    def test_managed_blueberry_on_ice_alias_allocates_blueberry_ice_order(self):
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES ('oms_managed_product_isolation_enabled','1')"
        )
        up_017(self.db)
        warehouse_id = self.db.execute(
            "SELECT id FROM warehouses WHERE code='PL-JYJG-TRANSIT'"
        ).fetchone()[0]
        self.add_order(
            "blueberry-ice", "PL", product_id=904, sku="",
            name="Fumot Leopard 40000 Puffs - Blueberry Ice",
        )

        result = plan_order(self.db, "blueberry-ice")
        mapped = self.db.execute(
            """SELECT s.sku_code FROM oms_order_items oi
               JOIN inv_skus s ON s.id=oi.sku_id WHERE oi.order_id='blueberry-ice'"""
        ).fetchone()["sku_code"]

        self.assertFalse(result["shortages"])
        self.assertEqual("40K-BOI", mapped)
        self.assertEqual(
            [{"warehouse_id": warehouse_id, "qty": 1}], self.allocations("blueberry-ice")
        )

    def test_jyjg_transit_warehouse_uses_finite_reserved_stock(self):
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES ('oms_managed_product_isolation_enabled','1')"
        )
        up_017(self.db)
        warehouse_id = self.db.execute(
            "SELECT id FROM warehouses WHERE code='PL-JYJG-TRANSIT'"
        ).fetchone()[0]
        self.assertEqual(
            (59, 590, 0),
            tuple(self.db.execute(
                """SELECT COUNT(*),SUM(on_hand),SUM(reserved)
                   FROM inv_stock WHERE warehouse_id=?""",
                (warehouse_id,),
            ).fetchone()),
        )

        self.add_order(
            "transit-first", "CZ", 7, product_id=901, sku="",
            name="Fumot Leopard 40000 Puffs - Cola Ice",
        )
        first = plan_order(self.db, "transit-first")
        self.assertFalse(first["shortages"])
        self.assertEqual(
            [{"warehouse_id": warehouse_id, "qty": 7}],
            self.allocations("transit-first"),
        )
        self.assertEqual("noop", plan_order(self.db, "transit-first")["action"])

        self.add_order(
            "transit-second", "HU", 4, product_id=902, sku="",
            name="Fumot Leopard 40000 Puffs - Cola Ice",
        )
        second = plan_order(self.db, "transit-second")
        self.assertEqual("stock_shortage", second["aggregate_status"])
        self.assertEqual(1, second["shortages"][0]["qty"])
        self.assertEqual(
            [{"warehouse_id": warehouse_id, "qty": 3}],
            self.allocations("transit-second"),
        )
        stock = self.db.execute(
            """SELECT st.on_hand,st.reserved
               FROM inv_stock st JOIN inv_skus s ON s.id=st.sku_id
               WHERE st.warehouse_id=? AND s.sku_code='40K-CI'""",
            (warehouse_id,),
        ).fetchone()
        self.assertEqual((10, 10), tuple(stock))

        first_fulfillment = self.db.execute(
            "SELECT id FROM oms_fulfillments WHERE order_id='transit-first' AND status!='superseded'"
        ).fetchone()[0]
        transition_fulfillment(self.db, first_fulfillment, "cancelled")
        self.assertEqual("planned", plan_order(self.db, "transit-second")["action"])
        self.assertEqual(
            [{"warehouse_id": warehouse_id, "qty": 4}],
            self.allocations("transit-second"),
        )
        stock = self.db.execute(
            """SELECT st.on_hand,st.reserved
               FROM inv_stock st JOIN inv_skus s ON s.id=st.sku_id
               WHERE st.warehouse_id=? AND s.sku_code='40K-CI'""",
            (warehouse_id,),
        ).fetchone()
        self.assertEqual((10, 4), tuple(stock))

    def test_jyjg_transit_shipment_deducts_once_and_rollback_disables_new_routing(self):
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES ('oms_managed_product_isolation_enabled','1')"
        )
        up_017(self.db)
        warehouse_id = self.db.execute(
            "SELECT id FROM warehouses WHERE code='PL-JYJG-TRANSIT'"
        ).fetchone()[0]
        self.add_order(
            "transit-ship", "PL", 2, product_id=903, sku="",
            name="Fumot RandM Tornado 9000 Puffs - Grape",
        )
        result = plan_order(self.db, "transit-ship")
        shipment = create_shipment(
            self.db,
            result["fulfillment_ids"][0],
            "Z-TRANSIT-001",
            carrier_slug="packeta",
        )
        duplicate = create_shipment(
            self.db,
            result["fulfillment_ids"][0],
            "Z-TRANSIT-001",
            carrier_slug="packeta",
        )
        self.assertEqual(shipment["id"], duplicate["id"])
        stock = self.db.execute(
            """SELECT st.on_hand,st.reserved
               FROM inv_stock st JOIN inv_skus s ON s.id=st.sku_id
               WHERE st.warehouse_id=? AND s.sku_code='9K-G'""",
            (warehouse_id,),
        ).fetchone()
        self.assertEqual((8, 0), tuple(stock))
        self.assertEqual(
            1,
            self.db.execute(
                """SELECT COUNT(*) FROM inv_movements
                   WHERE warehouse_id=? AND movement_type='sale_out'""",
                (warehouse_id,),
            ).fetchone()[0],
        )

        down_017(self.db)
        self.assertEqual(
            0,
            self.db.execute(
                "SELECT is_active FROM warehouses WHERE id=?", (warehouse_id,)
            ).fetchone()[0],
        )
        self.assertEqual(
            8,
            self.db.execute(
                """SELECT st.on_hand FROM inv_stock st JOIN inv_skus s ON s.id=st.sku_id
                   WHERE st.warehouse_id=? AND s.sku_code='9K-G'""",
                (warehouse_id,),
            ).fetchone()[0],
        )

    def test_joint_dispatch_uses_one_tracking_and_separate_stock_ledgers(self):
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES ('oms_managed_product_isolation_enabled','1')"
        )
        up_017(self.db)
        up_020(self.db)
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)",
            (
                "oms_joint_dispatch_groups",
                json.dumps({
                    "jyjg": {
                        "label": "Joint Poland dispatch",
                        "warehouse_codes": ["PL", "PL-JYJG-TRANSIT"],
                    }
                }),
            ),
        )
        self.db.execute(
            "DELETE FROM oms_sku_warehouses WHERE sku_id=1 AND warehouse_id=2"
        )
        self.add_order("joint-dispatch", "CZ", 1)
        items = json.loads(self.db.execute(
            "SELECT line_items FROM orders WHERE id='joint-dispatch'"
        ).fetchone()[0])
        items.append({
            "id": 502,
            "product_id": 902,
            "variation_id": 0,
            "sku": "",
            "name": "Fumot Leopard 40000 Puffs - Cola Ice",
            "quantity": 1,
            "total": "10.00",
            "total_tax": "0.00",
        })
        self.db.execute(
            "UPDATE orders SET line_items=?,total=20 WHERE id='joint-dispatch'",
            (json.dumps(items),),
        )
        planned = plan_order(self.db, "joint-dispatch")
        fulfillments = self.db.execute(
            """SELECT id,warehouse_id FROM oms_fulfillments
               WHERE order_id='joint-dispatch' AND status!='superseded' ORDER BY warehouse_id"""
        ).fetchall()
        self.assertEqual(2, len(fulfillments))
        self.assertEqual(
            ["PL", "PL-JYJG-TRANSIT"],
            [self.db.execute("SELECT code FROM warehouses WHERE id=?", (row["warehouse_id"],)).fetchone()[0]
             for row in fulfillments],
        )

        shipment = create_shipment(
            self.db, fulfillments[0]["id"], "JOINT-TRACK-001", carrier_slug="packeta"
        )
        recompute_order_status(self.db, "joint-dispatch")

        self.assertTrue(shipment["joint_dispatch"])
        self.assertEqual(2, len(shipment["fulfillment_ids"]))
        self.assertEqual(1, self.db.execute(
            "SELECT COUNT(*) FROM oms_shipments WHERE tracking_number='JOINT-TRACK-001'"
        ).fetchone()[0])
        self.assertEqual(2, self.db.execute(
            "SELECT COUNT(*) FROM oms_shipment_fulfillments WHERE shipment_id=?",
            (shipment["id"],),
        ).fetchone()[0])
        self.assertEqual(2, self.db.execute(
            "SELECT COUNT(*) FROM oms_shipment_items WHERE shipment_id=?",
            (shipment["id"],),
        ).fetchone()[0])
        parcel_rows = _all_order_shipments(self.db, "joint-dispatch", planned["revision"])
        self.assertEqual(1, len(parcel_rows))
        self.assertEqual(2, len(parcel_rows[0]["products"]))
        order = self.db.execute("SELECT * FROM orders WHERE id='joint-dispatch'").fetchone()
        notice = _customer_note_body(self.db, order, shipment)
        self.assertIn("Dobírka za tuto zásilku činí 20.00 EUR", notice)
        self.assertNotIn("samostatné zásilce", notice)
        self.assertIsNone(re.search(r"[\u4e00-\u9fff]", notice))
        self.assertEqual(
            ["shipped", "shipped"],
            [row[0] for row in self.db.execute(
                """SELECT status FROM oms_fulfillments
                   WHERE order_id='joint-dispatch' AND status!='superseded' ORDER BY warehouse_id"""
            ).fetchall()],
        )
        transit_stock = self.db.execute(
            """SELECT st.on_hand,st.reserved FROM inv_stock st
               JOIN inv_skus s ON s.id=st.sku_id
               JOIN warehouses w ON w.id=st.warehouse_id
               WHERE w.code='PL-JYJG-TRANSIT' AND s.sku_code='40K-CI'"""
        ).fetchone()
        self.assertEqual((9, 0), tuple(transit_stock))
        aggregate = self.db.execute(
            "SELECT aggregate_status FROM oms_order_fulfillment_state WHERE order_id='joint-dispatch'"
        ).fetchone()[0]
        self.assertEqual("shipped", aggregate)

        add_tracking_event(
            self.db, shipment["id"], "official", "delivered",
            raw_status="delivered", external_event_id="joint-delivered",
        )
        recompute_order_status(self.db, "joint-dispatch")
        self.assertEqual(
            ["delivered", "delivered"],
            [row[0] for row in self.db.execute(
                """SELECT status FROM oms_fulfillments
                   WHERE order_id='joint-dispatch' AND status!='superseded' ORDER BY warehouse_id"""
            ).fetchall()],
        )
        self.assertTrue(completion_guard(self.db, "joint-dispatch")[0])

        down_020(self.db)

    def test_jyjg_replenishment_warning_lifecycle(self):
        up_005(self.db)
        up_017(self.db)
        warehouse_id = self.db.execute(
            "SELECT id FROM warehouses WHERE code='PL-JYJG-TRANSIT'"
        ).fetchone()[0]
        sku_id = self.db.execute(
            "SELECT id FROM inv_skus WHERE sku_code='40K-CI'"
        ).fetchone()[0]

        record_movement(
            self.db, warehouse_id=warehouse_id, sku_id=sku_id,
            movement_type="reserve", reserved_delta=7,
            ref_type="test", ref_id="low-stock",
        )
        metrics = replenishment_metrics(self.db, warehouse_id, sku_id)
        notice = self.db.execute(
            "SELECT * FROM inv_notifications WHERE dedup_key=? AND status='unread'",
            (f"restock:{warehouse_id}:{sku_id}",),
        ).fetchone()
        self.assertEqual(3, metrics["available"])
        self.assertEqual(7, metrics["suggested_replenishment"])
        self.assertIn("建议补 7 支", notice["body"])

        record_movement(
            self.db, warehouse_id=warehouse_id, sku_id=sku_id,
            movement_type="release", reserved_delta=-7,
            ref_type="test", ref_id="restocked",
        )
        self.assertIsNone(self.db.execute(
            "SELECT 1 FROM inv_notifications WHERE dedup_key=? AND status='unread'",
            (f"restock:{warehouse_id}:{sku_id}",),
        ).fetchone())

    def test_manual_shipper_scope_migration_roundtrip(self):
        up_017(self.db)
        warehouse_id = self.db.execute(
            "SELECT id FROM warehouses WHERE code='PL-JYJG-TRANSIT'"
        ).fetchone()[0]
        self.db.execute(
            "INSERT INTO users (id,username,name) VALUES (10,'jinyi','金毅')"
        )
        self.db.execute("CREATE TABLE partners (id INTEGER PRIMARY KEY, name TEXT)")
        self.db.execute("INSERT INTO partners (id,name) VALUES (1,'金谷金毅（波兰）')")
        self.db.commit()

        up_018(self.db)
        permission = self.db.execute(
            """SELECT can_view,can_pick,can_pack,can_ship,can_cancel
               FROM oms_warehouse_user_permissions WHERE user_id=10 AND warehouse_id=?""",
            (warehouse_id,),
        ).fetchone()
        self.assertEqual((1, 0, 0, 1, 0), tuple(permission))
        self.assertEqual(
            1,
            self.db.execute(
                "SELECT partner_id FROM inv_warehouse_ext WHERE warehouse_id=?",
                (warehouse_id,),
            ).fetchone()[0],
        )
        self.assertEqual(
            ("3", "10"),
            tuple(row[0] for row in self.db.execute(
                "SELECT value FROM settings WHERE key IN ('inv_jyjg_reorder_point','inv_jyjg_target_stock') ORDER BY key"
            ).fetchall()),
        )

        down_018(self.db)
        self.assertIsNone(self.db.execute(
            "SELECT 1 FROM oms_warehouse_user_permissions WHERE user_id=10 AND warehouse_id=?",
            (warehouse_id,),
        ).fetchone())
        self.assertIsNone(self.db.execute(
            "SELECT partner_id FROM inv_warehouse_ext WHERE warehouse_id=?",
            (warehouse_id,),
        ).fetchone()[0])

    def test_managed_family_matching_is_exact_enough(self):
        self.db.execute(
            "INSERT OR REPLACE INTO settings (key,value) VALUES ('oms_managed_product_isolation_enabled','1')"
        )
        positives = [
            {"name": "Fumot Eco 4in1 80K Blueberry"},
            {"name": "Fumot Leopard 40K Mango"},
            {"name": "Fumot RandM Tornado 9000 Strawberry"},
            {"name": "Fumot RandM Tornado 15000 Grape"},
        ]
        for item in positives:
            with self.subTest(item=item):
                self.assertTrue(managed_product_family(self.db, item))
        self.assertFalse(managed_product_family(
            self.db, {"name": "Fumot RandM Tornado 90000 Strawberry"}
        ))
        self.assertFalse(managed_product_family(
            self.db, {"name": "Generic Tornado 9000 replacement coil"}
        ))
        self.add_order(
            "managed-detect", "HU", product_id=999, sku="",
            name=positives[0]["name"],
        )
        self.assertTrue(order_contains_managed_product(self.db, "managed-detect"))

    def test_packeta_hu_migration_is_fail_closed_and_reversible(self):
        up_009(self.db)
        up_010(self.db)
        self.db.execute(
            """CREATE TABLE IF NOT EXISTS shipping_carriers (
                 slug TEXT PRIMARY KEY, name TEXT, tracking_url TEXT, is_active INTEGER
               )"""
        )
        before = self.db.execute(
            "SELECT config_json,is_enabled,auto_submit FROM oms_warehouse_integrations WHERE provider='poland_wms'"
        ).fetchone()

        up_011(self.db)
        carrier = self.db.execute(
            "SELECT * FROM shipping_carriers WHERE slug='packeta-hu'"
        ).fetchone()
        integration = self.db.execute(
            "SELECT config_json,is_enabled,auto_submit FROM oms_warehouse_integrations WHERE provider='poland_wms'"
        ).fetchone()

        self.assertIn("tracking.expressone.hu", carrier["tracking_url"])
        self.assertEqual(0, integration["is_enabled"])
        self.assertEqual(0, integration["auto_submit"])
        self.assertEqual("managed_wms_skus", json.loads(integration["config_json"])["routing_policy"])

        down_011(self.db)
        restored = self.db.execute(
            "SELECT config_json,is_enabled,auto_submit FROM oms_warehouse_integrations WHERE provider='poland_wms'"
        ).fetchone()
        self.assertEqual(dict(before), dict(restored))
        self.assertIsNone(self.db.execute(
            "SELECT 1 FROM shipping_carriers WHERE slug='packeta-hu'"
        ).fetchone())

    def test_packeta_hu_primary_carrier_migration_roundtrip(self):
        up_009(self.db)
        up_010(self.db)
        self.db.execute(
            """CREATE TABLE IF NOT EXISTS shipping_carriers (
                 slug TEXT PRIMARY KEY, name TEXT, tracking_url TEXT, is_active INTEGER
               )"""
        )
        up_011(self.db)
        before = dict(self.db.execute(
            "SELECT slug,name,tracking_url,is_active FROM shipping_carriers WHERE slug='packeta-hu'"
        ).fetchone())

        up_012(self.db)
        updated = self.db.execute(
            "SELECT * FROM shipping_carriers WHERE slug='packeta-hu'"
        ).fetchone()
        self.assertIn("Packeta", updated["name"])
        self.assertIn("Express One", updated["name"])
        self.assertEqual(
            "https://tracking.packeta.com/en/{tracking}", updated["tracking_url"]
        )

        down_012(self.db)
        restored = dict(self.db.execute(
            "SELECT slug,name,tracking_url,is_active FROM shipping_carriers WHERE slug='packeta-hu'"
        ).fetchone())
        self.assertEqual(before, restored)

    def test_manual_partner_autoprovisions_czech_without_stock_quantity(self):
        self.db.execute(
            """UPDATE oms_warehouse_integrations
               SET inventory_authority='manual_partner' WHERE warehouse_id=1"""
        )
        self.set_stock(pl=0, hu=0)
        self.add_order(
            "cz-manual-1",
            "CZ",
            2,
            product_id=202,
            sku="",
            name="Partner-managed product",
        )

        result = plan_order(self.db, "cz-manual-1")

        self.assertFalse(result["shortages"])
        self.assertEqual([{"warehouse_id": 1, "qty": 2}], self.allocations("cz-manual-1"))
        self.assertIn("partner_reported_availability", result["reason"])
        created = self.db.execute(
            """SELECT s.sku_code, s.barcode, sw.warehouse_id
               FROM inv_site_sku_map sm
               JOIN inv_skus s ON s.id=sm.sku_id
               JOIN oms_sku_warehouses sw ON sw.sku_id=s.id
               WHERE sm.site_id=3 AND sm.wc_product_id=202"""
        ).fetchone()
        self.assertTrue(created["sku_code"].startswith("MANUAL-PL-3-202-"))
        self.assertIsNone(created["barcode"])
        self.assertEqual(1, created["warehouse_id"])

    def test_manual_partner_autoprovisions_hungary_market_fallback(self):
        self.db.execute(
            """UPDATE oms_warehouse_integrations
               SET inventory_authority='manual_partner' WHERE warehouse_id=1"""
        )
        self.add_order(
            "hu-unmapped-1",
            "HU",
            1,
            product_id=202,
            sku="",
            name="Partner-managed Hungary product",
        )

        result = plan_order(self.db, "hu-unmapped-1")

        self.assertFalse(result["shortages"])
        self.assertEqual([{"warehouse_id": 1, "qty": 1}], self.allocations("hu-unmapped-1"))
        created = self.db.execute(
            "SELECT id, barcode FROM inv_skus WHERE sku_code LIKE 'MANUAL-PL-2-202-%'"
        ).fetchone()
        self.assertIsNotNone(created)
        self.assertIsNone(created["barcode"])

    def test_packeta_and_manual_partner_migration_roundtrip(self):
        self.db.execute(
            """CREATE TABLE shipping_carriers (
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 slug TEXT NOT NULL UNIQUE,
                 name TEXT NOT NULL,
                 tracking_url TEXT,
                 is_active INTEGER DEFAULT 1
               )"""
        )

        up_009(self.db)

        carrier = self.db.execute(
            "SELECT * FROM shipping_carriers WHERE slug='packeta'"
        ).fetchone()
        authority = self.db.execute(
            "SELECT inventory_authority FROM oms_warehouse_integrations WHERE warehouse_id=1"
        ).fetchone()[0]
        self.assertEqual("https://tracking.packeta.com/en/{tracking}", carrier["tracking_url"])
        self.assertEqual("manual_partner", authority)

        down_009(self.db)

        self.assertIsNone(
            self.db.execute(
                "SELECT id FROM shipping_carriers WHERE slug='packeta'"
            ).fetchone()
        )
        authority = self.db.execute(
            "SELECT inventory_authority FROM oms_warehouse_integrations WHERE warehouse_id=1"
        ).fetchone()[0]
        self.assertEqual("local", authority)

    def test_new_poland_wms_migration_is_fail_closed_and_reversible(self):
        up_009(self.db)
        up_010(self.db)

        new_pl = self.db.execute(
            "SELECT id, name, is_active FROM warehouses WHERE code='PL-NEW-WMS'"
        ).fetchone()
        integration = self.db.execute(
            "SELECT * FROM oms_warehouse_integrations WHERE warehouse_id=?",
            (new_pl["id"],),
        ).fetchone()
        self.assertEqual("新波兰仓（API）", new_pl["name"])
        self.assertEqual("poland_wms", integration["provider"])
        self.assertEqual(0, integration["is_enabled"])
        self.assertEqual(0, integration["auto_submit"])
        self.assertEqual("manual_partner", integration["inventory_authority"])
        self.assertEqual("http://175.178.192.240:8082", integration["base_url"])
        self.assertEqual(
            3,
            self.db.execute(
                "SELECT COUNT(*) FROM inv_market_warehouses WHERE warehouse_id=? AND is_active=1",
                (new_pl["id"],),
            ).fetchone()[0],
        )
        self.assertEqual(
            0,
            self.db.execute(
                """SELECT COUNT(*) FROM inv_market_warehouses mw
                   JOIN oms_warehouse_integrations wi ON wi.warehouse_id=mw.warehouse_id
                   WHERE wi.provider='hungary_wms' AND mw.is_active=1"""
            ).fetchone()[0],
        )
        self.assertEqual(
            "0",
            self.db.execute(
                "SELECT value FROM settings WHERE key='oms_new_pl_wms_routing_enabled'"
            ).fetchone()[0],
        )

        down_010(self.db)
        self.assertIsNone(
            self.db.execute(
                "SELECT id FROM warehouses WHERE code='PL-NEW-WMS'"
            ).fetchone()
        )
        self.assertGreater(
            self.db.execute(
                """SELECT COUNT(*) FROM inv_market_warehouses mw
                   JOIN oms_warehouse_integrations wi ON wi.warehouse_id=mw.warehouse_id
                   WHERE wi.provider='hungary_wms' AND mw.is_active=1"""
            ).fetchone()[0],
            0,
        )
        down_009(self.db)

    def test_new_poland_wms_mapping_is_exclusive_and_never_falls_back(self):
        up_009(self.db)
        up_010(self.db)
        new_pl_id = self.db.execute(
            "SELECT id FROM warehouses WHERE code='PL-NEW-WMS'"
        ).fetchone()[0]
        self.db.execute(
            """UPDATE oms_warehouse_integrations
               SET is_enabled=1, base_url='https://pl-wms.example.test',
                   external_code='PL01', channel_code='777'
               WHERE warehouse_id=?""",
            (new_pl_id,),
        )
        self.db.execute(
            "UPDATE settings SET value='1' WHERE key='oms_new_pl_wms_routing_enabled'"
        )
        self.db.execute(
            """INSERT INTO oms_sku_warehouses
               (sku_id, warehouse_id, is_primary, is_enabled,
                wms_product_name_zh, wms_product_name_en, product_type)
               VALUES (1, ?, 1, 1, '测试产品', 'Test product', 'P')""",
            (new_pl_id,),
        )
        self.db.execute(
            """INSERT INTO oms_external_stock
               (warehouse_id, sku_barcode, sku_id, quantity,
                lock_quantity, available_quantity)
               VALUES (?, 'BAR1', 1, 10, 0, 10)""",
            (new_pl_id,),
        )
        self.add_order("new-pl-special-1", "CZ", 2)

        result = plan_order(self.db, "new-pl-special-1")

        self.assertEqual(
            [{"warehouse_id": new_pl_id, "qty": 2}],
            self.allocations("new-pl-special-1"),
        )
        self.assertIn("exclusive_sku_route", result["reason"])
        fulfillment = self.db.execute(
            "SELECT * FROM oms_fulfillments WHERE order_id='new-pl-special-1'"
        ).fetchone()
        self.assertEqual("poland_wms", fulfillment["provider"])
        self.assertEqual("external_wms", fulfillment["mode"])
        with self.assertRaisesRegex(Exception, "独立 API 适配器"):
            build_wms_payload(self.db, fulfillment["id"])
        payload = build_poland_wms_payload(
            self.db,
            fulfillment["id"],
            customer_id="C1",
            customer_userid="U1",
        )
        self.assertEqual("777", payload["product_id"])
        self.assertEqual("CZ", payload["country"])
        self.assertEqual("OMSnewplspecial1PL01R1", payload["order_customerinvoicecode"])
        self.assertEqual("BAR1", payload["orderInvoiceParam"][0]["sku_code"])
        self.assertEqual(Decimal("20"), Decimal(payload["order_codamount"]))

    def test_new_poland_wms_special_sku_uses_partner_reported_stock_only(self):
        up_009(self.db)
        up_010(self.db)
        new_pl_id = self.db.execute(
            "SELECT id FROM warehouses WHERE code='PL-NEW-WMS'"
        ).fetchone()[0]
        self.db.execute(
            """UPDATE oms_warehouse_integrations
               SET is_enabled=1, base_url='https://pl-wms.example.test',
                   external_code='PL01'
               WHERE warehouse_id=?""",
            (new_pl_id,),
        )
        self.db.execute(
            "UPDATE settings SET value='1' WHERE key='oms_new_pl_wms_routing_enabled'"
        )
        self.db.execute(
            """INSERT INTO oms_sku_warehouses
               (sku_id, warehouse_id, is_primary, is_enabled,
                wms_product_name_zh, wms_product_name_en, product_type)
               VALUES (1, ?, 1, 1, '测试产品', 'Test product', 'P')""",
            (new_pl_id,),
        )
        self.add_order("new-pl-shortage-1", "HU", 1)

        result = plan_order(self.db, "new-pl-shortage-1")

        self.assertEqual(
            [{"warehouse_id": new_pl_id, "qty": 1}],
            self.allocations("new-pl-shortage-1"),
        )
        self.assertFalse(result["shortages"])
        self.assertIn("partner_reported_availability", result["reason"])

    def test_paused_hungary_only_mapping_falls_back_to_manual_poland(self):
        self.db.execute(
            "DELETE FROM oms_sku_warehouses WHERE sku_id=1 AND warehouse_id=1"
        )
        up_009(self.db)
        up_010(self.db)
        self.add_order("hu-legacy-map-1", "HU", 1)

        result = plan_order(self.db, "hu-legacy-map-1")

        self.assertFalse(result["shortages"])
        self.assertEqual(
            [{"warehouse_id": 1, "qty": 1}],
            self.allocations("hu-legacy-map-1"),
        )
        self.assertIn("partner_reported_availability", result["reason"])

    def test_two_parcels_complete_only_after_both_delivered_and_ignore_late_event(self):
        self.set_stock(pl=1, hu=5)
        self.add_order("delivery-1", "PL", 3)
        plan_order(self.db, "delivery-1")
        fulfillments = self.db.execute(
            "SELECT id,warehouse_id FROM oms_fulfillments WHERE order_id='delivery-1' ORDER BY warehouse_id"
        ).fetchall()
        shipments = []
        for f in fulfillments:
            current = self.db.execute("SELECT status FROM oms_fulfillments WHERE id=?", (f["id"],)).fetchone()[0]
            if current == "ready_to_submit":
                transition_fulfillment(self.db, f["id"], "submitting")
                transition_fulfillment(self.db, f["id"], "accepted")
            shipments.append(create_shipment(self.db, f["id"], f"TRACK-{f['warehouse_id']}", carrier_slug="test"))
        add_tracking_event(self.db, shipments[0]["id"], "official", "delivered", raw_status="delivered", external_event_id="one-delivered")
        recompute_order_status(self.db, "delivery-1")
        self.assertFalse(completion_guard(self.db, "delivery-1")[0])
        add_tracking_event(self.db, shipments[1]["id"], "third_party", "delivered", raw_status="delivered", external_event_id="two-delivered")
        recompute_order_status(self.db, "delivery-1")
        self.assertTrue(completion_guard(self.db, "delivery-1")[0])

        # Duplicate and late out-of-order callbacks are audited but cannot regress.
        add_tracking_event(self.db, shipments[1]["id"], "third_party", "delivered", raw_status="delivered", external_event_id="two-delivered")
        add_tracking_event(self.db, shipments[1]["id"], "official", "in_transit", raw_status="transit", external_event_id="late-transit")
        status = self.db.execute("SELECT status FROM oms_shipments WHERE id=?", (shipments[1]["id"],)).fetchone()[0]
        self.assertEqual("delivered", status)
        self.assertEqual(2, self.db.execute("SELECT COUNT(*) FROM oms_tracking_events WHERE shipment_id=?", (shipments[1]["id"],)).fetchone()[0])

    def test_wms_payload_contract_and_cod(self):
        self.add_order("hu-wms-1", "HU", 1, shipping_total="3.49")
        result = plan_order(self.db, "hu-wms-1")
        payload = build_wms_payload(self.db, result["fulfillment_ids"][0])
        self.assertEqual("HU01", payload["storehouseCode"])
        self.assertEqual("欧洲直发-25", payload["channelCode"])
        self.assertEqual("匈牙利", payload["contry"])
        self.assertEqual(Decimal("13.49"), Decimal(payload["invoicePrice"]))
        self.assertEqual("测试产品", payload["invoiceDetailsCreateRequests"][0]["productName"])
        finance = self.db.execute(
            "SELECT * FROM oms_fulfillment_financials WHERE fulfillment_id=?",
            (result["fulfillment_ids"][0],),
        ).fetchone()
        self.assertEqual(Decimal("10"), Decimal(finance["merchandise_amount"]))
        self.assertEqual(Decimal("3.49"), Decimal(finance["customer_shipping_amount"]))

    def test_split_cod_allocates_shipping_to_poland_and_hungary_goods_only(self):
        self.set_stock(pl=1, hu=10)
        self.add_order("split-cod-1", "PL", 3, shipping_total="3.49")
        plan_order(self.db, "split-cod-1")
        rows = self.db.execute(
            '''SELECT f.id, w.country, ff.cod_collection_role, ff.cod_amount,
                      ff.merchandise_amount, ff.customer_shipping_amount,
                      ff.order_adjustment_amount, ff.source_order_total,
                      ff.source_shipping_total, ff.allocation_method,
                      ff.settlement_mode
               FROM oms_fulfillments f
               JOIN warehouses w ON w.id=f.warehouse_id
               JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
               WHERE f.order_id='split-cod-1' AND f.status!='superseded'
               ORDER BY w.country'''
        ).fetchall()
        by_country = {row["country"]: row for row in rows}
        self.assertEqual("collector", by_country["PL"]["cod_collection_role"])
        self.assertEqual(Decimal("13.49"), Decimal(by_country["PL"]["cod_amount"]))
        self.assertEqual(Decimal("10"), Decimal(by_country["PL"]["merchandise_amount"]))
        self.assertEqual(Decimal("3.49"), Decimal(by_country["PL"]["customer_shipping_amount"]))
        self.assertEqual("collector", by_country["HU"]["cod_collection_role"])
        self.assertEqual(Decimal("20"), Decimal(by_country["HU"]["cod_amount"]))
        self.assertEqual(Decimal("20"), Decimal(by_country["HU"]["merchandise_amount"]))
        self.assertEqual(Decimal("0"), Decimal(by_country["HU"]["customer_shipping_amount"]))
        self.assertEqual(
            Decimal("33.49"),
            sum(Decimal(row["cod_amount"]) for row in rows),
        )
        self.assertEqual(
            "woo_line_gross_residual_to_poland",
            by_country["HU"]["allocation_method"],
        )
        self.assertEqual("monthly_statement", by_country["HU"]["settlement_mode"])
        payload = build_wms_payload(self.db, by_country["HU"]["id"])
        self.assertEqual(Decimal("20"), Decimal(payload["invoicePrice"]))

    def test_huf_discount_tax_and_shipping_preserve_order_total(self):
        self.set_stock(pl=1, hu=10)
        self.add_order(
            "split-huf-1",
            "PL",
            3,
            shipping_total="3490",
            currency="HUF",
            line_total="27000",
            line_tax="3000",
            order_total="33490",
        )
        plan_order(self.db, "split-huf-1")
        rows = self.db.execute(
            '''SELECT w.country, ff.*
               FROM oms_fulfillments f
               JOIN warehouses w ON w.id=f.warehouse_id
               JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
               WHERE f.order_id='split-huf-1' AND f.status!='superseded'
               ORDER BY w.country'''
        ).fetchall()
        by_country = {row["country"]: row for row in rows}
        self.assertEqual(Decimal("13490"), Decimal(by_country["PL"]["cod_amount"]))
        self.assertEqual(Decimal("10000"), Decimal(by_country["PL"]["merchandise_amount"]))
        self.assertEqual(Decimal("3490"), Decimal(by_country["PL"]["customer_shipping_amount"]))
        self.assertEqual(Decimal("20000"), Decimal(by_country["HU"]["cod_amount"]))
        self.assertEqual(Decimal("0"), Decimal(by_country["HU"]["customer_shipping_amount"]))
        self.assertEqual(
            Decimal("33490"),
            sum(Decimal(row["cod_amount"]) for row in rows),
        )
        order = self.db.execute(
            "SELECT * FROM orders WHERE id='split-huf-1'"
        ).fetchone()
        hu_notice = _customer_note_body(
            self.db,
            order,
            {
                "fulfillment_id": by_country["HU"]["fulfillment_id"],
                "tracking_number": "HU-TRACK-1",
                "carrier_name": "GLS",
                "carrier_slug": "gls",
            },
        )
        self.assertIn("Kwota pobrania dla tej przesyłki: 20000.00 HUF", hu_notice)
        self.assertIn("koszt dostawy nie zostanie naliczony ponownie", hu_notice)
        pl_notice = _customer_note_body(
            self.db,
            order,
            {
                "fulfillment_id": by_country["PL"]["fulfillment_id"],
                "tracking_number": "PL-TRACK-1",
                "carrier_name": "DPD",
                "carrier_slug": "dpd",
            },
        )
        self.assertIn("Kwota pobrania dla tej przesyłki: 13490.00 HUF", pl_notice)
        self.assertIn("pełny koszt dostawy zamówienia w wysokości 3490.00 HUF", pl_notice)

        expected = {
            "PL": "Przesyłka dla Twojego zamówienia",
            "HU": "A rendeléséhez tartozó csomagot",
            "CZ": "Zásilka k vaší objednávce",
            "DE": "A parcel for your order",
        }
        for country, phrase in expected.items():
            self.db.execute("UPDATE sites SET country=? WHERE id=1", (country,))
            notice = _customer_note_body(
                self.db,
                order,
                {
                    "fulfillment_id": by_country["PL"]["fulfillment_id"],
                    "tracking_number": "LANG-TRACK-1",
                    "carrier_name": "DPD",
                    "carrier_slug": "dpd",
                },
            )
            self.assertIn(phrase, notice)
            self.assertIsNone(re.search(r"[\u4e00-\u9fff]", notice))

    def test_legacy_customer_note_uses_site_language_and_english_fallback(self):
        expected = {
            "PL": "Przesyłka dla Twojego zamówienia",
            "HU": "A rendeléséhez tartozó csomagot",
            "CZ": "Zásilka k vaší objednávce",
            "SK": "A parcel for your order",
        }
        for country, phrase in expected.items():
            note = basic_shipment_note(
                country,
                "Packeta",
                "Z1465635854",
                "https://tracking.packeta.com/en/Z1465635854",
            )
            self.assertIn(phrase, note)
            self.assertIn("Z1465635854", note)
            self.assertIsNone(re.search(r"[\u4e00-\u9fff]", note))

    def test_split_rounding_keeps_exact_cod_total(self):
        self.set_stock(pl=1, hu=10)
        self.add_order(
            "split-round-1",
            "PL",
            3,
            shipping_total="3.49",
            line_total="10.00",
            order_total="13.49",
        )
        plan_order(self.db, "split-round-1")
        amounts = {
            row["country"]: Decimal(row["cod_amount"])
            for row in self.db.execute(
                '''SELECT w.country, ff.cod_amount
                   FROM oms_fulfillments f
                   JOIN warehouses w ON w.id=f.warehouse_id
                   JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
                   WHERE f.order_id='split-round-1' AND f.status!='superseded' '''
            ).fetchall()
        }
        self.assertEqual(Decimal("6.82"), amounts["PL"])
        self.assertEqual(Decimal("6.67"), amounts["HU"])
        self.assertEqual(Decimal("13.49"), sum(amounts.values()))

    def test_order_level_fee_stays_with_poland_collector(self):
        self.set_stock(pl=1, hu=10)
        self.add_order(
            "split-fee-1",
            "PL",
            3,
            shipping_total="3.49",
            line_total="30",
            order_total="35.49",
        )
        plan_order(self.db, "split-fee-1")
        rows = self.db.execute(
            '''SELECT w.country, ff.cod_amount, ff.customer_shipping_amount,
                      ff.order_adjustment_amount
               FROM oms_fulfillments f
               JOIN warehouses w ON w.id=f.warehouse_id
               JOIN oms_fulfillment_financials ff ON ff.fulfillment_id=f.id
               WHERE f.order_id='split-fee-1' AND f.status!='superseded' '''
        ).fetchall()
        by_country = {row["country"]: row for row in rows}
        self.assertEqual(Decimal("15.49"), Decimal(by_country["PL"]["cod_amount"]))
        self.assertEqual(Decimal("3.49"), Decimal(by_country["PL"]["customer_shipping_amount"]))
        self.assertEqual(Decimal("2.00"), Decimal(by_country["PL"]["order_adjustment_amount"]))
        self.assertEqual(Decimal("20.00"), Decimal(by_country["HU"]["cod_amount"]))
        self.assertEqual(Decimal("35.49"), sum(Decimal(row["cod_amount"]) for row in rows))

    def test_prepaid_hungary_never_receives_cod_amount(self):
        self.add_order(
            "hu-prepaid-1",
            "HU",
            1,
            shipping_total="3.49",
            payment_method="bacs",
        )
        result = plan_order(self.db, "hu-prepaid-1")
        payload = build_wms_payload(self.db, result["fulfillment_ids"][0])
        finance = self.db.execute(
            "SELECT * FROM oms_fulfillment_financials WHERE fulfillment_id=?",
            (result["fulfillment_ids"][0],),
        ).fetchone()
        self.assertEqual("not_applicable", finance["cod_collection_role"])
        self.assertEqual("monthly_statement", finance["settlement_mode"])
        self.assertEqual(Decimal("10"), Decimal(finance["merchandise_amount"]))
        self.assertEqual(Decimal("3.49"), Decimal(finance["customer_shipping_amount"]))
        self.assertEqual(Decimal("0"), Decimal(payload["invoicePrice"]))

    def test_ast_excludes_wms_label_until_actual_outbound(self):
        self.set_stock(pl=1, hu=5)
        self.add_order("ast-1", "PL", 3)
        plan_order(self.db, "ast-1")
        order_item = self.db.execute(
            "SELECT id, raw_json FROM oms_order_items WHERE order_id='ast-1'"
        ).fetchone()
        raw_item = json.loads(order_item["raw_json"])
        raw_item.update({
            "product_id": 13792,
            "variation_id": 13794,
            "sku": "WATERMELON-ICE",
        })
        self.db.execute(
            "UPDATE oms_order_items SET raw_json=? WHERE id=?",
            (json.dumps(raw_item), order_item["id"]),
        )
        fulfillments = self.db.execute(
            "SELECT id,warehouse_id,status FROM oms_fulfillments WHERE order_id='ast-1' ORDER BY warehouse_id"
        ).fetchall()
        pl, hu = fulfillments
        create_shipment(self.db, pl["id"], "PL-SHIPPED", carrier_slug="inpost")
        transition_fulfillment(self.db, hu["id"], "submitting")
        transition_fulfillment(self.db, hu["id"], "accepted")
        create_shipment(
            self.db, hu["id"], "HU-NOT-OUTBOUND", carrier_slug="wms-auto",
            initial_status="label_ready",
        )
        shipments = _all_order_shipments(self.db, "ast-1", 1)
        self.assertEqual(["PL-SHIPPED"], [s["tracking_number"] for s in shipments])
        partial_items = _ast_items(shipments)
        self.assertEqual(1, len(partial_items))
        self.assertEqual("2", partial_items[0]["status_shipped"])
        self.assertEqual("13794", partial_items[0]["products_list"][0]["product"])
        self.assertEqual("WATERMELON-ICE", partial_items[0]["products_list"][0]["sku"])
        self.assertEqual("1", _ast_items(shipments, final=True)[0]["status_shipped"])

    def test_ast_hungary_packeta_uses_expressone_custom_link(self):
        self.db.execute(
            """CREATE TABLE IF NOT EXISTS shipping_carriers (
                 slug TEXT PRIMARY KEY, name TEXT, tracking_url TEXT, is_active INTEGER
               )"""
        )
        self.db.execute(
            """INSERT OR REPLACE INTO shipping_carriers (slug,name,tracking_url,is_active)
               VALUES ('packeta-hu','Packeta / Express One（匈牙利）',
                       'https://tracking.expressone.hu/?plc_number={tracking}',1)"""
        )
        items = _ast_items([{
            "tracking_number": "671555557697000013601086",
            "carrier_slug": "packeta-hu",
            "carrier_name": "Packeta / Express One（匈牙利）",
            "shipped_at": "2026-07-30 10:48:00",
            "products": [],
        }], conn=self.db)

        self.assertEqual("custom", items[0]["tracking_provider"])
        self.assertEqual(
            "https://tracking.expressone.hu/?plc_number=671555557697000013601086",
            items[0]["custom_tracking_link"],
        )

        packeta_items = _ast_items([{
            "tracking_number": "Z1465635854",
            "carrier_slug": "packeta-hu",
            "carrier_name": "Packeta（匈牙利，Express One 末端派送）",
            "shipped_at": "2026-07-30 10:48:00",
            "products": [],
        }], conn=self.db)
        self.assertEqual("packeta", packeta_items[0]["tracking_provider"])
        self.assertEqual("", packeta_items[0]["custom_tracking_link"])


if __name__ == "__main__":
    unittest.main()
