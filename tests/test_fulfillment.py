import json
import sqlite3
import unittest
from decimal import Decimal

from fulfillment_service import (
    add_tracking_event,
    build_wms_payload,
    completion_guard,
    create_shipment,
    plan_order,
    recompute_order_status,
    transition_fulfillment,
)
from inv_migrations import up_001, up_006, up_007, up_008
from fulfillment_woocommerce import _all_order_shipments, _ast_items, _customer_note_body


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
    ):
        site = {"PL": "https://pl.test", "HU": "https://hu.test", "CZ": "https://cz.test"}[market]
        if line_total is None:
            line_total = qty * 10
        if order_total is None:
            order_total = Decimal(str(line_total)) + Decimal(str(line_tax)) + Decimal(str(shipping_total))
        item = {
            "id": 501,
            "product_id": 101,
            "variation_id": 0,
            "sku": "SKU1",
            "name": "Test product",
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
        self.assertIn("货到付款金额：20000.00 HUF", hu_notice)
        self.assertIn("不重复收取订单运费", hu_notice)
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
        self.assertIn("货到付款金额：13490.00 HUF", pl_notice)
        self.assertIn("全部客户运费 3490.00 HUF", pl_notice)

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
        self.assertEqual(1, len(_ast_items(shipments)))


if __name__ == "__main__":
    unittest.main()
