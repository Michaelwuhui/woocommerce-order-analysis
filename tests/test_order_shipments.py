import json
import unittest

from order_shipments import build_shipping_log_parcels, extract_tracking_candidates


class OrderShipmentPresentationTests(unittest.TestCase):
    def setUp(self):
        self.lines = [
            {"id": 4034, "product_id": 13792, "variation_id": 13802, "name": "LEMON LIME", "sku": "LEMON", "quantity": 10},
            {"id": 4035, "product_id": 13792, "variation_id": 13794, "name": "WATERMELON ICE", "sku": "WATERMELON", "quantity": 10},
            {"id": 4036, "product_id": 13792, "variation_id": 13796, "name": "JUICY PEACH ICE", "sku": "PEACH", "quantity": 10},
        ]

    def test_builds_all_parcels_with_exact_variant_items(self):
        logs = [
            {
                "id": 1,
                "tracking_number": "DPD-ONE",
                "carrier_slug": "dpd",
                "shipped_at": "2026-08-06 10:00:00",
                "items_json": json.dumps([{"product": "13794", "item_id": "4035", "qty": "10"}]),
                "is_partial": 1,
                "is_reship": 0,
            },
            {
                "id": 2,
                "tracking_number": "DPD-TWO",
                "carrier_slug": "dpd",
                "shipped_at": "2026-08-07 10:00:00",
                "items_json": json.dumps([
                    {"product": "13802", "item_id": "4034", "qty": "10"},
                    {"product": "13796", "item_id": "4036", "qty": "10"},
                ]),
                "is_partial": 0,
                "is_reship": 0,
            },
        ]

        parcels = build_shipping_log_parcels(logs, self.lines)

        self.assertEqual(["DPD-ONE", "DPD-TWO"], [p["tracking_number"] for p in parcels])
        self.assertEqual("WATERMELON ICE", parcels[0]["items"][0]["name"])
        self.assertEqual(10, parcels[0]["items"][0]["qty"])
        self.assertTrue(parcels[0]["is_partial"])
        self.assertEqual(["LEMON LIME", "JUICY PEACH ICE"], [i["name"] for i in parcels[1]["items"]])

    def test_collects_every_ast_tracking_and_deduplicates_local_copy(self):
        meta = [{
            "key": "_wc_shipment_tracking_items",
            "value": [
                {"tracking_number": "DPD-ONE", "tracking_provider": "dpd-pl"},
                {"tracking_number": "DPD-TWO", "tracking_provider": "dpd-pl"},
            ],
        }]
        logs = [
            {"tracking_number": "dpd-one", "carrier_slug": "dpd"},
            {"tracking_number": "DPD-TWO", "carrier_slug": "dpd"},
        ]

        candidates = extract_tracking_candidates(meta, self.lines, [], logs)

        self.assertEqual(["DPD-ONE", "DPD-TWO"], [c["tracking_number"] for c in candidates])
        self.assertEqual(["dpd-pl", "dpd-pl"], [c["provider"] for c in candidates])

    def test_collects_multiple_villatheme_trackings(self):
        lines = [{
            "id": 1,
            "meta_data": [{
                "key": "_vi_wot_order_item_tracking_data",
                "value": json.dumps([
                    {"tracking_number": "VILLA-1", "carrier_slug": "gls"},
                    {"tracking_number": "VILLA-2", "carrier_name": "DPD"},
                ]),
            }],
        }]

        candidates = extract_tracking_candidates([], lines, [], [])

        self.assertEqual(
            [
                {"tracking_number": "VILLA-1", "provider": "gls"},
                {"tracking_number": "VILLA-2", "provider": "DPD"},
            ],
            candidates,
        )


if __name__ == "__main__":
    unittest.main()
