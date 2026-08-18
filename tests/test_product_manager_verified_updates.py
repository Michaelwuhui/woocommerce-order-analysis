import json
import os
import unittest

import product_manager_service as service


class FakeResponse:
    def __init__(self, payload, status_code=200):
        self._payload = payload
        self.status_code = status_code
        self.text = json.dumps(payload)

    def json(self):
        return self._payload


class FakeWooRequests:
    RequestException = Exception

    def __init__(self, *, ignore_status=False, wcms_stock_bridge=False):
        self.calls = []
        self.put_urls = []
        self.ignore_status = ignore_status
        self.wcms_stock_bridge = wcms_stock_bridge
        self.state = {
            "id": 101,
            "name": "Variable",
            "slug": "variable",
            "sku": "SKU-101",
            "type": "variation",
            "manage_stock": True,
            "stock_quantity": 200,
            "stock_status": "instock",
            "regular_price": "10.00",
            "sale_price": "",
            "price": "10.00",
        }
        if wcms_stock_bridge:
            self.state["meta_data"] = [
                {"key": "wcms_stock_manage", "value": "yes"},
                {"key": "wcms_stock_qty", "value": 200},
                {"key": "wcms_stock_status", "value": "instock"},
            ]

    def put(self, url, *, auth, json, timeout, headers):
        self.calls.append(("PUT", dict(json)))
        self.put_urls.append(url)
        if "manage_stock" in json:
            self.state["manage_stock"] = bool(json["manage_stock"])
            # Reproduce the WC behavior that caused the production defect:
            # positive legacy quantity keeps the first response in stock.
            if json["manage_stock"] is False:
                self.state["stock_status"] = "instock"
        if "stock_status" in json and not self.ignore_status:
            self.state["stock_status"] = json["stock_status"]
        if "meta_data" in json:
            current = {
                row["key"]: row["value"]
                for row in self.state.get("meta_data", [])
            }
            current.update({row["key"]: row["value"] for row in json["meta_data"]})
            self.state["meta_data"] = [
                {"key": key, "value": value}
                for key, value in current.items()
            ]
        if self.wcms_stock_bridge:
            current = {
                row["key"]: row["value"]
                for row in self.state.get("meta_data", [])
            }
            self.state["manage_stock"] = current["wcms_stock_manage"] == "yes"
            self.state["stock_quantity"] = (
                int(current["wcms_stock_qty"])
                if self.state["manage_stock"] else None
            )
            self.state["stock_status"] = current["wcms_stock_status"]
        return FakeResponse(dict(self.state))

    def get(self, url, *, auth, timeout, headers, params=None):
        self.calls.append(("GET", dict(params or {})))
        if params is not None:
            return FakeResponse([dict(self.state)])
        return FakeResponse(dict(self.state))


class ProductManagerVerifiedUpdateTests(unittest.TestCase):
    def test_hard_soldout_is_split_and_verified(self):
        req = FakeWooRequests()
        final, error, trace = service.wc_product_update_verified(
            req,
            "https://shop.test/wp-json/wc/v3/products/1/variations/101",
            ("ck", "cs"),
            {"manage_stock": False, "stock_status": "outofstock"},
        )

        self.assertIsNone(error)
        self.assertEqual(final["stock_status"], "outofstock")
        self.assertFalse(final["manage_stock"])
        self.assertEqual(
            req.calls,
            [
                ("PUT", {"manage_stock": False}),
                ("PUT", {"stock_status": "outofstock"}),
                ("GET", {}),
            ],
        )
        self.assertEqual(len(trace["phases"]), 2)

    def test_http_200_is_failure_when_readback_does_not_match(self):
        req = FakeWooRequests(ignore_status=True)
        final, error, trace = service.wc_product_update_verified(
            req,
            "https://shop.test/wp-json/wc/v3/products/1/variations/101",
            ("ck", "cs"),
            {"manage_stock": False, "stock_status": "outofstock"},
        )

        self.assertEqual(final["stock_status"], "instock")
        self.assertIn("写入未达到目标状态", error)
        self.assertIn("stock_status", error)
        self.assertEqual(trace["final_state"]["stock_status"], "instock")

    def test_child_sync_is_verified_by_sku(self):
        req = FakeWooRequests()
        req.state.update({
            "manage_stock": False,
            "stock_status": "outofstock",
        })
        site = {
            "url": "https://child.test",
            "consumer_key": "child-ck",
            "consumer_secret": "child-cs",
            "product_master_id": 2,
        }
        result = service.verify_product_child_sync(
            req,
            site,
            dict(req.state),
            {"manage_stock": False, "stock_status": "outofstock"},
        )

        self.assertEqual(result["status"], "verified")
        self.assertEqual(req.calls[-1][1]["sku"], "SKU-101")
        self.assertFalse(result.get("direct_update", False))

    def test_child_sync_mismatch_is_directly_repaired_and_verified(self):
        req = FakeWooRequests()
        req.state["parent_id"] = 77
        site = {
            "url": "https://child.test",
            "consumer_key": "child-ck",
            "consumer_secret": "child-cs",
            "product_master_id": 2,
        }
        master_item = dict(req.state)
        master_item.update({
            "parent_id": 77,
            "manage_stock": False,
            "stock_status": "outofstock",
        })

        result = service.verify_product_child_sync(
            req,
            site,
            master_item,
            {"manage_stock": False, "stock_status": "outofstock"},
        )

        self.assertEqual(result["status"], "verified")
        self.assertTrue(result["direct_update"])
        self.assertFalse(result["state"]["manage_stock"])
        self.assertEqual(result["state"]["stock_status"], "outofstock")
        self.assertEqual(
            req.put_urls,
            [
                "https://child.test/wp-json/wc/v3/products/77/variations/101",
                "https://child.test/wp-json/wc/v3/products/77/variations/101",
            ],
        )

    def test_child_sync_reports_failure_when_direct_write_does_not_converge(self):
        req = FakeWooRequests(ignore_status=True)
        req.state["parent_id"] = 77
        site = {
            "url": "https://child.test",
            "consumer_key": "child-ck",
            "consumer_secret": "child-cs",
            "product_master_id": 2,
        }
        master_item = dict(req.state)
        master_item.update({
            "parent_id": 77,
            "manage_stock": False,
            "stock_status": "outofstock",
        })

        result = service.verify_product_child_sync(
            req,
            site,
            master_item,
            {"manage_stock": False, "stock_status": "outofstock"},
        )

        self.assertEqual(result["status"], "error")
        self.assertTrue(result["direct_update"])
        self.assertIn("直接补写也失败", result["detail"])
        self.assertEqual(result["state"]["stock_status"], "instock")

    def test_child_sync_updates_wcms_stock_bridge_metadata(self):
        req = FakeWooRequests(wcms_stock_bridge=True)
        req.state["parent_id"] = 77
        site = {
            "url": "https://child.test",
            "consumer_key": "child-ck",
            "consumer_secret": "child-cs",
            "product_master_id": 2,
        }
        master_item = dict(req.state)
        master_item.update({
            "manage_stock": False,
            "stock_quantity": None,
            "stock_status": "outofstock",
        })

        result = service.verify_product_child_sync(
            req,
            site,
            master_item,
            {"manage_stock": False, "stock_status": "outofstock"},
        )

        self.assertEqual(result["status"], "verified")
        self.assertFalse(result["state"]["manage_stock"])
        self.assertEqual(result["state"]["stock_status"], "outofstock")
        first_phase = next(call[1] for call in req.calls if call[0] == "PUT")
        bridge_meta = {
            row["key"]: row["value"]
            for row in first_phase["meta_data"]
        }
        self.assertEqual(bridge_meta["wcms_stock_manage"], "no")
        self.assertEqual(bridge_meta["wcms_stock_qty"], 0)
        self.assertEqual(bridge_meta["wcms_stock_status"], "outofstock")

if __name__ == "__main__":
    unittest.main()
