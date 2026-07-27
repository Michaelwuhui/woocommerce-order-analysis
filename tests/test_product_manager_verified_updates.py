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

    def __init__(self, *, ignore_status=False):
        self.calls = []
        self.ignore_status = ignore_status
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

    def put(self, url, *, auth, json, timeout, headers):
        self.calls.append(("PUT", dict(json)))
        if "manage_stock" in json:
            self.state["manage_stock"] = bool(json["manage_stock"])
            # Reproduce the WC behavior that caused the production defect:
            # positive legacy quantity keeps the first response in stock.
            if json["manage_stock"] is False:
                self.state["stock_status"] = "instock"
        if "stock_status" in json and not self.ignore_status:
            self.state["stock_status"] = json["stock_status"]
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

if __name__ == "__main__":
    unittest.main()
