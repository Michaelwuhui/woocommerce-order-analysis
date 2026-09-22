import json
import os
import unittest
from unittest.mock import patch

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
                ("GET", {}),
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

    def test_child_read_timeout_is_retried_without_replaying_completed_write(self):
        req = FakeWooRequests()
        req.state.update(manage_stock=False, stock_status="outofstock")
        site = {"url": "https://child.test", "consumer_key": "ck",
                "consumer_secret": "cs", "product_master_id": 2}
        with patch.object(req, "get", side_effect=[
            req.RequestException("read timed out"), FakeResponse([dict(req.state)])
        ]) as get:
            result = service.verify_product_child_sync(
                req, site, dict(req.state),
                {"manage_stock": False, "stock_status": "outofstock"},
            )
        self.assertEqual(result["status"], "verified")
        self.assertEqual([c.kwargs["timeout"] for c in get.call_args_list], [(5, 20), (5, 15)])
        self.assertEqual(get.call_args_list[0].args, get.call_args_list[1].args)
        self.assertEqual(get.call_args_list[0].kwargs["params"], get.call_args_list[1].kwargs["params"])
        self.assertEqual(req.put_urls, [])

    def test_repeated_child_read_failure_stays_pending_without_writes(self):
        req = FakeWooRequests()
        site = {"url": "https://child.test", "consumer_key": "ck",
                "consumer_secret": "cs", "product_master_id": 2}
        with patch.object(req, "get", side_effect=req.RequestException("read timed out")) as get:
            result = service.verify_product_child_sync(
                req, site, dict(req.state),
                {"manage_stock": False, "stock_status": "outofstock"},
            )
        self.assertEqual(result["status"], "pending")
        self.assertEqual(get.call_count, 2)
        self.assertEqual(req.put_urls, [])

    def test_child_auth_failure_is_not_retried_or_written(self):
        req = FakeWooRequests()
        site = {"url": "https://child.test", "consumer_key": "ck",
                "consumer_secret": "cs", "product_master_id": 2}
        with patch.object(req, "get", return_value=FakeResponse({"message": "Unauthorized"}, 401)) as get:
            result = service.verify_product_child_sync(
                req, site, dict(req.state),
                {"manage_stock": False, "stock_status": "outofstock"},
            )
        self.assertEqual(result["status"], "pending")
        self.assertEqual(get.call_count, 1)
        self.assertEqual(req.put_urls, [])

    def test_child_gateway_failure_retries_read_only(self):
        req = FakeWooRequests()
        req.state.update(manage_stock=False, stock_status="outofstock")
        site = {"url": "https://child.test", "consumer_key": "ck",
                "consumer_secret": "cs", "product_master_id": 2}
        with patch.object(req, "get", side_effect=[
            FakeResponse({"message": "SSL handshake failed"}, 525),
            FakeResponse([dict(req.state)]),
        ]) as get:
            result = service.verify_product_child_sync(
                req, site, dict(req.state),
                {"manage_stock": False, "stock_status": "outofstock"},
            )
        self.assertEqual(result["status"], "verified")
        self.assertEqual(get.call_count, 2)
        self.assertEqual(req.put_urls, [])

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

    def test_already_soldout_is_verified_without_any_put(self):
        req = FakeWooRequests()
        req.state.update(manage_stock=False, stock_status="outofstock")
        final, error, trace = service.wc_product_update_verified(
            req, "https://shop.test/wp-json/wc/v3/products/101", ("ck", "cs"),
            {"manage_stock": False, "stock_status": "outofstock"},
        )
        self.assertIsNone(error)
        self.assertEqual(req.calls, [("GET", {})])
        self.assertTrue(trace["already_satisfied"])
        self.assertEqual(final["stock_status"], "outofstock")

    def test_applied_write_timeout_is_read_back_without_replaying_phase(self):
        req = FakeWooRequests()
        put = req.put
        def timeout_after_save(*args, **kwargs):
            put(*args, **kwargs)
            raise req.RequestException("remote hook still running after save")
        with patch.object(req, "put", side_effect=timeout_after_save) as writes:
            final, error, trace = service.wc_product_update_verified(
                req, "https://shop.test/wp-json/wc/v3/products/101", ("ck", "cs"),
                {"manage_stock": False, "stock_status": "outofstock"},
            )
        self.assertIsNone(error)
        self.assertEqual(writes.call_count, 2)
        self.assertEqual([p[1] for p in req.calls if p[0] == "PUT"],
                         [{"manage_stock": False}, {"stock_status": "outofstock"}])
        self.assertTrue(all(p["verified_after_error"] for p in trace["phases"]))
        self.assertEqual(final["stock_status"], "outofstock")

    def test_unconfirmed_first_phase_stops_without_replay_or_second_phase(self):
        req = FakeWooRequests()
        with patch.object(req, "put", side_effect=req.RequestException("connection lost")) as writes:
            final, error, trace = service.wc_product_update_verified(
                req, "https://shop.test/wp-json/wc/v3/products/101", ("ck", "cs"),
                {"manage_stock": False, "stock_status": "outofstock"},
            )
        self.assertIn("写入结果未确认", error)
        self.assertEqual(writes.call_count, 1)
        self.assertTrue(trace["unconfirmed_write"])
        self.assertTrue(final["manage_stock"])

    def test_failed_preflight_never_sends_a_write(self):
        req = FakeWooRequests()
        with patch.object(req, "get", side_effect=req.RequestException("read timed out")):
            final, error, trace = service.wc_product_update_verified(
                req, "https://shop.test/wp-json/wc/v3/products/101", ("ck", "cs"),
                {"manage_stock": False, "stock_status": "outofstock"},
            )
        self.assertIsNone(final)
        self.assertIn("未提交修改", error)
        self.assertEqual(req.put_urls, [])

    def test_master_and_child_share_one_deadline(self):
        now = [0]
        master, child = FakeWooRequests(), FakeWooRequests()
        child.state["parent_id"] = 77
        def timed(method):
            def call(*args, **kwargs):
                now[0] += sum(kwargs["timeout"])
                return method(*args, **kwargs)
            return call
        for req in (master, child):
            req.get, req.put = timed(req.get), timed(req.put)
        payload = {"manage_stock": False, "stock_status": "outofstock"}
        site = {"url": "https://child.test", "consumer_key": "ck",
                "consumer_secret": "cs", "product_master_id": 2}
        with patch.object(service.time, "monotonic", side_effect=lambda: now[0]):
            deadline = service.new_product_operation_deadline()
            final, error, _ = service.wc_product_update_verified(
                master, "https://shop.test/wp-json/wc/v3/products/101", ("ck", "cs"),
                payload, deadline=deadline,
            )
            self.assertIsNone(error)
            result = service.verify_product_child_sync(child, site, final, payload, deadline=deadline)
        self.assertEqual(result["status"], "error")
        self.assertIn("时间已用完", result["detail"])
        self.assertLessEqual(now[0], 90)
        self.assertEqual(child.put_urls, [])

if __name__ == "__main__":
    unittest.main()
