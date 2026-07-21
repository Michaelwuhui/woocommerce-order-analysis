import unittest

import requests

from hungary_wms import WmsClient, WmsError, normalize_wms_tracking_status


class FakeResponse:
    status_code = 200
    text = '{"code":1}'

    def json(self):
        return {"code": 1, "message": "ok", "datas": [{"sku": "A"}]}


class FakeSession:
    def __init__(self, exc=None):
        self.exc = exc
        self.calls = []

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        if self.exc:
            raise self.exc
        return FakeResponse()


class WmsAdapterTests(unittest.TestCase):
    def test_inventory_auth_header_and_audit_redaction(self):
        session = FakeSession()
        client = WmsClient(base_url="http://example.test/apicall", salt="super-secret", allow_insecure_http=True, session=session)
        result = client.inventory("HU01")
        self.assertEqual("super-secret", session.calls[0][2]["headers"]["salt"])
        self.assertNotIn("super-secret", repr(result.request_redacted))
        self.assertEqual({"storehouseCode": "HU01"}, result.request_redacted)

    def test_create_timeout_has_unknown_outcome(self):
        client = WmsClient(base_url="https://example.test", salt="secret", session=FakeSession(requests.exceptions.Timeout()))
        with self.assertRaises(WmsError) as ctx:
            client.create_invoices([{"invoiceCode": "ONE"}])
        self.assertTrue(ctx.exception.retryable)
        self.assertTrue(ctx.exception.unknown_outcome)

    def test_plain_http_is_explicitly_blocked(self):
        client = WmsClient(base_url="http://example.test", salt="secret", allow_insecure_http=False, session=FakeSession())
        with self.assertRaises(WmsError) as ctx:
            client.inventory()
        self.assertEqual("insecure_http_blocked", ctx.exception.code)

    def test_cancel_is_never_sent_to_an_undocumented_endpoint(self):
        session = FakeSession()
        client = WmsClient(
            base_url="https://example.test", salt="secret", session=session
        )
        with self.assertRaises(WmsError) as ctx:
            client.cancel_invoice("INV-1")
        self.assertEqual("cancel_not_supported", ctx.exception.code)
        self.assertFalse(ctx.exception.retryable)
        self.assertEqual([], session.calls)

    def test_tracking_normalization(self):
        self.assertEqual("delivered", normalize_wms_tracking_status("delivered"))
        self.assertEqual("in_transit", normalize_wms_tracking_status("transit"))
        self.assertEqual("returned", normalize_wms_tracking_status("returned"))


if __name__ == "__main__":
    unittest.main()
