import json
import hashlib
import unittest

import requests

from poland_wms import (
    PolandWmsClient,
    PolandWmsError,
    normalize_poland_tracking_status,
)


class FakeResponse:
    def __init__(self, body, *, status=200, text=None):
        self.body = body
        self.status_code = status
        self.text = text if text is not None else json.dumps(body)

    def json(self):
        if isinstance(self.body, Exception):
            raise self.body
        return self.body


class FakeSession:
    def __init__(self, responses=None, exc=None):
        self.responses = list(responses or [])
        self.exc = exc
        self.calls = []

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        if self.exc:
            raise self.exc
        return self.responses.pop(0)


class PolandWmsAdapterTests(unittest.TestCase):
    def client(self, session, **kwargs):
        return PolandWmsClient(
            order_base_url="http://order.example.test:8082",
            print_base_url="http://print.example.test:8089",
            username="partner-user",
            password="partner-password",
            cancel_auth="cancel-header-secret",
            allow_insecure_http=True,
            session=session,
            **kwargs,
        )

    def test_auth_accepts_documented_single_quoted_response_and_redacts_credentials(self):
        session = FakeSession(
            [
                FakeResponse(
                    ValueError("not strict json"),
                    text="{'customer_id':'6581','customer_userid':'6901','ack':'true'}",
                )
            ]
        )

        result = self.client(session).authenticate()

        self.assertEqual("6581", result.data["customer_id"])
        self.assertEqual("6901", result.data["customer_userid"])
        self.assertNotIn("partner-password", repr(result.request_redacted))
        self.assertNotIn("partner-user", repr(result.request_redacted))
        self.assertEqual("[REDACTED]", result.request_redacted["form"]["username"])
        self.assertEqual("[REDACTED]", result.request_redacted["form"]["password"])

    def test_create_order_uses_form_encoded_param_and_timeout_is_unknown(self):
        session = FakeSession(
            [
                FakeResponse(
                    {
                        "ack": "true",
                        "order_id": "35864",
                        "tracking_number": "RF000000204DEL",
                    }
                )
            ]
        )
        payload = {
            "order_customerinvoicecode": "OMS100",
            "customer_id": "6581",
            "customer_userid": "6901",
            "consignee_name": "Sensitive Name",
        }

        result = self.client(session).create_order(payload)

        sent = json.loads(session.calls[0][2]["data"]["param"])
        self.assertEqual("OMS100", sent["order_customerinvoicecode"])
        self.assertEqual("35864", result.data["order_id"])
        self.assertNotIn("Sensitive Name", repr(result.request_redacted))
        self.assertNotIn("6581", repr(result.request_redacted))

        timeout_client = self.client(
            FakeSession(exc=requests.exceptions.Timeout())
        )
        with self.assertRaises(PolandWmsError) as ctx:
            timeout_client.create_order(payload)
        self.assertTrue(ctx.exception.retryable)
        self.assertTrue(ctx.exception.unknown_outcome)

    def test_plain_http_is_fail_closed_and_label_url_is_deterministic(self):
        client = PolandWmsClient(
            order_base_url="http://order.example.test:8082",
            print_base_url="http://print.example.test:8089",
            username="u",
            password="p",
            allow_insecure_http=False,
            session=FakeSession(),
        )
        with self.assertRaises(PolandWmsError) as ctx:
            client.authenticate()
        self.assertEqual("insecure_http_blocked", ctx.exception.code)

        url = self.client(FakeSession()).label_url(
            "35864", print_type="lab10_10", print_goods=True
        )
        self.assertEqual(
            "http://print.example.test:8089/order/FastRpt/PDF_NEW.aspx?"
            "PrintType=lab10_10&order_id=35864&PrintGoods=1",
            url,
        )

    def test_tracking_status_normalization(self):
        self.assertEqual("delivered", normalize_poland_tracking_status("已签收"))
        self.assertEqual("in_transit", normalize_poland_tracking_status("运输途中"))
        self.assertEqual("undelivered", normalize_poland_tracking_status("拒收"))
        self.assertEqual("returning", normalize_poland_tracking_status("退回中"))

    def test_cancel_uses_separate_auth_and_documented_md5_signature(self):
        session = FakeSession(
            [FakeResponse({"ack": True, "msg_code": "200", "message": "ok"})]
        )
        client = self.client(session)
        request_time = "2020-08-06 14:28:00"

        result = client.cancel_order(
            order_id="278422",
            customer_id="16421",
            reason="customer cancelled",
            request_time=request_time,
        )

        call = session.calls[0]
        expected = hashlib.md5(
            "278422164212020-08-06 14:28:00order.cancel".encode("utf-8")
        ).hexdigest().upper()
        self.assertEqual(expected, call[2]["json"]["content"]["sign"])
        self.assertEqual("cancel-header-secret", call[2]["headers"]["auth"])
        self.assertNotIn("cancel-header-secret", repr(result.request_redacted))
        self.assertNotIn("16421", repr(result.request_redacted))


if __name__ == "__main__":
    unittest.main()
