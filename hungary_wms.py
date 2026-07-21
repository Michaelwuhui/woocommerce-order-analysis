"""Adapter for the supplier Hungary WMS API.

The public contract uses a static ``salt`` header and currently documents an
HTTP endpoint.  Insecure HTTP is rejected unless the deployment explicitly
sets ``WMS_ALLOW_INSECURE_HTTP=1``; this makes the security exception visible
and easy to remove when the supplier enables valid HTTPS.
"""

from __future__ import annotations

import os
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import quote

import requests

from fulfillment_common import redact


@dataclass
class WmsResult:
    data: Any
    raw: dict
    http_status: int
    business_code: int | str | None
    message: str
    duration_ms: int
    method: str
    endpoint: str
    request_redacted: Any


class WmsError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        code: str = "wms_error",
        retryable: bool = False,
        unknown_outcome: bool = False,
        http_status: int | None = None,
        business_code: int | str | None = None,
        duration_ms: int | None = None,
        response: Any = None,
    ):
        super().__init__(message)
        self.code = code
        self.retryable = retryable
        self.unknown_outcome = unknown_outcome
        self.http_status = http_status
        self.business_code = business_code
        self.duration_ms = duration_ms
        self.response = response


class WmsClient:
    def __init__(
        self,
        *,
        base_url: str | None = None,
        salt: str | None = None,
        timeout: tuple[int, int] = (8, 45),
        allow_insecure_http: bool | None = None,
        session=None,
    ):
        self.base_url = (base_url or os.environ.get("WMS_BASE_URL") or "http://cod.kuasuda.com/apicall").rstrip("/")
        self.salt = salt if salt is not None else os.environ.get("WMS_SALT", "")
        self.timeout = timeout
        if allow_insecure_http is None:
            allow_insecure_http = str(os.environ.get("WMS_ALLOW_INSECURE_HTTP", "0")).lower() in {
                "1", "true", "yes", "on"
            }
        self.allow_insecure_http = bool(allow_insecure_http)
        self.session = session or requests.Session()

    def _validate_config(self):
        if not self.salt:
            raise WmsError("WMS_SALT 未配置", code="wms_salt_missing")
        if self.base_url.lower().startswith("http://") and not self.allow_insecure_http:
            raise WmsError(
                "WMS 端点是明文 HTTP；必须先提供 HTTPS/专线，或显式设置 WMS_ALLOW_INSECURE_HTTP=1",
                code="insecure_http_blocked",
            )

    def _request(self, method: str, path: str, *, payload: Any = None, unknown_on_timeout: bool = False) -> WmsResult:
        self._validate_config()
        url = f"{self.base_url}/{path.lstrip('/')}"
        headers = {
            "salt": self.salt,
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "woo-analysis-fulfillment/2.0",
        }
        started = time.monotonic()
        try:
            response = self.session.request(
                method.upper(),
                url,
                json=payload if payload is not None else None,
                headers=headers,
                timeout=self.timeout,
            )
        except requests.exceptions.Timeout as exc:
            elapsed = int((time.monotonic() - started) * 1000)
            raise WmsError(
                "WMS 请求超时",
                code="timeout",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                duration_ms=elapsed,
            ) from exc
        except requests.exceptions.ConnectionError as exc:
            elapsed = int((time.monotonic() - started) * 1000)
            raise WmsError(
                "WMS 连接失败",
                code="connection",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                duration_ms=elapsed,
            ) from exc
        except requests.exceptions.RequestException as exc:
            elapsed = int((time.monotonic() - started) * 1000)
            raise WmsError(
                f"WMS 请求异常: {exc}",
                code="request_error",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                duration_ms=elapsed,
            ) from exc

        elapsed = int((time.monotonic() - started) * 1000)
        try:
            body = response.json()
        except ValueError as exc:
            raise WmsError(
                "WMS 返回的不是 JSON",
                code="invalid_json",
                retryable=response.status_code >= 500,
                unknown_outcome=unknown_on_timeout and response.status_code >= 500,
                http_status=response.status_code,
                duration_ms=elapsed,
                response=(response.text or "")[:500],
            ) from exc

        if response.status_code == 429 or response.status_code >= 500:
            raise WmsError(
                f"WMS HTTP {response.status_code}",
                code=f"http_{response.status_code}",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                http_status=response.status_code,
                duration_ms=elapsed,
                response=redact(body),
            )
        if response.status_code >= 400:
            raise WmsError(
                f"WMS HTTP {response.status_code}",
                code=f"http_{response.status_code}",
                retryable=False,
                http_status=response.status_code,
                duration_ms=elapsed,
                response=redact(body),
            )

        business_code = body.get("code") if isinstance(body, dict) else None
        message = str(body.get("message") or "") if isinstance(body, dict) else ""
        if str(business_code) != "1":
            retryable = str(business_code) not in {"-1", "-2", "-3", "0"}
            raise WmsError(
                message or f"WMS 业务错误 {business_code}",
                code=f"business_{business_code}",
                retryable=retryable,
                http_status=response.status_code,
                business_code=business_code,
                duration_ms=elapsed,
                response=redact(body),
            )
        return WmsResult(
            data=body.get("datas"),
            raw=body,
            http_status=response.status_code,
            business_code=business_code,
            message=message,
            duration_ms=elapsed,
            method=method.upper(),
            endpoint=path,
            request_redacted=redact(payload),
        )

    def create_invoices(self, invoices: list[dict]) -> WmsResult:
        return self._request("POST", "/invoice/create", payload=invoices, unknown_on_timeout=True)

    def invoice_status(self, invoice_codes: list[str]) -> WmsResult:
        return self._request(
            "POST", "/api/last/invoice/invoiceStatus", payload={"invoiceCodes": invoice_codes}
        )

    def cancel_invoice(self, invoice_code: str) -> WmsResult:
        # Supplier confirmation (2026-07-21): outbound cancellation/intercept
        # is not exposed through the API and must be handled in the logistics
        # group by their operations team.  Fail closed so no guessed endpoint
        # can be called accidentally.
        raise WmsError(
            "供应商 WMS 不支持 API 取消；请在物流群联系运营人工拦截",
            code="cancel_not_supported",
            retryable=False,
        )

    def labels(self, invoice_codes: list[str]) -> WmsResult:
        return self._request("PUT", "/invoice/getLabelUrl", payload={"invoiceCodes": invoice_codes})

    def inventory(self, storehouse_code: str = "HU01") -> WmsResult:
        return self._request(
            "POST", "/api/last/inventory/inventoryList", payload={"storehouseCode": storehouse_code}
        )

    def track(self, tracking_number: str) -> WmsResult:
        return self._request("GET", f"/api/logistics/track/{quote(str(tracking_number), safe='')}")

    def tracks(self, tracking_numbers: list[str]) -> WmsResult:
        if len(tracking_numbers) > 100:
            raise ValueError("WMS 批量物流查询单次最多 100 个运单号")
        # The public page names the list inconsistently.  ``expressCodes`` is
        # isolated here so a supplier correction only changes the adapter.
        return self._request(
            "POST", "/api/logistics/tracks", payload={"expressCodes": tracking_numbers}
        )


def normalize_wms_fulfillment_status(raw: str | None) -> str | None:
    value = (raw or "").strip().lower()
    if not value:
        return None
    if any(token in value for token in ("缺货", "lack", "shortage")):
        return "stock_shortage"
    if any(token in value for token in ("取消", "作废", "cancel")):
        return "cancelled"
    if any(token in value for token in ("已发货", "出库", "shipped", "outbound")):
        return "shipped"
    if any(token in value for token in ("打包", "packed")):
        return "packed"
    if any(token in value for token in ("拣货", "picking")):
        return "picking"
    if any(token in value for token in ("已接单", "已创建", "待处理", "accepted", "created")):
        return "accepted"
    return None


def normalize_wms_tracking_status(raw: str | None) -> str:
    value = (raw or "").strip().lower().replace("-", "_")
    mapping = {
        "create": "shipped",
        "notfound": "not_found",
        "not_found": "not_found",
        "transit": "in_transit",
        "in_transit": "in_transit",
        "pickup": "pickup_ready",
        "delivered": "delivered",
        "undelivered": "undelivered",
        "exception": "exception",
        "expired": "expired",
        "cancel": "cancelled",
        "cancelled": "cancelled",
        "returned": "returned",
        "returning": "returning",
        "other": "exception",
    }
    return mapping.get(value, "exception")
