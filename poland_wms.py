"""Adapter for the new Poland warehouse's HuaLei/SZ56T API.

The supplier exposes two plain-HTTP services:

* order/auth/tracking: ``http://175.178.192.240:8082``
* label rendering: ``http://175.178.192.240:8089``

The API is form-encoded and returns a mixture of JSON and Python-style
single-quoted dictionaries.  This adapter contains that legacy behaviour so
the fulfillment domain never needs to depend on it directly.
"""

from __future__ import annotations

import ast
import hashlib
import json
import os
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Any
from urllib.parse import unquote, urlencode

import requests

from fulfillment_common import json_dump, redact


DEFAULT_ORDER_BASE_URL = "http://175.178.192.240:8082"
DEFAULT_PRINT_BASE_URL = "http://175.178.192.240:8089"


@dataclass
class PolandWmsResult:
    data: Any
    raw: Any
    http_status: int
    business_code: str | int | None
    message: str
    duration_ms: int
    method: str
    endpoint: str
    request_redacted: Any


class PolandWmsError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        code: str = "poland_wms_error",
        retryable: bool = False,
        unknown_outcome: bool = False,
        http_status: int | None = None,
        business_code: str | int | None = None,
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


def _truthy(value: Any) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "y", "ok", "success"}


def _decode_legacy_body(response) -> Any:
    """Decode strict JSON first, then the documented single-quoted response."""

    try:
        return response.json()
    except ValueError:
        text = (response.text or "").strip().lstrip("\ufeff")
        if not text:
            return {}
        try:
            return json.loads(text)
        except ValueError:
            try:
                value = ast.literal_eval(text)
            except (SyntaxError, ValueError) as exc:
                raise PolandWmsError(
                    "新波兰仓接口返回的不是可识别 JSON",
                    code="invalid_json",
                    retryable=response.status_code >= 500,
                    http_status=response.status_code,
                    response=text[:500],
                ) from exc
            if isinstance(value, (dict, list)):
                return value
            raise PolandWmsError(
                "新波兰仓接口返回格式不受支持",
                code="invalid_response_type",
                http_status=response.status_code,
                response=text[:500],
            )


def _message(body: Any) -> str:
    if isinstance(body, dict):
        value = body.get("message") or body.get("msg") or body.get("error")
        return unquote(str(value or ""))
    return ""


def _audit_request(form: dict | None, params: dict | None) -> dict:
    form_copy = dict(form or {})
    encoded_payload = form_copy.get("param")
    if isinstance(encoded_payload, str):
        try:
            form_copy["param"] = json.loads(encoded_payload)
        except ValueError:
            form_copy["param"] = "[UNPARSEABLE_FORM_PAYLOAD]"
    return redact({"form": form_copy, "params": dict(params or {})})


class PolandWmsClient:
    def __init__(
        self,
        *,
        order_base_url: str | None = None,
        print_base_url: str | None = None,
        username: str | None = None,
        password: str | None = None,
        cancel_auth: str | None = None,
        timeout: tuple[int, int] = (8, 45),
        allow_insecure_http: bool | None = None,
        session=None,
    ):
        self.order_base_url = (
            order_base_url
            or os.environ.get("SZ56T_ORDER_BASE_URL")
            or DEFAULT_ORDER_BASE_URL
        ).rstrip("/")
        self.print_base_url = (
            print_base_url
            or os.environ.get("SZ56T_PRINT_BASE_URL")
            or DEFAULT_PRINT_BASE_URL
        ).rstrip("/")
        self.username = username if username is not None else os.environ.get("SZ56T_USERNAME", "")
        self.password = password if password is not None else os.environ.get("SZ56T_PASSWORD", "")
        self.cancel_auth = (
            cancel_auth
            if cancel_auth is not None
            else os.environ.get("SZ56T_CANCEL_AUTH", "")
        )
        self.timeout = timeout
        if allow_insecure_http is None:
            allow_insecure_http = _truthy(os.environ.get("SZ56T_ALLOW_INSECURE_HTTP", "0"))
        self.allow_insecure_http = bool(allow_insecure_http)
        self.session = session or requests.Session()

    def _validate_transport(self):
        urls = (self.order_base_url, self.print_base_url)
        if any(url.lower().startswith("http://") for url in urls) and not self.allow_insecure_http:
            raise PolandWmsError(
                "新波兰仓只提供明文 HTTP；需供应商提供 HTTPS/专线，或显式设置 "
                "SZ56T_ALLOW_INSECURE_HTTP=1",
                code="insecure_http_blocked",
            )

    def _request(
        self,
        method: str,
        path: str,
        *,
        form: dict | None = None,
        params: dict | None = None,
        unknown_on_timeout: bool = False,
    ) -> PolandWmsResult:
        self._validate_transport()
        url = f"{self.order_base_url}/{path.lstrip('/')}"
        started = time.monotonic()
        try:
            response = self.session.request(
                method.upper(),
                url,
                data=form,
                params=params,
                headers={
                    "Accept": "application/json,text/plain,*/*",
                    "Accept-Language": "zh-cn",
                    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
                    "User-Agent": "woo-analysis-fulfillment/2.0",
                },
                timeout=self.timeout,
            )
        except requests.exceptions.Timeout as exc:
            raise PolandWmsError(
                "新波兰仓接口请求超时",
                code="timeout",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                duration_ms=int((time.monotonic() - started) * 1000),
            ) from exc
        except requests.exceptions.ConnectionError as exc:
            raise PolandWmsError(
                "新波兰仓接口连接失败",
                code="connection",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                duration_ms=int((time.monotonic() - started) * 1000),
            ) from exc
        except requests.exceptions.RequestException as exc:
            raise PolandWmsError(
                f"新波兰仓接口请求异常: {exc}",
                code="request_error",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                duration_ms=int((time.monotonic() - started) * 1000),
            ) from exc

        elapsed = int((time.monotonic() - started) * 1000)
        if response.status_code == 429 or response.status_code >= 500:
            body = None
            try:
                body = _decode_legacy_body(response)
            except PolandWmsError:
                body = (response.text or "")[:500]
            raise PolandWmsError(
                f"新波兰仓接口 HTTP {response.status_code}",
                code=f"http_{response.status_code}",
                retryable=True,
                unknown_outcome=unknown_on_timeout,
                http_status=response.status_code,
                duration_ms=elapsed,
                response=redact(body),
            )
        if response.status_code >= 400:
            raise PolandWmsError(
                f"新波兰仓接口 HTTP {response.status_code}",
                code=f"http_{response.status_code}",
                http_status=response.status_code,
                duration_ms=elapsed,
                response=(response.text or "")[:500],
            )

        try:
            body = _decode_legacy_body(response)
        except PolandWmsError as exc:
            exc.duration_ms = elapsed
            exc.unknown_outcome = bool(unknown_on_timeout)
            raise
        business_code = None
        if isinstance(body, dict):
            business_code = body.get("ack")
            if business_code is None:
                business_code = body.get("resultCode")
            if business_code is None:
                business_code = body.get("status")
        return PolandWmsResult(
            data=body,
            raw=body,
            http_status=response.status_code,
            business_code=business_code,
            message=_message(body),
            duration_ms=elapsed,
            method=method.upper(),
            endpoint=path,
            request_redacted=_audit_request(form, params),
        )

    def authenticate(self) -> PolandWmsResult:
        if not self.username or not self.password:
            raise PolandWmsError(
                "SZ56T_USERNAME / SZ56T_PASSWORD 未配置",
                code="credentials_missing",
            )
        result = self._request(
            "POST",
            "/selectAuth.htm",
            form={"username": self.username, "password": self.password},
        )
        body = result.data if isinstance(result.data, dict) else {}
        if not _truthy(body.get("ack")):
            raise PolandWmsError(
                result.message or "新波兰仓账号认证失败",
                code="authentication_failed",
                business_code=body.get("ack"),
                http_status=result.http_status,
                duration_ms=result.duration_ms,
                response=redact(body),
            )
        if not body.get("customer_id") or not body.get("customer_userid"):
            raise PolandWmsError(
                "认证成功但未返回 customer_id/customer_userid",
                code="authentication_ids_missing",
                http_status=result.http_status,
                duration_ms=result.duration_ms,
                response=redact(body),
            )
        return result

    def product_list(self) -> PolandWmsResult:
        return self._request("GET", "/getProductList.htm")

    def create_order(self, payload: dict) -> PolandWmsResult:
        result = self._request(
            "POST",
            "/createOrderApi.htm",
            form={"param": json_dump(payload)},
            unknown_on_timeout=True,
        )
        body = result.data if isinstance(result.data, dict) else {}
        if (
            ("ack" in body and not _truthy(body.get("ack")))
            or ("resultCode" in body and not _truthy(body.get("resultCode")))
            or not body.get("order_id")
        ):
            raise PolandWmsError(
                result.message or "新波兰仓创建物流订单失败",
                code="create_order_rejected",
                business_code=body.get("ack") or body.get("resultCode"),
                http_status=result.http_status,
                duration_ms=result.duration_ms,
                response=redact(body),
            )
        return result

    def tracking_number(
        self,
        *,
        document_code: str | None = None,
        order_id: str | None = None,
    ) -> PolandWmsResult:
        if not document_code and not order_id:
            raise ValueError("document_code 或 order_id 至少提供一个")
        params = {"documentCode": document_code} if document_code else {"order_id": order_id}
        return self._request("GET", "/getOrderTrackingNumber.htm", params=params)

    def track(self, tracking_number: str) -> PolandWmsResult:
        return self._request(
            "GET",
            "/selectTrack.htm",
            params={"documentCode": str(tracking_number)},
        )

    def mark_shipped(self, customer_id: str, document_code: str) -> PolandWmsResult:
        return self._request(
            "POST",
            "/postOrderApi.htm",
            form={
                "customer_id": str(customer_id),
                "order_customerinvoicecode": str(document_code),
            },
            unknown_on_timeout=True,
        )

    def cancel_order(
        self,
        *,
        order_id: str,
        customer_id: str,
        reason: str,
        request_time: str | None = None,
    ) -> PolandWmsResult:
        """Cancel an order through the separately documented Open API."""

        self._validate_transport()
        if not self.cancel_auth:
            raise PolandWmsError(
                "SZ56T_CANCEL_AUTH 未配置，不能调用新波兰仓取消接口",
                code="cancel_auth_missing",
            )
        method_name = "order.cancel"
        request_time = request_time or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        signature_source = (
            f"{order_id}{customer_id}{request_time}{method_name}"
        )
        signature = hashlib.md5(
            signature_source.encode("utf-8")
        ).hexdigest().upper()
        payload = {
            "method": method_name,
            "req_time": request_time,
            "content": {
                "order_id": str(order_id),
                "customer_id": str(customer_id),
                "reason": str(reason),
                "sign": signature,
            },
        }
        url = f"{self.order_base_url}/logistics/api"
        started = time.monotonic()
        try:
            response = self.session.request(
                "POST",
                url,
                json=payload,
                headers={
                    "auth": self.cancel_auth,
                    "Accept": "application/json",
                    "Content-Type": "application/json;charset=UTF-8",
                    "User-Agent": "woo-analysis-fulfillment/2.0",
                },
                timeout=self.timeout,
            )
        except requests.exceptions.Timeout as exc:
            raise PolandWmsError(
                "新波兰仓取消请求超时",
                code="cancel_timeout",
                retryable=False,
                unknown_outcome=True,
                duration_ms=int((time.monotonic() - started) * 1000),
            ) from exc
        except requests.exceptions.ConnectionError as exc:
            raise PolandWmsError(
                "新波兰仓取消接口连接失败",
                code="cancel_connection",
                retryable=False,
                unknown_outcome=True,
                duration_ms=int((time.monotonic() - started) * 1000),
            ) from exc
        except requests.exceptions.RequestException as exc:
            raise PolandWmsError(
                f"新波兰仓取消接口请求异常: {exc}",
                code="cancel_request_error",
                retryable=False,
                unknown_outcome=True,
                duration_ms=int((time.monotonic() - started) * 1000),
            ) from exc

        elapsed = int((time.monotonic() - started) * 1000)
        try:
            body = _decode_legacy_body(response)
        except PolandWmsError as exc:
            exc.duration_ms = elapsed
            exc.unknown_outcome = response.status_code >= 500
            raise
        if response.status_code >= 500:
            raise PolandWmsError(
                f"新波兰仓取消接口 HTTP {response.status_code}",
                code=f"cancel_http_{response.status_code}",
                unknown_outcome=True,
                http_status=response.status_code,
                duration_ms=elapsed,
                response=redact(body),
            )
        if response.status_code >= 400:
            raise PolandWmsError(
                f"新波兰仓取消接口 HTTP {response.status_code}",
                code=f"cancel_http_{response.status_code}",
                http_status=response.status_code,
                duration_ms=elapsed,
                response=redact(body),
            )
        ack = body.get("ack") if isinstance(body, dict) else None
        msg_code = body.get("msg_code") if isinstance(body, dict) else None
        if not _truthy(ack) or (msg_code not in (None, "", 200, "200")):
            raise PolandWmsError(
                _message(body) or "新波兰仓取消订单失败",
                code=f"cancel_business_{msg_code or ack}",
                business_code=msg_code or ack,
                http_status=response.status_code,
                duration_ms=elapsed,
                response=redact(body),
            )
        return PolandWmsResult(
            data=body,
            raw=body,
            http_status=response.status_code,
            business_code=msg_code or ack,
            message=_message(body),
            duration_ms=elapsed,
            method="POST",
            endpoint="/logistics/api",
            request_redacted=redact(
                {"headers": {"auth": self.cancel_auth}, "json": payload}
            ),
        )

    def label_url(
        self,
        order_id: str,
        *,
        print_type: str = "lab10_10",
        label_format: str | None = None,
        print_goods: bool = False,
    ) -> str:
        self._validate_transport()
        params = {"PrintType": print_type, "order_id": str(order_id)}
        if label_format:
            params["Format"] = label_format
        if print_goods:
            params["PrintGoods"] = "1"
        return f"{self.print_base_url}/order/FastRpt/PDF_NEW.aspx?{urlencode(params)}"


def normalize_poland_tracking_status(raw: str | None) -> str:
    value = (raw or "").strip().lower().replace("-", "_")
    if any(token in value for token in ("签收", "妥投", "delivered", "success")):
        return "delivered"
    if any(token in value for token in ("退回中", "returning")):
        return "returning"
    if any(token in value for token in ("退回", "退件", "returned")):
        return "returned"
    if any(token in value for token in ("拒收", "派送失败", "未妥投", "undelivered")):
        return "undelivered"
    if any(token in value for token in ("异常", "扣关", "丢失", "exception")):
        return "exception"
    if any(token in value for token in ("待取", "自提", "pickup")):
        return "pickup_ready"
    if any(token in value for token in ("运输", "转运", "派送", "揽收", "in_transit", "transit")):
        return "in_transit"
    if any(token in value for token in ("创建", "预报", "已发货", "shipped")):
        return "shipped"
    if any(token in value for token in ("无轨迹", "not_found", "notfound")):
        return "not_found"
    return "exception"
