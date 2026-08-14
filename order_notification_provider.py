"""Official WeCom image provider plus non-network test/manual providers."""

from __future__ import annotations

import base64
import hashlib
import os
import re
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import requests
from PIL import Image


MAX_IMAGE_BYTES = 2 * 1024 * 1024


class ProviderError(RuntimeError):
    def __init__(self, message: str, *, code: str, retryable: bool = False, unknown_outcome: bool = False, http_status: int | None = None):
        super().__init__(message)
        self.code = code
        self.retryable = retryable
        self.unknown_outcome = unknown_outcome
        self.http_status = http_status


def resolve_secret_ref(secret_ref: str | None) -> str:
    """Resolve only environment-backed references; never accept inline URLs."""
    value = str(secret_ref or "")
    if not value.startswith("env:"):
        raise ProviderError("目标未配置安全的环境变量引用", code="secret_ref_invalid")
    name = value[4:]
    if not re.fullmatch(r"[A-Z][A-Z0-9_]{2,100}", name):
        raise ProviderError("密钥引用名称无效", code="secret_ref_invalid")
    secret = os.environ.get(name, "").strip()
    if not secret:
        raise ProviderError("密钥引用未注入", code="secret_missing")
    return secret


def validate_wecom_webhook(url: str) -> None:
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    if (
        parsed.scheme != "https"
        or parsed.hostname != "qyapi.weixin.qq.com"
        or parsed.path != "/cgi-bin/webhook/send"
        or parsed.port not in (None, 443)
        or parsed.username is not None
        or parsed.password is not None
        or bool(parsed.fragment)
        or set(query) != {"key"}
        or len(query.get("key", [])) != 1
        or not query["key"][0].strip()
    ):
        raise ProviderError("企业微信 Webhook 主机或路径不合法", code="webhook_invalid")


def _image_payload(path: str) -> tuple[dict, int]:
    image_path = Path(path)
    raw = image_path.read_bytes()
    if not raw or len(raw) >= MAX_IMAGE_BYTES:
        raise ProviderError("图片为空或超过 2 MB", code="image_size_invalid")
    try:
        with Image.open(image_path) as image:
            if image.format not in {"PNG", "JPEG"}:
                raise ProviderError("图片格式必须为 PNG/JPG", code="image_format_invalid")
            image.verify()
    except ProviderError:
        raise
    except Exception as exc:
        raise ProviderError("图片不可解码", code="image_decode_failed") from exc
    return {
        "msgtype": "image",
        "image": {
            "base64": base64.b64encode(raw).decode("ascii"),
            "md5": hashlib.md5(raw).hexdigest(),  # WeCom contract requires MD5.
        },
    }, len(raw)


class FakeProvider:
    """Test-only provider. It never performs network I/O."""

    channel_type = "FAKE"

    def send_images(self, image_paths: list[str], target: dict) -> dict:
        byte_count = 0
        for path in image_paths:
            _, size = _image_payload(path)
            byte_count += size
        return {"accepted": True, "provider": "fake", "images": len(image_paths), "bytes": byte_count}


class ManualWechatProvider:
    """Ordinary WeChat fallback: render/download only, never automate WeChat."""

    channel_type = "MANUAL_WECHAT"

    def send_images(self, image_paths: list[str], target: dict) -> dict:
        for path in image_paths:
            _image_payload(path)
        return {"accepted": False, "manual_ready": True, "images": len(image_paths)}


class WeComBotProvider:
    channel_type = "WECOM_BOT"

    def __init__(self, session: requests.Session | None = None):
        self.session = session or requests.Session()

    def _send_payload(self, webhook: str, payload: dict) -> dict:
        try:
            response = self.session.post(
                webhook,
                json=payload,
                timeout=(5, 15),
                allow_redirects=False,
                headers={"User-Agent": "woo-analysis-order-notification/1.0"},
            )
        except requests.ConnectTimeout as exc:
            raise ProviderError("连接企业微信超时", code="connect_timeout", retryable=True) from exc
        except requests.ReadTimeout as exc:
            raise ProviderError(
                "企业微信响应超时，发送结果未知",
                code="delivery_unknown",
                unknown_outcome=True,
            ) from exc
        except requests.ConnectionError as exc:
            raise ProviderError("连接企业微信失败", code="connection_error", retryable=True) from exc

        if response.status_code == 429 or response.status_code >= 500:
            raise ProviderError(
                "企业微信暂时不可用",
                code="provider_transient",
                retryable=True,
                http_status=response.status_code,
            )
        if response.status_code != 200:
            raise ProviderError(
                "企业微信拒绝请求",
                code="provider_http_error",
                http_status=response.status_code,
            )
        try:
            body = response.json()
        except ValueError as exc:
            raise ProviderError("企业微信返回非 JSON", code="provider_bad_response", http_status=200) from exc
        errcode = int(body.get("errcode", -999))
        if errcode != 0:
            if errcode == -1:
                raise ProviderError("企业微信系统繁忙", code="provider_busy", retryable=True, http_status=200)
            raise ProviderError(
                "企业微信业务错误",
                code=f"wecom_{errcode}",
                http_status=200,
            )
        return body

    def send_images(self, image_paths: list[str], target: dict) -> dict:
        webhook = resolve_secret_ref(target.get("secret_ref"))
        validate_wecom_webhook(webhook)
        accepted = 0
        for path in image_paths:
            payload, _ = _image_payload(path)
            self._send_payload(webhook, payload)
            accepted += 1
        return {"accepted": True, "provider": "wecom", "images": accepted}

    def send_text(self, content: str, target: dict) -> dict:
        """Send a privacy-minimized operational alert through the same official bot."""
        content = str(content or "").strip()
        if not content or len(content.encode("utf-8")) > 2048:
            raise ProviderError("企业微信文本为空或超过 2048 字节", code="text_size_invalid")
        webhook = resolve_secret_ref(target.get("secret_ref"))
        validate_wecom_webhook(webhook)
        self._send_payload(
            webhook,
            {"msgtype": "text", "text": {"content": content}},
        )
        return {"accepted": True, "provider": "wecom", "messages": 1}


def provider_for(channel_type: str, *, session: requests.Session | None = None):
    if channel_type == "FAKE":
        return FakeProvider()
    if channel_type == "MANUAL_WECHAT":
        return ManualWechatProvider()
    if channel_type == "WECOM_BOT":
        return WeComBotProvider(session=session)
    raise ProviderError("未知通知通道", code="channel_invalid")
