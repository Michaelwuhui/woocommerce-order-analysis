"""Render the exact logged WooCommerce admin order email as safe PNG files.

The source is the site's existing read-only ``woo-tracking`` endpoint.  The
browser never receives WooCommerce credentials and is never allowed to make a
network request: trusted same-site images are downloaded separately, validated
and embedded as data URIs before Chromium sees the document.
"""

from __future__ import annotations

import base64
import hashlib
import ipaddress
import io
import json
import os
import re
import shutil
import socket
import tempfile
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urljoin, urlparse

import requests
from PIL import Image

from oid_utils import woo_post_id


EMAIL_TEMPLATE_VERSION = "woo-admin-email-v1"
MAX_EMAIL_HTML_BYTES = 2 * 1024 * 1024
MAX_IMAGE_BYTES = 2 * 1024 * 1024
MAX_TOTAL_IMAGE_BYTES = 8 * 1024 * 1024
MAX_IMAGES = 40
MAX_RENDER_HEIGHT = 25_000
PAGE_HEIGHT = 5_000
MAX_PAGES = 5
MAX_OUTPUT_BYTES = 1_900_000

_ADMIN_NEW_ORDER_PHRASES = (
    "new order",
    "new customer order",
    "you've got a new order",
    "you’ve got a new order",
    "nowe zamówienie",
    "nowe zamowienie",
    "nová objednávka",
    "nova objednavka",
    "új rendelés",
    "uj rendeles",
    "neue bestellung",
    "nouvelle commande",
    "nieuwe bestelling",
    "nuevo pedido",
)
_SUCCESS_VALUES = {"1", "true", "sent", "success", "successful", "delivered", "completed"}
_DANGEROUS_TAGS = {
    "script", "iframe", "frame", "frameset", "object", "embed", "applet",
    "form", "input", "button", "textarea", "select", "option", "video",
    "audio", "source", "track", "canvas", "svg", "math", "base", "link",
}
_URL_ATTRIBUTES = {
    "href", "action", "formaction", "poster", "background", "data", "ping",
    "cite", "longdesc", "usemap", "manifest", "srcdoc", "xlink:href",
}


class EmailRenderError(RuntimeError):
    def __init__(self, message: str, *, code: str, retryable: bool = False):
        super().__init__(message)
        self.code = code
        self.retryable = retryable


def _json(value: Any, default: Any) -> Any:
    if isinstance(value, (dict, list)):
        return value
    try:
        return json.loads(value or "")
    except (TypeError, ValueError):
        return default


def _flatten_addresses(value: Any) -> list[str]:
    """Return lowercase addresses from FluentSMTP's varying recipient shapes."""
    value = _json(value, value)
    values: list[Any]
    if isinstance(value, dict):
        values = list(value.keys()) + list(value.values())
    elif isinstance(value, list):
        values = value
    else:
        values = [value]
    addresses: list[str] = []
    for item in values:
        if isinstance(item, dict):
            addresses.extend(_flatten_addresses(item))
            continue
        for match in re.findall(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", str(item or ""), re.I):
            lowered = match.lower()
            if lowered not in addresses:
                addresses.append(lowered)
    return addresses


def _has_exact_order_reference(value: Any, number: str) -> bool:
    if not number:
        return False
    return bool(re.search(rf"(?<!\d){re.escape(number)}(?!\d)", str(value or "")))


def _is_success(log: dict) -> bool:
    value = log.get("status")
    if value is None:
        value = log.get("success")
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in _SUCCESS_VALUES


def select_admin_new_order_log(logs: list[dict], order_number: str, billing_email: str = "") -> dict:
    """Select one successful admin-new-order template, never a customer receipt.

    The recipient address is not an email-type discriminator. WooCommerce can
    legitimately send its administrator template to an address that also
    appears as the order billing address (common for internal/test orders).
    Positive template/subject classification therefore takes precedence.
    """
    candidates = []
    for log in logs or []:
        if not isinstance(log, dict) or not _is_success(log):
            continue
        subject = str(log.get("subject") or "")
        folded = subject.casefold()
        if not _has_exact_order_reference(subject, order_number):
            continue
        if not any(phrase in folded for phrase in _ADMIN_NEW_ORDER_PHRASES):
            continue
        if log.get("id") in (None, ""):
            continue
        candidates.append(log)
    if not candidates:
        raise EmailRenderError(
            "未找到该订单已成功发送的管理员新订单邮件",
            code="admin_new_order_email_not_found",
            retryable=True,
        )
    def sort_key(item):
        try:
            log_id = int(item.get("id") or 0)
        except (TypeError, ValueError):
            log_id = 0
        return str(item.get("sent_at") or item.get("created_at") or ""), log_id

    candidates.sort(key=sort_key, reverse=True)
    return candidates[0]


def _get_json(response, *, code: str) -> dict:
    if response.status_code != 200:
        retryable = response.status_code >= 500 or response.status_code in {408, 429}
        raise EmailRenderError(
            f"邮件日志接口返回 HTTP {response.status_code}", code=code, retryable=retryable
        )
    try:
        payload = response.json()
    except (TypeError, ValueError) as exc:
        raise EmailRenderError("邮件日志接口没有返回 JSON", code=code, retryable=True) from exc
    if not isinstance(payload, dict) or payload.get("success") is False:
        raise EmailRenderError(
            str(payload.get("error") if isinstance(payload, dict) else "邮件日志响应无效"),
            code=code,
            retryable=False,
        )
    return payload


def fetch_admin_new_order_email(conn, order_id: str, *, session=None) -> dict:
    """Fetch the exact logged admin email using GET-only, site-scoped credentials."""
    order = conn.execute(
        "SELECT id,woo_id,number,source,billing FROM orders WHERE id=?", (order_id,)
    ).fetchone()
    if not order:
        raise EmailRenderError("订单不存在", code="order_not_found")
    order = dict(order)
    site = conn.execute(
        "SELECT url,consumer_key,consumer_secret FROM sites WHERE url=?", (order["source"],)
    ).fetchone()
    if not site or not site["consumer_key"] or not site["consumer_secret"]:
        raise EmailRenderError("站点邮件日志凭据未配置", code="email_log_credentials_missing")
    site = dict(site)
    request_session = session or requests.Session()
    remote_order_id = order.get("woo_id") or woo_post_id(order["id"])
    base = site["url"].rstrip("/") + f"/wp-json/woo-tracking/v1/orders/{remote_order_id}/email-logs"
    headers = {
        "User-Agent": "HongKong-Order-Notification/1.0",
        "Accept": "application/json",
        "X-Woo-Tracking-Key": site["consumer_key"],
        "X-Woo-Tracking-Secret": site["consumer_secret"],
    }
    try:
        listing_response = request_session.get(base, headers=headers, timeout=(5, 20))
    except requests.RequestException as exc:
        raise EmailRenderError("读取邮件日志失败", code="email_log_unavailable", retryable=True) from exc
    listing = _get_json(listing_response, code="email_log_list_failed")
    billing = _json(order.get("billing"), {})
    selected = select_admin_new_order_log(
        listing.get("logs") or [], str(order.get("number") or remote_order_id), billing.get("email", "")
    )
    try:
        detail_response = request_session.get(
            f"{base}/{int(selected['id'])}", headers=headers, timeout=(5, 25)
        )
    except (requests.RequestException, TypeError, ValueError) as exc:
        raise EmailRenderError("读取邮件正文失败", code="email_log_detail_unavailable", retryable=True) from exc
    detail = _get_json(detail_response, code="email_log_detail_failed")
    subject = str(detail.get("subject") or selected.get("subject") or "")
    number = str(order.get("number") or remote_order_id)
    detail_has_status = "status" in detail or "success" in detail
    if not (_is_success(detail) if detail_has_status else _is_success(selected)) or not _has_exact_order_reference(subject, number):
        raise EmailRenderError("邮件正文与订单不匹配", code="email_log_order_mismatch")
    if not any(phrase in subject.casefold() for phrase in _ADMIN_NEW_ORDER_PHRASES):
        raise EmailRenderError("该邮件不是管理员新订单通知", code="email_log_type_mismatch")
    body = detail.get("body")
    if not isinstance(body, str) or not body.strip():
        raise EmailRenderError("邮件日志没有保存正文", code="email_body_missing")
    if len(body.encode("utf-8")) > MAX_EMAIL_HTML_BYTES:
        raise EmailRenderError("邮件正文超过安全上限", code="email_body_too_large")
    return {
        "order_id": order["id"],
        "order_number": number,
        "site_url": site["url"],
        "log_id": int(selected["id"]),
        "plugin": detail.get("plugin") or listing.get("plugin") or selected.get("source") or "unknown",
        "source": detail.get("source") or selected.get("source") or "FluentSMTP",
        "subject": subject,
        "sent_at": detail.get("sent_at") or selected.get("sent_at"),
        "body": body,
    }


def _normal_host(host: str) -> str:
    value = (host or "").rstrip(".").lower()
    return value[4:] if value.startswith("www.") else value


def _public_host(host: str) -> bool:
    try:
        addresses = {item[4][0] for item in socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)}
    except OSError:
        return False
    if not addresses:
        return False
    for raw in addresses:
        try:
            if not ipaddress.ip_address(raw).is_global:
                return False
        except ValueError:
            return False
    return True


def _clean_css(value: str) -> str:
    css = str(value or "")
    css = re.sub(r"@import\s+[^;]+;?", "", css, flags=re.I)
    css = re.sub(r"url\s*\([^)]*\)", "none", css, flags=re.I)
    css = re.sub(r"expression\s*\([^)]*\)", "", css, flags=re.I)
    css = re.sub(r"(?:javascript|vbscript)\s*:", "", css, flags=re.I)
    css = re.sub(r"(?:behavior|-moz-binding)\s*:[^;]+;?", "", css, flags=re.I)
    return css


def _read_response_bytes(response, limit: int) -> bytes:
    chunks = []
    size = 0
    iterator = response.iter_content(64 * 1024) if hasattr(response, "iter_content") else [response.content]
    for chunk in iterator:
        if not chunk:
            continue
        size += len(chunk)
        if size > limit:
            raise EmailRenderError("邮件图片超过安全上限", code="email_image_too_large")
        chunks.append(chunk)
    return b"".join(chunks)


def _validated_data_image(src: str) -> tuple[bytes, str] | None:
    match = re.fullmatch(r"data:(image/(?:png|jpeg|gif|webp));base64,([A-Za-z0-9+/=\s]+)", src, re.I)
    if not match:
        return None
    try:
        raw = base64.b64decode(match.group(2), validate=True)
    except (ValueError, TypeError):
        return None
    if len(raw) > MAX_IMAGE_BYTES:
        return None
    try:
        image = Image.open(io.BytesIO(raw))
        image.verify()
    except Exception:
        return None
    return raw, match.group(1).lower()


def sanitize_email_html(
    html: str,
    site_url: str,
    *,
    image_session=None,
    allowed_hosts: list[str] | None = None,
    public_host_validator: Callable[[str], bool] = _public_host,
) -> tuple[str, dict]:
    """Remove active content and inline validated images without leaking credentials."""
    try:
        from bs4 import BeautifulSoup, Comment
    except ImportError as exc:
        raise EmailRenderError("服务器缺少 BeautifulSoup", code="beautifulsoup_missing") from exc
    soup = BeautifulSoup(html, "html.parser")
    for comment in soup.find_all(string=lambda text: isinstance(text, Comment)):
        comment.extract()
    for tag in list(soup.find_all(list(_DANGEROUS_TAGS))):
        tag.decompose()
    for meta in list(soup.find_all("meta")):
        meta.decompose()

    for tag in soup.find_all(True):
        for attribute in list(tag.attrs):
            lowered = attribute.lower()
            if lowered.startswith("on") or lowered in _URL_ATTRIBUTES or lowered == "srcset":
                del tag.attrs[attribute]
        if tag.has_attr("style"):
            tag["style"] = _clean_css(str(tag.get("style") or ""))
    for style in soup.find_all("style"):
        style.string = _clean_css(style.get_text())

    site_host = urlparse(site_url).hostname or ""
    trusted_hosts = {_normal_host(site_host)}
    if allowed_hosts is None:
        allowed_hosts = [
            host.strip()
            for host in os.environ.get("ORDER_NOTIFICATION_IMAGE_HOST_ALLOWLIST", "").split(",")
            if host.strip()
        ]
    trusted_hosts.update(_normal_host(host) for host in allowed_hosts if host)
    fetcher = image_session or requests.Session()
    fetched = 0
    removed = 0
    total_bytes = 0
    for image_tag in list(soup.find_all("img"))[:MAX_IMAGES]:
        src = str(image_tag.get("src") or "").strip()
        if not src:
            image_tag.attrs.pop("src", None)
            continue
        inline = _validated_data_image(src)
        if inline:
            total_bytes += len(inline[0])
            if total_bytes > MAX_TOTAL_IMAGE_BYTES:
                image_tag.attrs.pop("src", None)
                removed += 1
            continue
        canonical_src = urljoin(site_url.rstrip("/") + "/", src)
        parsed = urlparse(canonical_src)
        host = parsed.hostname or ""
        try:
            port = parsed.port
        except ValueError:
            port = -1
        if (
            parsed.scheme.lower() != "https"
            or not host
            or parsed.username
            or parsed.password
            or port not in (None, 443)
            or _normal_host(host) not in trusted_hosts
            or not public_host_validator(host)
            or len(canonical_src) > 2048
        ):
            image_tag.attrs.pop("src", None)
            removed += 1
            continue
        try:
            response = fetcher.get(
                canonical_src,
                headers={"User-Agent": "HongKong-Order-Email-Image/1.0", "Accept": "image/png,image/jpeg,image/gif,image/webp"},
                timeout=(3, 12),
                allow_redirects=False,
                stream=True,
            )
            content_type = str(response.headers.get("Content-Type") or "").split(";", 1)[0].lower()
            if response.status_code != 200 or content_type not in {"image/png", "image/jpeg", "image/gif", "image/webp"}:
                raise ValueError("invalid image response")
            raw = _read_response_bytes(response, MAX_IMAGE_BYTES)
            candidate = _validated_data_image(
                "data:" + content_type + ";base64," + base64.b64encode(raw).decode("ascii")
            )
            if not candidate or total_bytes + len(raw) > MAX_TOTAL_IMAGE_BYTES:
                raise ValueError("invalid image")
            image_tag["src"] = "data:" + content_type + ";base64," + base64.b64encode(raw).decode("ascii")
            total_bytes += len(raw)
            fetched += 1
        except Exception:
            image_tag.attrs.pop("src", None)
            removed += 1
    for overflow in list(soup.find_all("img"))[MAX_IMAGES:]:
        overflow.attrs.pop("src", None)
        removed += 1

    if not soup.html:
        wrapper = BeautifulSoup("<html><head></head><body></body></html>", "html.parser")
        for child in list(soup.contents):
            wrapper.body.append(child.extract())
        soup = wrapper
    if not soup.head:
        soup.html.insert(0, soup.new_tag("head"))
    csp = soup.new_tag("meta")
    csp["http-equiv"] = "Content-Security-Policy"
    csp["content"] = "default-src 'none'; img-src data:; style-src 'unsafe-inline'; font-src 'none'; media-src 'none'; connect-src 'none'; frame-src 'none'"
    soup.head.insert(0, csp)
    viewport = soup.new_tag("meta")
    viewport["name"] = "viewport"
    viewport["content"] = "width=device-width, initial-scale=1"
    soup.head.insert(1, viewport)
    return str(soup), {"images_inlined": fetched, "images_removed": removed, "inline_image_bytes": total_bytes}


def _chromium_path(playwright_chromium) -> str | None:
    configured = os.environ.get("ORDER_NOTIFICATION_CHROMIUM_PATH")
    candidates = [
        configured,
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
        shutil.which("google-chrome"),
        shutil.which("google-chrome-stable"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        getattr(playwright_chromium, "executable_path", None),
    ]
    return next((str(path) for path in candidates if path and Path(path).is_file()), None)


def _browser_launch_env(runtime_home: str | Path) -> dict[str, str]:
    """Give hardened services a private writable HOME for Chromium/Crashpad."""
    home = Path(runtime_home).resolve()
    config_home = home / ".config"
    cache_home = home / ".cache"
    config_home.mkdir(parents=True, exist_ok=True)
    cache_home.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env.update(
        {
            "HOME": str(home),
            "XDG_CONFIG_HOME": str(config_home),
            "XDG_CACHE_HOME": str(cache_home),
        }
    )
    return env


def _save_bounded_png(image: Image.Image, path: Path) -> int:
    image = image.convert("RGB")
    image.save(path, format="PNG", optimize=True, compress_level=9)
    if path.stat().st_size <= MAX_OUTPUT_BYTES:
        return path.stat().st_size
    quantized = image.quantize(colors=192, method=Image.Quantize.FASTOCTREE)
    quantized.save(path, format="PNG", optimize=True, compress_level=9)
    if path.stat().st_size > MAX_OUTPUT_BYTES:
        raise EmailRenderError("渲染图片超过群机器人大小上限", code="rendered_image_too_large")
    return path.stat().st_size


def render_email_html(html: str, output_dir: str, job_id: str, *, chromium_path: str | None = None) -> list[dict]:
    """Render already-sanitized HTML with all browser networking aborted."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise EmailRenderError("服务器缺少 Playwright", code="playwright_missing") from exc

    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "-", str(job_id))[:120] or "email"
    full_path = output / f"{safe_id}-full.png"
    blocked_requests = 0
    try:
        with tempfile.TemporaryDirectory(prefix=".chromium-home-", dir=output) as browser_home:
            with sync_playwright() as playwright:
                executable = chromium_path or _chromium_path(playwright.chromium)
                if not executable:
                    raise EmailRenderError("服务器没有可用的 Chromium", code="chromium_missing")
                browser = playwright.chromium.launch(
                    headless=True,
                    executable_path=executable,
                    env=_browser_launch_env(browser_home),
                )
                try:
                    page = browser.new_page(viewport={"width": 720, "height": 900}, device_scale_factor=1)

                    def block_route(route):
                        nonlocal blocked_requests
                        blocked_requests += 1
                        route.abort()

                    page.route("**/*", block_route)
                    page.set_content(html, wait_until="load", timeout=20_000)
                    page.emulate_media(media="screen")
                    height = int(page.evaluate("Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"))
                    if height <= 0 or height > MAX_RENDER_HEIGHT:
                        raise EmailRenderError("邮件页面高度超出安全上限", code="email_render_height_invalid")
                    page.screenshot(path=str(full_path), full_page=True, type="png")
                finally:
                    browser.close()
    except EmailRenderError:
        raise
    except Exception as exc:
        raise EmailRenderError("邮件 HTML 渲染失败", code="email_html_render_failed", retryable=True) from exc

    rendered = []
    try:
        with Image.open(full_path) as source:
            width, height = source.size
            page_count = (height + PAGE_HEIGHT - 1) // PAGE_HEIGHT
            if page_count > MAX_PAGES:
                raise EmailRenderError("邮件图片页数超过安全上限", code="email_render_too_many_pages")
            for page_number in range(1, page_count + 1):
                top = (page_number - 1) * PAGE_HEIGHT
                page_image = source.crop((0, top, width, min(height, top + PAGE_HEIGHT)))
                path = output / f"{safe_id}-p{page_number:02d}.png"
                size = _save_bounded_png(page_image, path)
                raw_hash = hashlib.sha256(path.read_bytes()).hexdigest()
                rendered.append(
                    {
                        "page": page_number,
                        "path": str(path),
                        "sha256": raw_hash,
                        "width": page_image.width,
                        "height": page_image.height,
                        "bytes": size,
                        "blocked_requests": blocked_requests,
                    }
                )
    finally:
        try:
            full_path.unlink()
        except OSError:
            pass
    return rendered


def render_logged_admin_email(
    conn,
    order_id: str,
    output_dir: str,
    job_id: str,
    *,
    session=None,
    image_session=None,
    chromium_path: str | None = None,
) -> tuple[list[dict], dict]:
    email = fetch_admin_new_order_email(conn, order_id, session=session)
    sanitized, safety = sanitize_email_html(
        email["body"], email["site_url"], image_session=image_session
    )
    rendered = render_email_html(sanitized, output_dir, job_id, chromium_path=chromium_path)
    metadata = {key: value for key, value in email.items() if key != "body"}
    metadata.update(safety)
    metadata["template_version"] = EMAIL_TEMPLATE_VERSION
    metadata["html_sha256"] = hashlib.sha256(sanitized.encode("utf-8")).hexdigest()
    metadata["blocked_browser_requests"] = sum(item.get("blocked_requests", 0) for item in rendered)
    return rendered, metadata
