"""Deterministic, network-free PNG rendering for order notifications."""

from __future__ import annotations

import hashlib
import math
import os
import re
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont


CARD_WIDTH = 1080
MAX_IMAGE_BYTES = 2 * 1024 * 1024
TEMPLATE_VERSION = "order-card-v1"

EVENT_STYLES = {
    "ORDER_READY": ("新订单", "#136f63", "#e8f7f3"),
    "ORDER_UPDATED": ("订单变更 · 请以本卡为准", "#b45309", "#fff7ed"),
    "ORDER_CANCELLED": ("停止处理", "#b91c1c", "#fef2f2"),
    "ORDER_HOLD": ("暂停处理", "#a16207", "#fefce8"),
    "MANUAL_RESEND": ("人工重发", "#475569", "#f1f5f9"),
}


def _font_candidates(*, bold: bool = False) -> list[str]:
    configured = os.environ.get("ORDER_NOTIFICATION_FONT", "").strip()
    regular = [
        configured,
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    if not bold:
        return regular
    return [
        configured,
        "C:/Windows/Fonts/msyhbd.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc",
        *regular,
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]


def _font_stack(size: int, *, bold: bool = False) -> tuple:
    fonts = []
    seen = set()
    for candidate in _font_candidates(bold=bold):
        if not candidate or candidate in seen:
            continue
        seen.add(candidate)
        if candidate and Path(candidate).is_file():
            fonts.append(ImageFont.truetype(candidate, size=size))
    return tuple(fonts) or (ImageFont.load_default(size=size),)


_GLYPH_CACHE: dict[tuple[int, str], bool] = {}
_MISSING_GLYPH = "\U0010ffff"


def _mask_signature(font, char: str) -> tuple:
    mask = font.getmask(char, mode="L")
    return mask.size, bytes(mask)


def _supports(font, char: str) -> bool:
    if char.isspace():
        return True
    key = (id(font), char)
    cached = _GLYPH_CACHE.get(key)
    if cached is not None:
        return cached
    try:
        supported = _mask_signature(font, char) != _mask_signature(font, _MISSING_GLYPH)
    except Exception:
        supported = True
    _GLYPH_CACHE[key] = supported
    return supported


def _select_font(fonts: tuple, char: str):
    return next((font for font in fonts if _supports(font, char)), fonts[0])


def _font_runs(text: str, fonts: tuple):
    current_font = None
    current = ""
    for char in text:
        selected = _select_font(fonts, char)
        if current and selected is not current_font:
            yield current, current_font
            current = ""
        current_font = selected
        current += char
    if current:
        yield current, current_font


def _text_width(draw: ImageDraw.ImageDraw, text: str, fonts: tuple) -> float:
    return sum(draw.textlength(run, font=font) for run, font in _font_runs(text, fonts))


def _draw_text(draw: ImageDraw.ImageDraw, xy, text: Any, fonts: tuple, *, fill) -> float:
    x, y = xy
    start = x
    for run, font in _font_runs(str(text), fonts):
        draw.text((x, y), run, font=font, fill=fill)
        x += draw.textlength(run, font=font)
    return x - start


def _clean(value: Any, limit: int = 500) -> str:
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", str(value or ""))
    text = re.sub(r"\s+", " ", text).strip()
    return text[:limit]


def _wrap(draw: ImageDraw.ImageDraw, text: Any, fonts: tuple, max_width: int, max_lines: int = 4) -> list[str]:
    value = _clean(text)
    if not value:
        return ["—"]
    lines: list[str] = []
    current = ""
    for char in value:
        candidate = current + char
        if _text_width(draw, candidate, fonts) <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = char
        if len(lines) >= max_lines:
            break
    if current and len(lines) < max_lines:
        lines.append(current)
    consumed = "".join(lines)
    if len(consumed) < len(value) and lines:
        tail = lines[-1]
        while tail and _text_width(draw, tail + "…", fonts) > max_width:
            tail = tail[:-1]
        lines[-1] = tail + "…"
    return lines or ["—"]


def _text_block(draw, xy, label, value, *, label_font, value_font, width, color="#111827") -> int:
    x, y = xy
    _draw_text(draw, (x, y), label, label_font, fill="#64748b")
    y += 42
    lines = _wrap(draw, value, value_font, width, max_lines=4)
    for line in lines:
        _draw_text(draw, (x, y), line, value_font, fill=color)
        y += 44
    return y


def _chunks(items: list[dict], size: int = 7) -> list[list[dict]]:
    if not items:
        return [[]]
    return [items[i : i + size] for i in range(0, len(items), size)]


def _measure_page(snapshot: dict, items: list[dict]) -> int:
    note = _clean(snapshot.get("customer_note"), 500)
    changes = list(snapshot.get("changes") or [])[:5]
    # Fixed sections plus a 156 px footer safety zone.  Earlier revisions
    # underestimated this area, allowing notes/changes to collide with the
    # authenticated detail link on compact cards.
    base = 900
    item_height = 160 * max(1, len(items))
    note_height = min(220, 48 + 38 * max(1, math.ceil(len(note) / 44))) if note else 0
    change_height = 0
    if changes:
        wrapped_lines = 0
        for change in changes:
            approximate_length = sum(
                len(_clean(change.get(key), limit))
                for key, limit in (("field", 60), ("before", 80), ("after", 80))
            ) + 6
            wrapped_lines += min(2, max(1, math.ceil(approximate_length / 42)))
        change_height = 52 + 36 * wrapped_lines
    return max(1120, base + item_height + note_height + change_height)


def render_order_cards(
    snapshot: dict,
    event_type: str,
    output_dir: str | os.PathLike[str],
    job_id: str,
    *,
    template_version: str = TEMPLATE_VERSION,
) -> list[dict]:
    """Render one or more private PNGs and return verifiable metadata."""
    title, accent, pale = EVENT_STYLES.get(event_type, EVENT_STYLES["ORDER_HOLD"])
    items = list(snapshot.get("items") or [])
    pages = _chunks(items)
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    results: list[dict] = []

    title_font = _font_stack(48, bold=True)
    h_font = _font_stack(34, bold=True)
    body_font = _font_stack(30)
    small_font = _font_stack(24)
    label_font = _font_stack(22, bold=True)

    for page_no, page_items in enumerate(pages, 1):
        height = _measure_page(snapshot, page_items)
        image = Image.new("RGB", (CARD_WIDTH, height), "#f8fafc")
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((34, 34, CARD_WIDTH - 34, height - 34), radius=28, fill="white", outline="#cbd5e1", width=2)
        draw.rounded_rectangle((34, 34, CARD_WIDTH - 34, 178), radius=28, fill=accent)
        draw.rectangle((34, 132, CARD_WIDTH - 34, 178), fill=accent)
        _draw_text(draw, (72, 72), title, title_font, fill="white")
        _draw_text(draw, (CARD_WIDTH - 250, 88), f"{page_no}/{len(pages)} · {template_version}", small_font, fill="#e2e8f0")

        y = 215
        order_number = _clean(snapshot.get("number") or snapshot.get("order_id"), 80)
        store = _clean(snapshot.get("store_label") or snapshot.get("store_id"), 100)
        _draw_text(draw, (72, y), f"{store}  ·  订单 #{order_number}", h_font, fill="#0f172a")
        y += 64
        _draw_text(draw, (72, y), f"下单 {_clean(snapshot.get('created_at'), 50)}   通知 {_clean(snapshot.get('notification_at'), 50)}", small_font, fill="#64748b")
        y += 58

        left, right = 72, 565
        box_top = y
        draw.rounded_rectangle((60, box_top, CARD_WIDTH - 60, box_top + 220), radius=18, fill=pale)
        left_y = _text_block(
            draw, (left, box_top + 22), "仓库 / 配送", f"{snapshot.get('warehouse_name') or '待分配'} · {snapshot.get('shipping_method') or '未提供'}",
            label_font=label_font, value_font=body_font, width=430,
        )
        right_y = _text_block(
            draw, (right, box_top + 22), "付款 / 金额", f"{snapshot.get('payment_method') or '未提供'} · {snapshot.get('currency') or ''} {snapshot.get('total') or '0'}",
            label_font=label_font, value_font=body_font, width=390,
        )
        y = max(left_y, right_y, box_top + 206) + 30

        recipient = snapshot.get("recipient") or {}
        recipient_text = " · ".join(
            part for part in (
                _clean(recipient.get("name_masked"), 80),
                _clean(recipient.get("phone_masked"), 40),
                _clean(recipient.get("postal_code"), 20),
                _clean(recipient.get("city"), 80),
                _clean(recipient.get("delivery_point"), 100),
            ) if part
        )
        y = _text_block(
            draw, (72, y), "收件信息（最小化）", recipient_text,
            label_font=label_font, value_font=body_font, width=CARD_WIDTH - 144,
        ) + 16

        _draw_text(draw, (72, y), f"商品（本页 {len(page_items)} 项）", label_font, fill="#64748b")
        y += 40
        if not page_items:
            _draw_text(draw, (82, y), "无商品明细", body_font, fill="#991b1b")
            y += 62
        for item in page_items:
            draw.rounded_rectangle((62, y, CARD_WIDTH - 62, y + 142), radius=14, fill="#f8fafc", outline="#e2e8f0")
            name = _clean(item.get("name"), 240)
            variation = _clean(item.get("variation"), 160)
            sku = _clean(item.get("sku"), 100)
            qty = item.get("quantity") or 0
            name_lines = _wrap(draw, name, body_font, 740, max_lines=2)
            for line_no, line in enumerate(name_lines):
                _draw_text(draw, (82, y + 12 + line_no * 38), line, body_font, fill="#0f172a")
            details = " · ".join(p for p in (variation, sku) if p) or "SKU / 规格未提供"
            detail_line = _wrap(draw, details, small_font, 740, max_lines=1)[0]
            _draw_text(draw, (82, y + 92), detail_line, small_font, fill="#64748b")
            _draw_text(draw, (870, y + 45), f"× {qty}", h_font, fill=accent)
            y += 158

        changes = snapshot.get("changes") or []
        if changes:
            _draw_text(draw, (72, y), "变更摘要", label_font, fill=accent)
            y += 40
            for change in changes[:5]:
                line = f"• {_clean(change.get('field'), 60)}：{_clean(change.get('before'), 80)} → {_clean(change.get('after'), 80)}"
                for wrapped in _wrap(draw, line, small_font, CARD_WIDTH - 160, max_lines=2):
                    _draw_text(draw, (82, y), wrapped, small_font, fill="#7c2d12")
                    y += 36
            y += 12

        note = _clean(snapshot.get("customer_note"), 500)
        if note:
            _draw_text(draw, (72, y), "客户备注", label_font, fill="#64748b")
            y += 38
            for line in _wrap(draw, note, small_font, CARD_WIDTH - 160, max_lines=4):
                _draw_text(draw, (82, y), line, small_font, fill="#334155")
                y += 36
            y += 12

        detail = _clean(snapshot.get("internal_order_url"), 160)
        footer_y = height - 116
        draw.line((72, footer_y - 20, CARD_WIDTH - 72, footer_y - 20), fill="#e2e8f0", width=2)
        footer = f"内部详情（需登录）：{detail}" if detail else "内部详情：请登录订单系统查看"
        _draw_text(draw, (72, footer_y), footer, small_font, fill="#475569")
        _draw_text(draw, (72, footer_y + 40), "接口已接受 ≠ 群成员已阅读", small_font, fill="#94a3b8")

        path = output / f"{re.sub(r'[^A-Za-z0-9_-]', '_', job_id)}-{page_no}.png"
        image.save(path, format="PNG", optimize=True)
        raw = path.read_bytes()
        if len(raw) >= MAX_IMAGE_BYTES:
            path.unlink(missing_ok=True)
            raise ValueError("rendered_image_too_large")
        decoded = Image.open(path)
        decoded.verify()
        results.append(
            {
                "path": str(path),
                "sha256": hashlib.sha256(raw).hexdigest(),
                "width": CARD_WIDTH,
                "height": height,
                "bytes": len(raw),
                "page": page_no,
                "pages": len(pages),
                "template_version": template_version,
            }
        )
    return results
