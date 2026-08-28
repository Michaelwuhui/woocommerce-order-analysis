"""Helpers for the Polish InPost Paczkomat text export."""

import re
from decimal import Decimal, InvalidOperation
from io import BytesIO


def _clean_field(value):
    text = str(value or '').strip()
    return re.sub(r'[;\r\n]+', ' ', text)


def normalize_polish_phone(value):
    """Return a Polish phone number in the +48XXXXXXXXX format."""
    digits = ''.join(character for character in str(value or '') if character.isdigit())
    if digits.startswith('0048') and len(digits) == 13:
        digits = digits[2:]
    if len(digits) == 9:
        digits = f'48{digits}'
    if len(digits) != 11 or not digits.startswith('48'):
        return ''
    return f'+{digits}'


def extract_inpost_locker_code(value):
    """Extract the Paczkomat code from a code-plus-address checkout value."""
    first_part = _clean_field(value).split(',', 1)[0].strip().upper()
    for token in re.findall(r'[A-Z0-9-]+', first_part):
        if token != 'PACZKOMAT' and any(ch.isalpha() for ch in token) and any(ch.isdigit() for ch in token):
            return token
    return ''


def format_collection_amount(value):
    """Match the partner sample: decimal notation without redundant zeroes."""
    try:
        amount = Decimal(str(value).strip())
    except (InvalidOperation, TypeError, ValueError):
        return ''
    if not amount.is_finite() or amount < 0:
        return ''
    text = format(amount, 'f')
    if '.' in text:
        text = text.rstrip('0').rstrip('.')
    return text or '0'


def build_inpost_txt(rows):
    """Build an UTF-8, no-BOM, CRLF-delimited InPost import file."""
    lines = []
    for item in rows:
        fields = [
            _clean_field(item.get('email')),
            normalize_polish_phone(item.get('phone')),
            'A',
            'paczkomat',
            _clean_field(item.get('customer_number')),
            format_collection_amount(item.get('cod_amount')),
            extract_inpost_locker_code(item.get('locker_code')),
        ]
        lines.append(';'.join(fields))
    return BytesIO('\r\n'.join(lines).encode('utf-8'))
