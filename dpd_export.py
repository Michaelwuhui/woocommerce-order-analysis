"""Pure helpers for Polish DPD export identity and reference rules."""

import json
import re
from collections import defaultdict
from datetime import datetime, timezone
from urllib.parse import urlparse


SITE_ABBREVIATIONS = {
    'vapesklep': 'vsklep',
}


def _json_value(value, expected_type):
    if isinstance(value, expected_type):
        return value
    if not value:
        return expected_type()
    try:
        parsed = json.loads(value)
    except (TypeError, ValueError):
        return expected_type()
    return parsed if isinstance(parsed, expected_type) else expected_type()


def _is_placeholder_phone(digits):
    if not digits or len(set(digits)) <= 1:
        return True
    if len(digits) >= 6:
        differences = {
            int(digits[index + 1]) - int(digits[index])
            for index in range(len(digits) - 1)
        }
        if differences == {1} or differences == {-1}:
            return True
    return False


def _normalize_phone(value):
    digits = ''.join(character for character in str(value or '') if character.isdigit())
    if len(digits) < 7:
        return None
    canonical = digits[-9:] if len(digits) >= 9 else digits
    if _is_placeholder_phone(canonical) or _is_placeholder_phone(digits):
        return None
    return canonical


def _normalize_address(address):
    street = re.sub(r'\s+', ' ', str(address.get('address_1') or '').strip().lower())
    postcode = str(address.get('postcode') or '').strip().replace(' ', '').lower()
    city = re.sub(r'\s+', ' ', str(address.get('city') or '').strip().lower())
    parts = [part for part in (street, postcode, city) if part]
    if not street or len(parts) < 2:
        return None
    return ' | '.join(parts)


def _parse_datetime(value):
    raw = str(value or '').strip().replace('Z', '+00:00')
    if not raw:
        return None
    try:
        parsed = datetime.fromisoformat(raw)
    except ValueError:
        return None
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(timezone.utc).replace(tzinfo=None)
    return parsed


def _identity(order):
    billing = _json_value(order.get('billing'), dict)
    shipping = _json_value(order.get('shipping'), dict)
    address = dict(shipping if shipping.get('address_1') else billing)

    if not address.get('address_1'):
        custom = {}
        for item in _json_value(order.get('meta_data'), list):
            if isinstance(item, dict):
                custom[item.get('key')] = item.get('value') or ''
        street = ' '.join(filter(None, [
            custom.get('_billing_adres_dpd'),
            custom.get('_billing_numer_domu'),
        ])).strip()
        if street:
            address['address_1'] = street
            address['postcode'] = address.get('postcode') or custom.get('_billing_kod_pocztowy', '')
            address['city'] = address.get('city') or custom.get('_billing_miejscowosc', '')

    keys = set()
    email = str(billing.get('email') or '').strip().lower()
    phone = _normalize_phone(address.get('phone') or billing.get('phone'))
    normalized_address = _normalize_address(address)
    if email:
        keys.add(('email', email))
    if phone:
        keys.add(('phone', phone))
    if normalized_address:
        keys.add(('address', normalized_address))
    return _parse_datetime(order.get('date_created')), keys


def find_recent_duplicate_order_ids(orders, candidate_ids, window_hours=72):
    """Return candidates sharing email, phone, or full address within the window."""
    candidate_ids = set(candidate_ids)
    index = defaultdict(list)
    identities = {}
    for order in orders:
        order_id = order.get('id')
        created_at, keys = _identity(order)
        if order_id is None or created_at is None or not keys:
            continue
        identities[order_id] = (created_at, keys)
        for key in keys:
            index[key].append((order_id, created_at))

    duplicate_ids = set()
    maximum_seconds = float(window_hours) * 3600
    for order_id in candidate_ids:
        identity = identities.get(order_id)
        if not identity:
            continue
        created_at, keys = identity
        if any(
            other_id != order_id
            and abs((other_created_at - created_at).total_seconds()) <= maximum_seconds
            for key in keys
            for other_id, other_created_at in index[key]
        ):
            duplicate_ids.add(order_id)
    return duplicate_ids


def site_order_reference(source, order_number):
    """Build the partner customer reference from site label plus order number."""
    parsed = urlparse(str(source or ''))
    hostname = (parsed.hostname or parsed.path or '').lower().split(':', 1)[0]
    if hostname.startswith('www.'):
        hostname = hostname[4:]
    label = hostname.split('.', 1)[0]
    site_label = SITE_ABBREVIATIONS.get(label, label)
    number = str(order_number or '').strip().lstrip('#').replace(' ', '')
    return f'{site_label}{number}'
