#!/usr/bin/env python3
"""
Carrier delivery-status lookup (Phase 2 — auto-resolve shipped-order outcomes).

Scope today: InPost (Polish parcel lockers) via the public ShipX tracking API,
which needs no credentials and was verified reachable from this server.
DPD's public tracking page is behind an Imperva WAF (rejects non-browser
requests), so DPD is intentionally NOT handled here yet — classify_carrier
returns 'dpd' so the resolver can skip those rows until a DPD path is chosen.

This module does pure lookups: no DB access, no writes. The resolver
(resolve_outcomes.py) owns all DB side effects and policy.
"""
import time

import requests

INPOST_TRACKING_URL = "https://api-shipx-pl.easypack24.net/v1/tracking/{number}"
_HEADERS = {"User-Agent": "woo-analysis-tracker/1.0", "Accept": "application/json"}

# Map InPost ShipX raw status -> our outcome category.
#   delivered  -> 客户已签收 (auto-confirm)
#   returned   -> 包裹退回/拒收 (goes to human review list)
#   attention  -> 需关注 (failed pickup window etc.) — surfaced but not auto-acted
#   in_transit -> still moving, leave in queue
# Anything not listed defaults to 'in_transit' (the safe, no-op bucket).
INPOST_STATUS_MAP = {
    'delivered': 'delivered',
    'returned_to_sender': 'returned',
    'returned_to_agency': 'returned',
    'canceled': 'attention',
    'avizo': 'attention',
    'undelivered': 'attention',
    'claimed': 'attention',
    'rejected_by_receiver': 'returned',
}


def classify_carrier(provider, tracking_number=None):
    """Decide which carrier a shipment belongs to.

    Prefer the plugin-stored provider (AST '_wc_shipment_tracking_items'.
    tracking_provider, e.g. 'inpost-paczkomaty' / 'dpd-pl') because the
    human-facing shipping-method label is sometimes wrong. Fall back to the
    tracking-number shape: InPost numbers are 24 digits (often '6055…'),
    DPD-PL numbers are ~14 digits.
    """
    p = (provider or '').lower()
    if 'inpost' in p or 'paczko' in p:
        return 'inpost'
    if 'dpd' in p:
        return 'dpd'
    # Provider names some OTHER real carrier (australia-post, ems, gls, …):
    # do NOT guess from the number shape — hand it to Track718 to auto-detect.
    if p and not p.startswith('custom') and p not in ('', 'unknown', 'auto', 'other'):
        return 'unknown'
    # No useful provider → guess by number FORMAT.
    raw = (tracking_number or '').strip()
    digits = ''.join(ch for ch in raw if ch.isdigit())
    if raw.isdigit() and len(raw) >= 20:   # InPost = a PURE 24-digit number
        return 'inpost'                    # (AusPost 'R…' / others have letters → not InPost)
    if 11 <= len(digits) <= 16:            # DPD-PL = ~13 digits (often + trailing letter)
        return 'dpd'
    return 'unknown'


def inpost_status(tracking_number, timeout=12, session=None):
    """Query InPost ShipX for one parcel.

    Returns a dict:
      {'ok': True, 'raw': 'delivered', 'outcome': 'delivered',
       'delivered_at': '2026-04-09T09:03:05...', 'last_event': '...'}
    or on failure:
      {'ok': False, 'error': 'http_404' | 'timeout' | 'conn' | 'parse', 'detail': ...}
    A 404 means InPost has no record (number not theirs, or purged).
    """
    tn = (tracking_number or '').strip()
    if not tn:
        return {'ok': False, 'error': 'empty'}
    url = INPOST_TRACKING_URL.format(number=tn)
    getter = session or requests
    try:
        r = getter.get(url, headers=_HEADERS, timeout=timeout)
    except requests.exceptions.Timeout:
        return {'ok': False, 'error': 'timeout'}
    except requests.exceptions.RequestException as e:
        return {'ok': False, 'error': 'conn', 'detail': str(e)}

    if r.status_code == 404:
        return {'ok': False, 'error': 'http_404'}
    if r.status_code != 200:
        return {'ok': False, 'error': f'http_{r.status_code}'}

    try:
        d = r.json()
    except ValueError:
        return {'ok': False, 'error': 'parse'}

    raw = (d.get('status') or '').strip()
    outcome = INPOST_STATUS_MAP.get(raw, 'in_transit')
    details = d.get('tracking_details') or []
    last = details[-1] if details else {}
    delivered_at = None
    if outcome == 'delivered':
        # find the timestamp of the delivered/confirmed event if present
        for ev in reversed(details):
            if (ev.get('status') or '') in ('delivered', 'confirmed'):
                delivered_at = ev.get('datetime')
                break
        delivered_at = delivered_at or last.get('datetime')
    return {
        'ok': True,
        'raw': raw,
        'outcome': outcome,
        'delivered_at': delivered_at,
        'last_event': last.get('status', ''),
        'last_event_at': last.get('datetime', ''),
    }


# ───────────────────────── DPD via Track718 aggregator ─────────────────────
# DPD-PL's public page is behind an F5/TSPD JS challenge (un-scrapable from a
# plain HTTP client), so DPD goes through Track718. Track718 is async, like
# most aggregators: you ADD a number (POST /v2/tracks), it crawls the carrier
# in the background, then you QUERY (POST /v2/tracking/query). A just-added
# number returns result 0 (NotFound) / 10 (NotOnline) until the crawl lands,
# so it resolves over subsequent (cron) runs. Auth is a plain header, no sign.
TRACK718_ADD_URL = "https://apigetway.track718.net/v2/tracks"
TRACK718_QUERY_URL = "https://apigetway.track718.net/v2/tracking/query"
TRACK718_DPD_PL = "dpd-pl"          # Track718 courier code for DPD Poland
TRACK718_INPOST_PL = "in-post"      # Track718 courier code for InPost Poland (波兰Inpost, id 358).
                                    # Fallback when InPost's own ShipX business API 404s — e.g.
                                    # Paczkomat self-service parcels that aren't in ShipX at all.
                                    # (NB: 'paczkomaty'/id 1101 exists too but returned no data in
                                    # testing; 'in-post' is what the consumer site/website uses.)
TRACK718_ADD_BATCH = 100            # API max per /v2/tracks call
TRACK718_QUERY_BATCH = 20           # API max per /v2/tracking/query call

# Track718 main "result" code -> our normalized outcome.
#   40 Delivered | 65/66/67 Returned-family | 8/11/12/13/30/31 in transit
#   20 Undelivered / 50 Alert / 51 TooLong -> attention
#   0 NotFound / 10 NotOnline -> not actionable yet (mapped to None -> retry)
TRACK718_RESULT_MAP = {
    40: 'delivered',
    65: 'returned', 66: 'returned', 67: 'returned',
    8: 'in_transit', 11: 'in_transit', 12: 'in_transit', 13: 'in_transit', 30: 'in_transit', 31: 'in_transit',
    20: 'attention', 50: 'attention', 51: 'attention',
}


def _chunks(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


def _t718_headers(key):
    return {'Content-Type': 'application/json', 'Track718-API-Key': key}


def track718_add(items, key, timeout=25, session=None):
    """Register DPD shipments so Track718 starts crawling them.
    items: list of dicts, each at least {'trackNum': ...}; we also pass
    code/innerNum/country/zip/ondate when available. Returns count added."""
    getter = session or requests
    added = 0
    for batch in _chunks(list(items), TRACK718_ADD_BATCH):
        try:
            r = getter.post(TRACK718_ADD_URL, json=batch, headers=_t718_headers(key), timeout=timeout)
            added += int((r.json().get('data') or {}).get('added') or 0)
        except (requests.exceptions.RequestException, ValueError, TypeError):
            continue
    return added


def track718_query(numbers, key, code=TRACK718_DPD_PL, timeout=25, session=None):
    """Query status for already-added numbers. Returns
    {trackNum: {'ok':True,'result':int,'outcome':..,'event':..,'event_at':..}}.
    Numbers without a usable status yet (result 0/10) come back ok=False."""
    getter = session or requests
    out = {}
    for batch in _chunks(list(numbers), TRACK718_QUERY_BATCH):
        body = [{'trackNum': n, 'code': code} for n in batch]
        try:
            r = getter.post(TRACK718_QUERY_URL, json=body, headers=_t718_headers(key), timeout=timeout)
            d = r.json()
        except (requests.exceptions.RequestException, ValueError):
            for n in batch:
                out[n] = {'ok': False, 'error': 'conn'}
            continue
        for row in ((d.get('data') or {}).get('list') or []):
            num = row.get('trackNum')
            result = row.get('result', 0)
            outcome = TRACK718_RESULT_MAP.get(result)
            latest = row.get('latest') or {}
            if outcome:
                out[num] = {'ok': True, 'result': result, 'outcome': outcome,
                            'event': latest.get('trackContent', ''), 'event_at': latest.get('trackTime', '')}
            else:
                out[num] = {'ok': False, 'error': 'no_info', 'result': result}
    return out


def track718_detail(number, key, code=None, timeout=25, session=None, poll=3, poll_wait=2.0):
    """On-demand single-number lookup returning the FULL event timeline
    (for the 查物流 button). `code` forces a carrier (e.g. 'dpd-pl'); when None,
    Track718 AUTO-DETECTS the carrier — covering EMS/中国邮政, Australia Post,
    GLS, etc. Track718 is async (add → crawl → query), so we poll the query a
    few times over a few seconds to catch numbers that crawl quickly. Returns
    detected `carrier` + time-sorted events flattened from fromDetail+toDetail."""
    getter = session or requests
    item = {'trackNum': number}
    if code:
        item['code'] = code
    # Register; if Track718 can't auto-detect the carrier (errorCode 40013) it
    # returns candidate codes in `otherCodes` — retry the add with the first
    # one (e.g. 'australia-post') so the number actually gets registered/crawled.
    try:
        ar = getter.post(TRACK718_ADD_URL, json=[item], headers=_t718_headers(key), timeout=timeout).json()
        errs = ((ar.get('data') or {}).get('errors')) or []
        if errs and not item.get('code'):
            oc = errs[0].get('otherCodes') or []
            if oc:
                item['code'] = oc[0]
                getter.post(TRACK718_ADD_URL, json=[item], headers=_t718_headers(key), timeout=timeout)
    except (requests.exceptions.RequestException, ValueError):
        pass
    last = {'ok': False, 'error': 'no_info', 'events': []}
    for attempt in range(max(1, poll)):
        if attempt:
            time.sleep(poll_wait)
        try:
            r = getter.post(TRACK718_QUERY_URL, json=[item], headers=_t718_headers(key), timeout=timeout)
            rows = ((r.json().get('data') or {}).get('list')) or []
        except (requests.exceptions.RequestException, ValueError):
            last = {'ok': False, 'error': 'conn', 'events': []}
            continue
        if not rows:
            continue
        row = rows[0]
        result = row.get('result', 0)
        outcome = TRACK718_RESULT_MAP.get(result)
        events = []
        for seg in ('toDetail', 'fromDetail'):
            for e in (row.get(seg) or []):
                t = e.get('date', '')          # Track718 event fields: date + status
                if not t:
                    continue
                ai = e.get('addressInfo') or {}
                loc = e.get('address') or ' '.join(p for p in (ai.get('city', ''), ai.get('country', '')) if p) or ''
                events.append({'time': t, 'status': ((loc + ' · ') if loc else '') + (e.get('status', '') or '')})
        events.sort(key=lambda x: x['time'], reverse=True)
        last = {'ok': bool(events or outcome), 'result': result, 'outcome': outcome or 'unknown',
                'carrier': row.get('code') or code or '', 'events': events,
                'error': None if (events or outcome) else 'no_info'}
        # Stop as soon as we have the timeline OR a definitive status. Track718
        # often returns the result code (e.g. 40=delivered) a few polls before the
        # event list fills in, so don't keep waiting once the outcome is known.
        if events or (outcome and outcome != 'unknown'):
            break
    return last


def lookup(carrier, tracking_number, key=None, **kw):
    """Single-number dispatch. InPost is synchronous. DPD via Track718 needs
    add → crawl → query, so a first lookup right after add may return
    error='no_info' (retry shortly)."""
    if carrier == 'inpost':
        return inpost_status(tracking_number, **kw)
    if carrier == 'dpd':
        if not key:
            return {'ok': False, 'error': 'no_track718_key'}
        track718_add([{'trackNum': tracking_number, 'code': TRACK718_DPD_PL}], key)
        res = track718_query([tracking_number], key)
        return res.get(tracking_number, {'ok': False, 'error': 'no_info'})
    return {'ok': False, 'error': 'unknown_carrier'}
