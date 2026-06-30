#!/usr/bin/env python3
"""
Auto-DETECT shipped-order delivery status from carriers (Phase 2).

Per the agreed workflow this script ONLY detects and records the carrier
status into orders.carrier_status — it never auto-confirms delivery, never
marks undelivered, and never touches WooCommerce. A human (发货管理员) reviews
the detected status in the 「待确认结局」queue and approves; approval is what
sets delivery_confirmed / pushes WooCommerce. So this job is safe to run
unattended on cron.

Carriers:
  InPost — free public ShipX API (synchronous).
  DPD    — via 17track aggregator (async: register once, then poll). The
           17track key lives in settings.track17_api_key.

carrier_status stores the NORMALIZED outcome:
  'delivered' | 'returned' | 'attention' | 'in_transit' | 'unknown'

    venv/bin/python resolve_outcomes.py            # DRY-RUN (no writes, no 17track quota spend)
    venv/bin/python resolve_outcomes.py --live      # write carrier_status; registers DPD numbers w/ 17track
    venv/bin/python resolve_outcomes.py --live --carrier inpost   # one carrier only
"""
import sqlite3
import json
import time
import sys
import argparse
from datetime import datetime, timedelta
from collections import Counter, defaultdict

import os
import requests

# Run from the script's own dir so relative paths (DB_FILE, `import
# carrier_tracking`) work no matter the caller's cwd (e.g. cron without cd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import carrier_tracking as ct

DB_FILE = 'woocommerce_orders.db'
DEFAULT_MIN_AGE_DAYS = None
INPOST_THROTTLE = 0.35


def get_conn():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def get_setting(conn, key):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row['value'] if row else None


def extract_tracking(order):
    """Mirror app.process_shipped_order priority. Returns (number, provider)."""
    try:
        md = json.loads(order['meta_data'] or '[]')
    except (ValueError, TypeError):
        md = []
    try:
        li = json.loads(order['line_items'] or '[]')
    except (ValueError, TypeError):
        li = []
    try:
        sl = json.loads(order['shipping_lines'] or '[]')
    except (ValueError, TypeError):
        sl = []

    for m in md if isinstance(md, list) else []:
        if isinstance(m, dict) and m.get('key') == '_wc_shipment_tracking_items':
            v = m.get('value') or []
            if isinstance(v, list) and v and isinstance(v[0], dict) and v[0].get('tracking_number'):
                return str(v[0]['tracking_number']).strip(), str(v[0].get('tracking_provider', ''))
    for it in li if isinstance(li, list) else []:
        if not isinstance(it, dict):
            continue
        for m in it.get('meta_data', []):
            if not isinstance(m, dict):
                continue
            if m.get('key') == '_vi_wot_order_item_tracking_data':
                try:
                    td = m.get('value')
                    td = json.loads(td) if isinstance(td, str) else td
                    if isinstance(td, list) and td and td[0].get('tracking_number'):
                        return str(td[0]['tracking_number']).strip(), str(td[0].get('carrier_slug') or td[0].get('carrier_name') or '')
                except (ValueError, TypeError, AttributeError, KeyError):
                    pass
            if m.get('key') == 'tracking_number' and str(m.get('value', '')).strip():
                return str(m['value']).strip(), ''
    for it in sl if isinstance(sl, list) else []:
        if isinstance(it, dict):
            for m in it.get('meta_data', []):
                if isinstance(m, dict) and m.get('key') == 'tracking_number' and str(m.get('value', '')).strip():
                    return str(m['value']).strip(), ''
    provider = ''
    for m in md if isinstance(md, list) else []:
        if isinstance(m, dict) and m.get('key') == '_tracking_provider':
            provider = str(m.get('value', ''))
    for m in md if isinstance(md, list) else []:
        if isinstance(m, dict) and m.get('key') == '_tracking_number' and str(m.get('value', '')).strip():
            return str(m['value']).strip(), provider
    return None, ''


def fetch_candidates(conn, min_age_days, limit, recheck_hours, live, au_sites=None):
    cutoff = None
    if min_age_days is not None:
        cutoff = (datetime.now() - timedelta(days=min_age_days)).strftime('%Y-%m-%dT%H:%M:%S')
    recheck_clause = ''
    if live and recheck_hours is not None:
        # On cron runs, skip rows whose carrier_status was refreshed recently,
        # UNLESS they're terminal-but-unconfirmed (delivered/returned) — those
        # we keep showing; they leave the candidate set once a human acts.
        recheck_clause = f"""
          AND (carrier_status_at IS NULL
               OR carrier_status NOT IN ('delivered','returned')
               AND carrier_status_at <= datetime('now', '-{int(recheck_hours)} hours'))"""
    # Payment scope: always COD orders (every market). When AU auto-track is on,
    # ALSO include Australian-site orders — they're online-paid (not COD), so
    # they'd otherwise never be picked up here.
    params = []
    if au_sites:
        ph = ','.join(['?'] * len(au_sites))
        pay_clause = f"(o.payment_method = 'cod' OR o.source IN ({ph}))"
        params.extend(au_sites)
    else:
        pay_clause = "o.payment_method = 'cod'"
    age_clause = ''
    shipped_expr = "datetime(replace(substr(COALESCE(sl.shipped_at, o.date_modified, o.date_created), 1, 19), 'T', ' '))"
    if cutoff is not None:
        age_clause = "AND " + shipped_expr + " <= ?"
        params.append(cutoff)
    else:
        age_clause = """
          AND datetime(replace(substr(COALESCE(sl.shipped_at, o.date_modified, o.date_created), 1, 19), 'T', ' ')) <= datetime(
              'now',
              '-' || CASE COALESCE(s.country, '')
                       WHEN 'PL' THEN 1
                       WHEN 'AU' THEN 10
                       ELSE 7
                     END || ' days'
          )
        """
    q = f"""
        SELECT o.id, o.number, o.date_created, o.billing, o.meta_data, o.line_items, o.shipping_lines
        FROM orders o
        LEFT JOIN sites s ON o.source = s.url
        LEFT JOIN shipping_logs sl ON sl.id = (
            SELECT id FROM shipping_logs WHERE order_id = o.id ORDER BY id DESC LIMIT 1
        )
        WHERE {pay_clause}
          AND o.status IN ('on-hold', 'shipped', 'partial-shipped')
          AND COALESCE(o.is_undelivered, 0) = 0
          AND COALESCE(o.is_problem_return, 0) = 0
          AND COALESCE(o.delivery_confirmed, 0) = 0
          {age_clause}
          {recheck_clause}
        ORDER BY o.date_created DESC
    """
    rows = conn.execute(q, params).fetchall()
    return rows[:limit] if limit else rows


def write_status(conn, order_id, outcome):
    conn.execute("UPDATE orders SET carrier_status=?, carrier_status_at=datetime('now') WHERE id=?",
                 (outcome, order_id))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--live', action='store_true', help='write carrier_status (default: dry-run, no writes, no 17track quota spend)')
    ap.add_argument('--limit', type=int, default=0)
    ap.add_argument('--min-age-days', type=int, default=DEFAULT_MIN_AGE_DAYS,
                    help='override country defaults (PL=3, AU=10, other=7)')
    ap.add_argument('--recheck-hours', type=int, default=12)
    ap.add_argument('--carrier', choices=['inpost', 'dpd', 'other', 'all'], default='all')
    ap.add_argument('--throttle', type=float, default=INPOST_THROTTLE)
    args = ap.parse_args()
    live = args.live

    conn = get_conn()
    has_cols = 'carrier_status' in {r[1] for r in conn.execute("PRAGMA table_info(orders)")}
    if live and not has_cols:
        print("ERROR: carrier_status columns missing. ALTER TABLE orders ADD COLUMN carrier_status TEXT / carrier_status_at TEXT first.")
        sys.exit(1)
    key718 = get_setting(conn, 'track718_api_key')

    # AU auto-track toggle (系统设置). When ON, Australian-site orders (online-paid,
    # non-COD, typically EMS) join the candidate set and are resolved via Track718.
    au_enabled = (get_setting(conn, 'auto_track_au') or '').strip().lower() in ('1', 'true', 'on', 'yes')
    au_sites = []
    if au_enabled:
        au_sites = [r['url'] for r in conn.execute("SELECT url FROM sites WHERE country='AU'").fetchall()]

    candidates = fetch_candidates(conn, args.min_age_days, args.limit, args.recheck_hours, live, au_sites)
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] resolve_outcomes "
          f"{'LIVE' if live else 'DRY-RUN'} — {len(candidates)} candidates, carrier={args.carrier}, "
          f"AU auto-track={'ON' if au_enabled else 'OFF'}({len(au_sites)} sites)")

    # Bucket by carrier
    inpost = []   # (order_id, number)
    dpd = []      # (order_id, number, row)
    generic = []  # (order_id, number, provider) — EMS/中国邮政/Australia Post/… via Track718 auto-detect
    skipped = Counter()
    for o in candidates:
        number, provider = extract_tracking(o)
        if not number:
            skipped['no_tracking'] += 1
            continue
        carrier = ct.classify_carrier(provider, number)
        if carrier == 'inpost' and args.carrier in ('inpost', 'all'):
            inpost.append((o['id'], number))
        elif carrier == 'dpd' and args.carrier in ('dpd', 'all'):
            dpd.append((o['id'], number, o))
        elif carrier == 'unknown' and args.carrier in ('other', 'all'):
            generic.append((o['id'], number, provider))
        else:
            skipped[carrier] += 1

    outcomes = Counter()
    written = 0
    session = requests.Session()

    # ---- InPost (synchronous, free public ShipX API) ----
    inpost_fb = []   # ShipX 404 (often Paczkomat self-service) → Track718 fallback
    for oid, num in inpost:
        info = ct.inpost_status(num, session=session)
        if not info['ok']:
            if info['error'] == 'http_404' and key718:
                inpost_fb.append((oid, num))   # recover via Track718 'in-post' below
            else:
                outcomes['inpost:err:' + info['error']] += 1
            time.sleep(args.throttle)
            continue
        outcomes['inpost:' + info['outcome']] += 1
        if live:
            write_status(conn, oid, info['outcome'])
            written += 1
        time.sleep(args.throttle)

    # ---- InPost ShipX-404 fallback via Track718 (async: add -> crawl -> query) ----
    # Parcels the ShipX business API doesn't hold (e.g. Paczkomat self-service)
    # are tracked by Track718 under the 'in-post' courier code. Same pattern as
    # DPD: register now, a later cron run reads the crawled status. This both
    # clears these from the queue AND restores refusal visibility (returned→🔴).
    if inpost_fb and key718:
        if live:
            items = [{'trackNum': num, 'code': ct.TRACK718_INPOST_PL,
                      'innerNum': str(oid), 'country': 'PL'} for oid, num in inpost_fb]
            added = ct.track718_add(items, key718, session=session)
            print(f"  Track718 InPost(in-post) add: {added} 个已登记（刚登记的本轮多为未上线，下次 cron 出状态）")
        nums = [num for _, num in inpost_fb]
        statuses = ct.track718_query(nums, key718, code=ct.TRACK718_INPOST_PL, session=session)
        for oid, num in inpost_fb:
            r = statuses.get(num, {})
            if not r.get('ok'):
                outcomes['inpost718:no_info'] += 1
                continue
            outcomes['inpost718:' + r['outcome']] += 1
            if live:
                write_status(conn, oid, r['outcome'])
                written += 1

    # ---- DPD (Track718, async: add -> crawl -> query) ----
    if dpd:
        if not key718:
            print("  ⚠ DPD 跳过: settings.track718_api_key 未配置")
        else:
            if live:
                items = []
                for oid, num, o in dpd:
                    try:
                        b = json.loads(o['billing'] or '{}')
                    except (ValueError, TypeError):
                        b = {}
                    items.append({'trackNum': num, 'code': ct.TRACK718_DPD_PL,
                                  'innerNum': str(o['number']), 'country': (b.get('country') or 'PL'),
                                  'zip': (b.get('postcode') or ''),
                                  'ondate': ((o['date_created'] or '')[:10] or '2020-01-01') + 'T00:00:00Z'})
                added = ct.track718_add(items, key718, session=session)
                print(f"  Track718 add: {added} 个已登记（刚登记的本轮多为未上线，下次 cron 出状态）")
            nums = [num for _, num, _ in dpd]
            statuses = ct.track718_query(nums, key718, session=session)
            for oid, num, o in dpd:
                r = statuses.get(num, {})
                if not r.get('ok'):
                    outcomes['dpd:no_info'] += 1
                    continue
                outcomes['dpd:' + r['outcome']] += 1
                if live:
                    write_status(conn, oid, r['outcome'])
                    written += 1

    # ---- Other carriers (EMS/中国邮政, Australia Post, GLS, …) via Track718 ----
    # Non-InPost/non-DPD shipments — chiefly the AU market's EMS parcels. Track718
    # AUTO-DETECTS the carrier; track718_detail handles add (+otherCodes retry) and
    # polls the query. Already-crawled numbers return on the first poll, so steady
    # state is fast; brand-new numbers may need a later run to surface a status.
    # It registers numbers (spends quota), so we only call it on --live runs.
    if generic:
        if not key718:
            print("  ⚠ 其他物流(EMS等) 跳过: settings.track718_api_key 未配置")
        elif not live:
            print(f"  [DRY-RUN] 其他物流(EMS等) 候选 {len(generic)} 单，未查询（避免 Track718 额度消耗）")
            for _ in generic:
                outcomes['other:dry'] += 1
        else:
            for oid, num, prov in generic:
                res = ct.track718_detail(num, key718, code=None, session=session, poll=2, poll_wait=2.0)
                oc = res.get('outcome')
                if oc and oc != 'unknown':
                    outcomes['other:' + oc] += 1
                    write_status(conn, oid, oc)
                    written += 1
                else:
                    outcomes['other:no_info'] += 1
                time.sleep(args.throttle)

    if live:
        conn.commit()
    conn.close()

    print("\n— 候选归类 —")
    print(f"   InPost {len(inpost)}   InPost查无→Track718 {len(inpost_fb)}   DPD {len(dpd)}   其他(EMS等) {len(generic)}   跳过 {dict(skipped)}")
    print("\n— 检测结果（normalized outcome）—")
    agg = Counter()
    for k, n in outcomes.most_common():
        print(f"   {k:26s} {n}")
        if ':' in k and not k.endswith(('err', 'no_info')):
            agg[k.split(':', 1)[1]] += n
    print("\n— 汇总 —")
    for k in ('delivered', 'returned', 'attention', 'in_transit'):
        if agg.get(k):
            print(f"   {k:12s} {agg[k]}")
    if live:
        print(f"\n已写入 carrier_status: {written} 单（仅检测标记，未确认、未推 WooCommerce）")
    else:
        print(f"\n[DRY-RUN] 已签收 {agg.get('delivered',0)} · 退回 {agg.get('returned',0)} · 在途 {agg.get('in_transit',0)}。未写库、未花 17track 额度。")


if __name__ == '__main__':
    main()
