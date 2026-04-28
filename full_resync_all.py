#!/usr/bin/env python3
"""One-off: real full resync for every site, no date filter, fetch every page.

Run: ./venv/bin/python full_resync_all.py
"""
import sqlite3
import time
import random
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sync_utils
from woocommerce import API

DB_FILE = 'woocommerce_orders.db'

def main():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    sites = conn.execute(
        'SELECT id, url, consumer_key, consumer_secret FROM sites '
        'WHERE consumer_key IS NOT NULL AND consumer_key != ""'
    ).fetchall()
    conn.close()

    summary = []
    for site in sites:
        url = site['url']
        print(f"\n=========================")
        print(f">>> Full resync: {url}")
        print(f"=========================")
        before = _count(url)
        try:
            wcapi = API(
                url=url,
                consumer_key=site['consumer_key'],
                consumer_secret=site['consumer_secret'],
                version='wc/v3',
                timeout=60,
            )
            total = 0
            page = 1
            per_page = 100
            while True:
                time.sleep(random.uniform(0.4, 0.8))
                try:
                    resp = wcapi.get('orders', params={
                        'per_page': per_page, 'page': page,
                        'expand': 'line_items,shipping_lines,tax_lines,fee_lines,coupon_lines,refunds'
                    })
                except Exception as e:
                    print(f"  Network error on page {page}: {e}")
                    time.sleep(5)
                    continue
                if resp.status_code != 200:
                    print(f"  HTTP {resp.status_code} on page {page}: {resp.text[:200]}")
                    break
                data = resp.json()
                if not data:
                    print(f"  Page {page}: empty, finished")
                    break
                for o in data:
                    o['source'] = url
                sync_utils.save_orders_to_db(data)
                total += len(data)
                first_date = data[0].get('date_created', '')
                last_date = data[-1].get('date_created', '')
                print(f"  Page {page}: {len(data)} orders ({first_date} ~ {last_date}), running total {total}")
                page += 1
                if page > 500:
                    print("  Safety break at page 500")
                    break

            after = _count(url)
            summary.append((url, before, after, total))
            print(f">>> {url}: DB before={before}, after={after}, fetched={total}")
        except Exception as e:
            print(f">>> ERROR on {url}: {e}")
            summary.append((url, before, _count(url), -1))

    # Print summary
    print("\n\n=== Summary ===")
    for url, before, after, fetched in summary:
        delta = after - before
        flag = "+" if delta > 0 else " "
        print(f"  {flag}{delta:>6}  {url}  ({before} -> {after}, fetched {fetched})")


def _count(url):
    conn = sqlite3.connect(DB_FILE)
    n = conn.execute('SELECT COUNT(*) FROM orders WHERE source = ?', (url,)).fetchone()[0]
    conn.close()
    return n


if __name__ == '__main__':
    main()
