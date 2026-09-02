#!/usr/bin/env python3
"""Read-only diagnostic for one WooCommerce order.

Credentials are resolved from the application database. The script performs
GET requests only and deliberately has no write-permission probe.
"""

from __future__ import annotations

import argparse

from woocommerce import API

from app import app, get_db_connection


def main(argv=None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--site-url", required=True)
    parser.add_argument("--order-id", required=True, type=int)
    args = parser.parse_args(argv)

    with app.app_context():
        connection = get_db_connection()
        try:
            site = connection.execute(
                "SELECT url,consumer_key,consumer_secret FROM sites WHERE url = ?",
                (args.site_url,),
            ).fetchone()
        finally:
            connection.close()

    if not site:
        print("Site not found")
        return 2

    client = API(
        url=site["url"],
        consumer_key=site["consumer_key"],
        consumer_secret=site["consumer_secret"],
        version="wc/v3",
        timeout=30,
    )
    response = client.get(
        f"orders/{args.order_id}",
        params={"_fields": "id,status"},
    )
    if response.status_code != 200:
        print(f"Read failed: HTTP {response.status_code}")
        return 1
    order = response.json()
    print(f"Read OK: order={order.get('id')} status={order.get('status')}")
    print("Write permission: NOT TESTED (production-safe read-only mode)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
