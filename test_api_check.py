#!/usr/bin/env python3
"""Read-only WooCommerce API connectivity diagnostic.

This file intentionally performs no write-permission probe.  Production write
paths are verified with mocks plus their durable operation/read-back ledger.
"""
import sys
sys.path.insert(0, '/www/wwwroot/woo-analysis')

from app import app, get_db_connection
from woocommerce import API

# Test the function
with app.app_context():
    try:
        conn = get_db_connection()
        sites = conn.execute('SELECT * FROM sites LIMIT 1').fetchall()
        
        if not sites:
            print("No sites found")
            sys.exit(1)
        
        site = sites[0]
        print(f"Testing site: {site['url']}")
        print(f"Site ID: {site['id']}")
        
        # Test the API call
        wcapi = API(
            url=site['url'],
            consumer_key=site['consumer_key'],
            consumer_secret=site['consumer_secret'],
            version="wc/v3",
            timeout=15
        )
        
        # Test read
        print("\nTesting READ permission...")
        response = wcapi.get("orders", params={"per_page": 1})
        print(f"Status: {response.status_code}")
        
        if response.status_code == 200:
            print("✓ Read permission OK")
            print("Write permission: NOT TESTED (production-safe read-only mode)")
        else:
            print(f"✗ Read permission FAILED: {response.status_code}")
            
        conn.close()
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
