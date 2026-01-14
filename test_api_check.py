#!/usr/bin/env python3
"""Test the check_all_sites_api function"""
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
            
            # Test write
            print("\nTesting WRITE permission...")
            orders = response.json()
            if orders and len(orders) > 0:
                test_order_id = orders[0]['id']
                print(f"Using order ID: {test_order_id}")
                
                test_note_response = wcapi.post(
                    f"orders/{test_order_id}/notes",
                    data={
                        "note": "[API权限测试] 此消息用于验证写权限，将立即删除",
                        "customer_note": False
                    }
                )
                print(f"Note creation status: {test_note_response.status_code}")
                
                if test_note_response.status_code in (200, 201):
                    print("✓ Write permission OK")
                    note_id = test_note_response.json().get('id')
                    if note_id:
                        delete_resp = wcapi.delete(f"orders/{test_order_id}/notes/{note_id}")
                        print(f"Note deletion status: {delete_resp.status_code}")
                else:
                    print(f"✗ Write permission FAILED: {test_note_response.text}")
            else:
                print("No orders to test write permission")
        else:
            print(f"✗ Read permission FAILED: {response.status_code}")
            
        conn.close()
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
