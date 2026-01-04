
from woocommerce import API
import json

url = "https://www.vapeprimeau.com"
consumer_key = "ck_e8ff54af41a9de304614bb504665e1de53162da5"
consumer_secret = "cs_77a37538dd6c480227a1f57e28b2903411220a6c"

wcapi = API(
    url=url,
    consumer_key=consumer_key,
    consumer_secret=consumer_secret,
    version="wc/v3",
    timeout=30
)

order_id = 30898

print(f"Checking order #{order_id} on {url}...")

# 1. Check specific order
try:
    response = wcapi.get(f"orders/{order_id}")
    if response.status_code == 200:
        order = response.json()
        print(f"Order found! Status: {order.get('status')}")
        print(f"ID: {order.get('id')}")
    else:
        print(f"Order check failed: HTTP {response.status_code}")
        print(response.text)
except Exception as e:
    print(f"Error checking order: {e}")

# 2. Check if it appears in the list (cleanup logic simulation)
print("\nSimulating cleanup logic (fetching all IDs)...")
try:
    # Use same logic as app.py
    response = wcapi.get("orders", params={
        "per_page": 100, 
        "page": 1, 
        "_fields": "id,status",
        "include": [order_id] # Try to specifically include it in list if possible, or just search for it
    })
    
    # Actually, cleanup fetches ALL IDs. Let's just standard list fetch and see if we can find it in recent.
    # But to be precise, let's try to filter by ID to see if list endpoint returns it
    
    response_list = wcapi.get("orders", params={
        "include": [order_id]
    })
    
    if response_list.status_code == 200:
        orders = response_list.json()
        found = False
        for o in orders:
            if o['id'] == order_id:
                print(f"Order #{order_id} IS returned by list endpoint. Status: {o.get('status')}")
                found = True
                break
        if not found:
            print(f"Order #{order_id} is NOT returned by list endpoint (with include param).")
    else:
        print(f"List check failed: HTTP {response_list.status_code}")

except Exception as e:
    print(f"Error listing orders: {e}")
