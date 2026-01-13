
import sqlite3
import json

DB_FILE = '/www/wwwroot/woo-analysis/woocommerce_orders.db'

def inspect_order(order_number):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("SELECT * FROM orders WHERE number = ?", (order_number,))
    orders = cursor.fetchall()
    
    if not orders:
        print(f"Order #{order_number} not found.")
        return

    print(f"Found {len(orders)} orders with number {order_number}.")
    
    for order in orders:
        print(f"\n--- Order ID: {order['id']} Source: {order['source']} ---")
        
        # Check billing
        try:
            billing = json.loads(order['billing'])
            print("Billing Phone:", billing.get('phone'))
            print("Billing Data:", json.dumps(billing, indent=2))
        except:
            print("Billing: Raw/Error", order['billing'])
            
        # Check shipping
        try:
            shipping = json.loads(order['shipping'])
            print("Shipping Phone:", shipping.get('phone')) # Shipping usually doesn't have phone in WC, but maybe?
        except:
            print("Shipping: Raw/Error", order['shipping'])

        # Check meta_data
        try:
            meta_data = json.loads(order['meta_data'])
            # Look for any value resembling the phone number 48796223605
            print("Scanning meta_data for phone number...")
            found_in_meta = False
            for meta in meta_data:
                val = str(meta.get('value', ''))
                if '48796223605' in val or '796223605' in val:
                    print(f"Found in meta key '{meta.get('key')}': {val}")
                    found_in_meta = True
            
            if not found_in_meta:
                print("Phone number not found in meta_data.")
                
        except:
            print("Meta Data: Raw/Error")

    conn.close()

if __name__ == "__main__":
    inspect_order('44945')
