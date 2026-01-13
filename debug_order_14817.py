
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
            print("Billing:", json.dumps(billing, indent=2, ensure_ascii=False))
        except:
            print("Billing: Raw/Error")
            
        # Check shipping
        try:
            shipping = json.loads(order['shipping'])
            print("Shipping:", json.dumps(shipping, indent=2, ensure_ascii=False))
        except:
            print("Shipping: Raw/Error")

        # Check meta_data
        try:
            meta_data = json.loads(order['meta_data'])
            print("Meta Data (Keys only or specific values):")
            for meta in meta_data:
                key = meta.get('key')
                value = meta.get('value')
                # Print if key looks like address related or if value matches "Bieniądzice"
                if any(x in str(key).lower() for x in ['adres', 'dpd', 'domu', 'kod', 'pocztowy', 'miejscowość', 'shipping', 'billing']):
                    print(f"  {key}: {value}")
                elif any(x in str(value) for x in ['Bieniądzice', '98-300', 'Wieluń', '8', '661517307']):
                    print(f"  {key}: {value}")
                
        except:
            print("Meta Data: Raw/Error")

    conn.close()

if __name__ == "__main__":
    inspect_order('14817')
