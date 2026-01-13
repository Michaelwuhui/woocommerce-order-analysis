
import sqlite3
import json

DB_FILE = '/www/wwwroot/woo-analysis/woocommerce_orders.db'

def inspect_customer_orders(email):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    print(f"Searching for orders with email: {email}")
    cursor.execute("SELECT id, number, date_created, billing FROM orders WHERE billing LIKE ? ORDER BY date_created DESC", (f'%"{email}"%',))
    orders = cursor.fetchall()
    
    if not orders:
        print("No orders found for this email.")
        return

    print(f"Found {len(orders)} orders.")
    
    for order in orders:
        try:
            billing = json.loads(order['billing'])
            phone = billing.get('phone', 'N/A')
            print(f"Order #{order['number']} ({order['date_created']}): Phone='{phone}'")
        except:
            print(f"Order #{order['number']}: Error parsing billing")

    conn.close()

if __name__ == "__main__":
    inspect_customer_orders('danikowalewski12@gmail.com')
