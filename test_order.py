import sqlite3
import json

conn = sqlite3.connect('woocommerce_orders.db')
conn.row_factory = sqlite3.Row
order = conn.execute("SELECT * FROM orders WHERE number = '10'").fetchone()

if order:
    print(json.dumps(dict(order), indent=2))
else:
    print("Order not found")
