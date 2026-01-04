
import sqlite3

db_file = 'woocommerce_orders.db'
conn = sqlite3.connect(db_file)
cursor = conn.cursor()

print("--- SITES ---")
cursor.execute("SELECT id, url FROM sites WHERE url LIKE '%vapeprimeau.com%'")
sites = cursor.fetchall()
for s in sites:
    print(f"ID: {s[0]}, URL repr: {repr(s[1])}")

print("\n--- ORDERS ---")
# Check source for the specific order
cursor.execute("SELECT id, source FROM orders WHERE id='30898' OR number='30898'")
orders = cursor.fetchall()
for o in orders:
    print(f"Order ID: {o[0]}, Source repr: {repr(o[1])}")

conn.close()
