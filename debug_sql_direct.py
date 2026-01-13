import sqlite3
import os

db_path = '/www/wwwroot/woo-analysis/woocommerce_orders.db'

try:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    # Query with new field
    query = '''
        SELECT o.id, o.number, o.status, o.total, o.currency, o.date_created, o.date_modified,
               o.source, o.billing, o.shipping, o.line_items, o.meta_data, o.shipping_lines,
               s.manager,
               sl.tracking_number, sl.carrier_slug, sl.shipped_at
        FROM orders o
        LEFT JOIN sites s ON o.source = s.url
        LEFT JOIN shipping_logs sl ON o.id = sl.order_id
        WHERE o.status = 'on-hold' LIMIT 1
    '''
    cursor = conn.execute(query)
    row = cursor.fetchone()
    if row:
        print("Query successful")
        d = dict(row)
        print("Row dict created successfully")
        print("shipping_lines:", d['shipping_lines'][:50] if d['shipping_lines'] else 'None')
    else:
        print("No shipped orders found")
except Exception as e:
    print("Error:", e)
