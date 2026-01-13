from app import app, get_db_connection
import sqlite3

try:
    with app.app_context():
        conn = get_db_connection()
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
        row = conn.execute(query).fetchone()
        if row:
            print("Query successful")
            # Iterate to verify all fields access
            d = dict(row)
            print("Row dict created successfully")
            if 'shipping_lines' in d:
                print("shipping_lines present:", bool(d['shipping_lines']))
        else:
            print("No shipped orders found")
except Exception as e:
    print("Error:", e)
