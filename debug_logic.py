import sqlite3
import json
import traceback
from datetime import datetime

db_path = '/www/wwwroot/woo-analysis/woocommerce_orders.db'

def parse_json_field(field_value):
    if not field_value:
        return []
    try:
        if isinstance(field_value, str):
            return json.loads(field_value)
        return field_value
    except:
        return []

try:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    
    query = '''
        SELECT o.id, o.number, o.status, o.total, o.currency, o.date_created, o.date_modified,
               o.source, o.billing, o.shipping, o.line_items, o.meta_data, o.shipping_lines,
               s.manager,
               sl.tracking_number, sl.carrier_slug, sl.shipped_at
        FROM orders o
        LEFT JOIN sites s ON o.source = s.url
        LEFT JOIN shipping_logs sl ON o.id = sl.order_id
        WHERE o.status = 'on-hold'
    '''
    orders = conn.execute(query).fetchall()
    print(f"Found {len(orders)} orders")

    for row in orders:
        order = dict(row)
        # Verify shipping_lines access
        # print(f"Order {order['number']} shipping_lines type: {type(order['shipping_lines'])}")
        
        tracking_number = order['tracking_number']
        carrier_slug = order['carrier_slug']
        
        meta_data = parse_json_field(order['meta_data'])

        # If still no tracking, try meta_data for Advanced Shipment Tracking Pro
        if not tracking_number and meta_data:
             for meta in meta_data:
                # ... existing logic skipped for brevity ...
                pass

        # If still no tracking, try line_items meta_data
        if not tracking_number:
            line_items = parse_json_field(order['line_items'])
            if line_items:
                for item in line_items:
                    if isinstance(item, dict):
                        item_meta = item.get('meta_data', [])
                        for meta in item_meta:
                            # VillaTheme skipped
                            
                            # Simple 'tracking_number' key (Custom implementation)
                            if isinstance(meta, dict) and meta.get('key') == 'tracking_number':
                                tracking_number = meta.get('value', '').strip()
                                
                    if tracking_number:
                        break
        
        # If still no tracking, try shipping_lines meta_data 'tracking_number'
        shipping_lines = parse_json_field(order['shipping_lines'])
        if not tracking_number and shipping_lines:
            for item in shipping_lines:
                if isinstance(item, dict):
                    for meta in item.get('meta_data', []):
                        if isinstance(meta, dict) and meta.get('key') == 'tracking_number':
                            tracking_number = meta.get('value', '').strip()
                            if tracking_number:
                                break
                if tracking_number:
                    break

        # If still no tracking, try order meta_data '_tracking_number'
        if not tracking_number and meta_data:
             for meta in meta_data:
                if isinstance(meta, dict) and meta.get('key') == '_tracking_number':
                    tracking_number = meta.get('value', '').strip()
                    if tracking_number:
                        break
        
        if tracking_number:
            print(f"Order {order['number']}: Found tracking {tracking_number}")

    print("Finished processing")

except Exception as e:
    traceback.print_exc()
