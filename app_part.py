@app.route('/api/shipping/print/pending')
@login_required
@shipper_required
def print_pending_list():
    """Get all pending orders for printing"""
    conn = get_db_connection()
    
    query = '''
        SELECT o.id, o.number, o.total, o.currency, o.source, o.billing, o.shipping, o.line_items,
               s.manager
        FROM orders o
        LEFT JOIN sites s ON o.source = s.url
        WHERE o.status = 'processing'
    '''
    params = []
    
    allowed_sources = get_user_allowed_sources(current_user.id, current_user.is_admin(), current_user.is_viewer())
    if allowed_sources is not None:
        placeholders = ','.join(['?' for _ in allowed_sources])
        query += f' AND o.source IN ({placeholders})'
        params.extend(allowed_sources)
    
    query += ' ORDER BY o.date_created'
    
    orders = conn.execute(query, params).fetchall()
    conn.close()
    
    result = []
    for order in orders:
        billing = parse_json_field(order['billing'])
        shipping_info = parse_json_field(order['shipping'])
        line_items = parse_json_field(order['line_items'])
        
        addr = shipping_info if shipping_info and shipping_info.get('address_1') else billing
        
        result.append({
            'order_number': order['number'],
            'source': order['source'].replace('https://www.', '').replace('https://', ''),
            'manager': order['manager'] or '',
            'customer_name': f"{addr.get('first_name', '')} {addr.get('last_name', '')}".strip(),
            'total': f"{float(order['total'] or 0):.2f} {order['currency']}",
            'products': [{'name': item.get('name', ''), 'qty': item.get('quantity', 1)} for item in (line_items or [])]
        })
    
    from datetime import datetime
    today = datetime.now().strftime('%Y-%m-%d')
