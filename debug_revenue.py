#!/usr/bin/env python3
"""Debug script to compare revenue calculations between Dashboard and Products page"""

import sqlite3
import json
from datetime import date

def get_db_connection():
    conn = sqlite3.connect('/www/wwwroot/woo-analysis/woocommerce_orders.db')
    conn.row_factory = sqlite3.Row
    return conn

def parse_json_field(field):
    if not field:
        return field
    if isinstance(field, str):
        try:
            return json.loads(field)
        except (json.JSONDecodeError, TypeError):
            return field
    return field

def main():
    conn = get_db_connection()
    
    # This month filter
    today = date.today()
    date_from = today.replace(day=1).isoformat()
    date_to = today.isoformat() + 'T23:59:59'
    
    # Dashboard calculation: order-level net revenue for successful orders
    dashboard_query = """
        SELECT currency, SUM(total) as revenue, SUM(shipping_total) as shipping
        FROM orders
        WHERE date_created >= ? AND date_created <= ?
        AND status NOT IN ('failed', 'cancelled')
        GROUP BY currency
    """
    
    dashboard_results = conn.execute(dashboard_query, (date_from, date_to)).fetchall()
    
    print("=" * 60)
    print("DASHBOARD CALCULATION (Order Level)")
    print("=" * 60)
    
    total_dashboard_net = 0
    for row in dashboard_results:
        revenue = row['revenue'] or 0
        shipping = row['shipping'] or 0
        net = revenue - shipping
        total_dashboard_net += net
        print(f"  {row['currency']}: {revenue:.2f} total - {shipping:.2f} shipping = {net:.2f} net")
    
    print(f"\nTotal Net Revenue: {total_dashboard_net:.2f}")
    
    # Products calculation: item-level totals
    print("\n" + "=" * 60)
    print("PRODUCTS CALCULATION (Item Level)")
    print("=" * 60)
    
    orders_query = """
        SELECT id, line_items, currency, total as order_total, shipping_total
        FROM orders
        WHERE date_created >= ? AND date_created <= ?
        AND status NOT IN ('failed', 'cancelled')
    """
    
    orders = conn.execute(orders_query, (date_from, date_to)).fetchall()
    
    item_totals_by_currency = {}
    gross_totals_by_currency = {}
    order_totals_by_currency = {}
    
    for order in orders:
        currency = order['currency'] or 'N/A'
        items = parse_json_field(order['line_items'])
        order_shipping = float(order['shipping_total'] or 0)
        
        if not isinstance(items, list):
            continue
        
        # Calculate item-level totals with discount ratio
        items_sum = 0
        for item in items:
            item_total = float(item.get('total', 0))
            items_sum += item_total
        
        # Calculate discount ratio
        expected_items_value = float(order['order_total'] or 0) - order_shipping
        if items_sum > 0:
            discount_ratio = expected_items_value / items_sum
        else:
            discount_ratio = 1.0
        
        # Pro-rate shipping
        for item in items:
            item_total = float(item.get('total', 0))
            # Apply discount
            actual_item_revenue = item_total * discount_ratio
            
            if items_sum > 0:
                item_shipping = order_shipping * (item_total / items_sum)
            else:
                item_shipping = 0
            
            gross_total = actual_item_revenue + item_shipping
            
            if currency not in item_totals_by_currency:
                item_totals_by_currency[currency] = 0
                gross_totals_by_currency[currency] = 0
            
            item_totals_by_currency[currency] += actual_item_revenue
            gross_totals_by_currency[currency] += gross_total
        
        # Track order-level totals for comparison
        if currency not in order_totals_by_currency:
            order_totals_by_currency[currency] = 0
        order_totals_by_currency[currency] += float(order['order_total'] or 0)
    
    total_product_net = 0
    total_product_gross = 0
    
    for currency, amount in item_totals_by_currency.items():
        gross = gross_totals_by_currency.get(currency, 0)
        order_total = order_totals_by_currency.get(currency, 0)
        total_product_net += amount
        total_product_gross += gross
        
        print(f"  {currency}:")
        print(f"    Item totals (net): {amount:.2f}")
        print(f"    + Pro-rated shipping (gross): {gross:.2f}")
        print(f"    Order totals (reference): {order_total:.2f}")
        
        # Calculate difference
        diff = order_total - gross
        if abs(diff) > 0.01:
            print(f"    ⚠️  Difference (order - gross): {diff:.2f}")
    
    print(f"\nTotal Net Revenue (item totals): {total_product_net:.2f}")
    print(f"Total Gross Revenue (w/ shipping): {total_product_gross:.2f}")
    
    # Summary comparison
    print("\n" + "=" * 60)
    print("COMPARISON SUMMARY")
    print("=" * 60)
    print(f"Dashboard Net Revenue:  {total_dashboard_net:.2f}")
    print(f"Products Net Revenue:   {total_product_net:.2f}")
    print(f"Products Gross Revenue: {total_product_gross:.2f}")
    
    difference = total_dashboard_net - total_product_net
    if abs(difference) > 1:
        print(f"\n⚠️  Net Revenue Difference: {difference:.2f}")
        print("   This is likely due to order-level discounts not reflected in item totals")
    else:
        print(f"\n✅ Net Revenue calculations match!")
    
    conn.close()

if __name__ == '__main__':
    main()
