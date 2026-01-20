import sqlite3
import pandas as pd
from datetime import date

DB_FILE = 'woocommerce_orders.db'

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def verify_dashboard_stats():
    print("--- Verifying Dashboard Stats ---")
    conn = get_db_connection()
    
    # 1. Total Orders count should exclude checkout-draft
    # First get raw count of checkout-draft
    draft_count = conn.execute("SELECT COUNT(*) FROM orders WHERE status IN ('checkout-draft', 'trash')").fetchone()[0]
    print(f"Total draft/trash orders in DB: {draft_count}")
    
    # Simulate dashboard query condition (simplified)
    today = date.today()
    date_from = today.replace(day=1).isoformat()
    
    # Get total orders matching dashboard logic
    dashboard_total = conn.execute(f"SELECT COUNT(*) FROM orders WHERE date_created >= ? AND status NOT IN ('checkout-draft', 'trash')", (date_from,)).fetchone()[0]
    
    # Get total orders WITHOUT exclusion
    raw_total = conn.execute(f"SELECT COUNT(*) FROM orders WHERE date_created >= ?", (date_from,)).fetchone()[0]
    
    print(f"Dashboard Total Orders (filtered): {dashboard_total}")
    print(f"Raw Total Orders (unfiltered): {raw_total}")
    
    diff = raw_total - dashboard_total
    print(f"Difference (should match drafts in this period): {diff}")
    
    # Verify drafted orders in this period
    drafts_in_period = conn.execute(f"SELECT COUNT(*) FROM orders WHERE date_created >= ? AND status IN ('checkout-draft', 'trash')", (date_from,)).fetchone()[0]
    print(f"Drafts in period: {drafts_in_period}")
    
    if diff == drafts_in_period:
        print("✅ Dashboard total orders verification PASSED")
    else:
        print("❌ Dashboard total orders verification FAILED")
        
    conn.close()

def verify_products_stats():
    print("\n--- Verifying Products Stats ---")
    conn = get_db_connection()
    
    # Products page condition
    conditions = ["status NOT IN ('failed', 'cancelled', 'checkout-draft', 'trash', 'cheat')"]
    where_clause = 'WHERE ' + ' AND '.join(conditions)
    
    # Count orders considered for products
    considered_orders = conn.execute(f"SELECT COUNT(*) FROM orders {where_clause}").fetchone()[0]
    
    # Count total excluding only failed/cancelled (old logic)
    old_conditions = ["status NOT IN ('failed', 'cancelled')"]
    old_where_clause = 'WHERE ' + ' AND '.join(old_conditions)
    old_considered_orders = conn.execute(f"SELECT COUNT(*) FROM orders {old_where_clause}").fetchone()[0]
    
    print(f"Product Analysis Considered Orders: {considered_orders}")
    print(f"Old Logic Considered Orders: {old_considered_orders}")
    
    diff = old_considered_orders - considered_orders
    print(f"Difference: {diff}")
    
    excluded_count = conn.execute("SELECT COUNT(*) FROM orders WHERE status IN ('checkout-draft', 'trash', 'cheat')").fetchone()[0]
    print(f"Excluded orders (draft/trash/cheat) in DB: {excluded_count}")
    
    if diff == excluded_count:
         print("✅ Products stats verification PASSED")
    else:
         print("❌ Products stats verification FAILED")

    conn.close()

if __name__ == "__main__":
    verify_dashboard_stats()
    verify_products_stats()
