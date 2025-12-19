import sqlite3

def init_db():
    conn = sqlite3.connect('woocommerce_orders.db')
    c = conn.cursor()
    
    # Create customer_settings table
    c.execute('''
        CREATE TABLE IF NOT EXISTS customer_settings (
            email TEXT PRIMARY KEY,
            quality_tier TEXT,
            note TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    conn.commit()
    conn.close()
    print("Database initialized successfully.")

if __name__ == '__main__':
    init_db()
