"""
WooCommerce Order Analysis Web Dashboard
Flask application with user authentication and data visualization
"""
import sqlite3
import json
from datetime import datetime
from functools import wraps

from flask import Flask, render_template, render_template_string, request, redirect, url_for, flash, jsonify, session
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd

app = Flask(__name__)
app.secret_key = 'woocommerce-order-analysis-secret-key-2024'

# Flask-Login setup
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = '请先登录以访问此页面'

# Database configuration
DB_FILE = 'woocommerce_orders.db'

# Simple user storage (in production, use a proper database)
USERS = {
    'admin': {
        'id': '1',
        'username': 'admin',
        'password': generate_password_hash('admin123'),
        'name': '管理员'
    }
}


class User(UserMixin):
    def __init__(self, user_id, username, name, role='user'):
        self.id = user_id
        self.username = username
        self.name = name
        self.role = role
    
    def is_admin(self):
        return self.role == 'admin'


@login_manager.user_loader
def load_user(user_id):
    conn = get_db_connection()
    user_row = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
    conn.close()
    if user_row:
        return User(user_row['id'], user_row['username'], user_row['name'], user_row['role'])
    return None


def get_db_connection():
    """Create database connection"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def get_all_managers():
    """Get list of all unique site managers"""
    conn = get_db_connection()
    managers = conn.execute('SELECT DISTINCT manager FROM sites WHERE manager IS NOT NULL AND manager != "" ORDER BY manager').fetchall()
    conn.close()
    return [m['manager'] for m in managers]


def parse_json_field(value):
    """Safely parse JSON field"""
    if not value:
        return {}
    try:
        return json.loads(value)
    except:
        return {}


def extract_flavor_from_meta(item):
    """
    Extract flavor/variation attribute from WooCommerce line_item meta_data.
    Looks for common variation attribute keys like 'flavour', 'flavor', 'pa_flavour', 'pa_flavor', etc.
    Returns the display_value if found, otherwise empty string.
    """
    meta_data = item.get('meta_data', [])
    if not isinstance(meta_data, list):
        return ''
    
    # Common flavor-related keys in WooCommerce (also include Polish 'smak/smaki')
    flavor_keys = ['pa_flavour', 'pa_flavor', 'flavour', 'flavor', 'pa_taste', 'taste', 'pa_variant', 'variant', 'pa_smak', 'smak', 'pa_smaki', 'smaki']
    
    for meta in meta_data:
        if not isinstance(meta, dict):
            continue
        key = meta.get('key', '').lower()
        if key in flavor_keys:
            # Prefer display_value over value for human-readable format
            return meta.get('display_value', '') or meta.get('value', '')
    
    return ''


def extract_puffs_from_meta(item):
    """
    Extract puffs count from WooCommerce line_item meta_data.
    Looks for 'puffs', 'pa_puffs', 'puff_count' keys.
    Returns the numeric puffs value if found, otherwise None.
    """
    import re
    
    meta_data = item.get('meta_data', [])
    if not isinstance(meta_data, list):
        return None
    
    # Common puffs-related keys in WooCommerce (including Polish 'liczba-zaciagniec')
    puffs_keys = ['pa_puffs', 'puffs', 'puff_count', 'pa_puff_count', 'pa_liczba-zaciagniec', 'liczba-zaciagniec']
    
    for meta in meta_data:
        if not isinstance(meta, dict):
            continue
        key = meta.get('key', '').lower()
        if key in puffs_keys:
            value = meta.get('display_value', '') or meta.get('value', '')
            # Extract numeric value from strings like "15000 puffs" or "15000"
            if value:
                match = re.search(r'(\d+)', str(value))
                if match:
                    return int(match.group(1))
    
    return None


def calculate_customer_tier(successful_orders, total_spending, avg_days_between):
    """Calculate customer tier based on orders, spending, and frequency"""
    quality_score = 0
    quality_score += min(successful_orders * 10, 30)  # Max 30 for orders
    quality_score += min(total_spending / 100, 40)    # Max 40 for spending
    
    # Frequency bonus
    if avg_days_between > 0:
        if avg_days_between < 60:
            quality_score += 30
        elif avg_days_between < 120:
            quality_score += 15
    
    quality_score = min(quality_score, 100)
    
    if quality_score >= 80:
        return 'vip'
    elif quality_score >= 60:
        return 'good'
    elif quality_score >= 40:
        return 'normal'
    else:
        return 'new'


def get_full_product_name(item):
    """
    Get full product name including variation attributes (like flavor).
    Combines the product name with any variation flavor found in meta_data.
    Returns: (full_name, flavor_only, puffs_from_meta)
    """
    name = item.get('name', '')
    flavor = extract_flavor_from_meta(item)
    puffs = extract_puffs_from_meta(item)
    
    if flavor:
        # If flavor is not already in the name, append it
        if flavor.upper() not in name.upper():
            full_name = f"{name} - {flavor}"
        else:
            full_name = name
    else:
        full_name = name
        flavor = ''
    
    return full_name, flavor, puffs


def get_user_allowed_sources(user_id, is_admin=False):
    """
    Get list of source URLs that a user has permission to access.
    Admin users can access all sources.
    Returns None if user has no restrictions (can see all), or a list of allowed URLs.
    """
    if is_admin:
        return None  # No restrictions for admin
    
    conn = get_db_connection()
    
    # Get site_ids from user_site_permissions
    perms = conn.execute('''
        SELECT s.url FROM user_site_permissions p
        JOIN sites s ON p.site_id = s.id
        WHERE p.user_id = ?
    ''', (user_id,)).fetchall()
    
    conn.close()
    
    if not perms:
        # No permissions set - return empty list (no access)
        return []
    
    return [p['url'] for p in perms]


def build_source_filter_clause(allowed_sources, table_alias=''):
    """
    Build SQL WHERE clause for filtering by allowed sources.
    Returns (clause_string, params_list)
    """
    if allowed_sources is None:
        # No restrictions
        return '', []
    
    if not allowed_sources:
        # Empty list - no access to anything
        return 'AND 1=0', []  # Always false condition
    
    prefix = f'{table_alias}.' if table_alias else ''
    placeholders = ', '.join(['?' for _ in allowed_sources])
    return f'AND {prefix}source IN ({placeholders})', allowed_sources


def get_cny_rate(currency, year_month):
    """
    Get CNY exchange rate for a currency in a specific month.
    Falls back to the most recent rate if not found for the specific month.
    Returns (rate, actual_year_month) or (None, None) if not found.
    """
    if not currency or currency == 'CNY':
        return 1.0, year_month  # CNY to CNY is 1:1
    
    conn = get_db_connection()
    
    # Try exact month first
    rate = conn.execute('''
        SELECT rate_to_cny, year_month FROM exchange_rates
        WHERE currency = ? AND year_month = ?
    ''', (currency, year_month)).fetchone()
    
    if rate:
        conn.close()
        return rate['rate_to_cny'], rate['year_month']
    
    # Fall back to most recent rate before this month
    rate = conn.execute('''
        SELECT rate_to_cny, year_month FROM exchange_rates
        WHERE currency = ? AND year_month <= ?
        ORDER BY year_month DESC
        LIMIT 1
    ''', (currency, year_month)).fetchone()
    
    if rate:
        conn.close()
        return rate['rate_to_cny'], rate['year_month']
    
    # Fall back to any rate for this currency
    rate = conn.execute('''
        SELECT rate_to_cny, year_month FROM exchange_rates
        WHERE currency = ?
        ORDER BY year_month DESC
        LIMIT 1
    ''', (currency,)).fetchone()
    
    conn.close()
    
    if rate:
        return rate['rate_to_cny'], rate['year_month']
    
    return None, None


def convert_to_cny(amount, currency, year_month):
    """
    Convert an amount to CNY using the exchange rate for the given month.
    Returns (cny_amount, rate, actual_month) or (None, None, None) if no rate found.
    """
    if amount is None:
        return None, None, None
    
    rate, actual_month = get_cny_rate(currency, year_month)
    if rate is None:
        return None, None, None
    
    return round(amount * rate, 2), rate, actual_month


def get_all_exchange_rates():
    """Get all exchange rates from database."""
    conn = get_db_connection()
    rates = conn.execute('''
        SELECT id, year_month, currency, rate_to_cny, updated_at
        FROM exchange_rates
        ORDER BY year_month DESC, currency
    ''').fetchall()
    conn.close()
    return [dict(r) for r in rates]


def admin_required(f):
    """Decorator to require admin role"""
    from functools import wraps
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or not current_user.is_admin():
            # Return a beautiful access denied page
            return render_template_string('''
<!DOCTYPE html>
<html lang="zh">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>访问受限 - WooCommerce 订单分析</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            color: #e0e0e0;
        }
        .container {
            text-align: center;
            padding: 60px 40px;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 24px;
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            max-width: 480px;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
        }
        .icon {
            width: 120px;
            height: 120px;
            margin: 0 auto 30px;
            background: linear-gradient(135deg, rgba(239, 68, 68, 0.2), rgba(239, 68, 68, 0.1));
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            border: 2px solid rgba(239, 68, 68, 0.3);
        }
        .icon i { font-size: 48px; color: #ef4444; }
        h1 {
            font-size: 28px;
            font-weight: 600;
            margin-bottom: 16px;
            background: linear-gradient(135deg, #f87171, #ef4444);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        p { color: #9ca3af; font-size: 16px; line-height: 1.6; margin-bottom: 30px; }
        .user-info {
            background: rgba(0, 0, 0, 0.2);
            padding: 12px 20px;
            border-radius: 12px;
            margin-bottom: 30px;
            font-size: 14px;
        }
        .user-info span { color: #60a5fa; }
        .btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 14px 32px;
            background: linear-gradient(135deg, #3b82f6, #2563eb);
            color: white;
            text-decoration: none;
            border-radius: 12px;
            font-weight: 500;
            transition: all 0.3s ease;
        }
        .btn:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(59, 130, 246, 0.3); }
    </style>
</head>
<body>
    <div class="container">
        <div class="icon"><i class="bi bi-shield-lock"></i></div>
        <h1>访问受限</h1>
        <p>抱歉，您没有权限访问此页面。<br>此功能仅限管理员使用。</p>
        <div class="user-info">当前登录: <span>{{ current_user.name or current_user.username }}</span></div>
        <a href="/" class="btn"><i class="bi bi-house"></i> 返回首页</a>
    </div>
</body>
</html>
            '''), 403
        return f(*args, **kwargs)
    return decorated_function


# ============== ROUTES ==============

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        
        conn = get_db_connection()
        user_row = conn.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
        conn.close()
        
        if user_row and check_password_hash(user_row['password_hash'], password):
            user = User(user_row['id'], username, user_row['name'], user_row['role'])
            login_user(user)
            flash('登录成功！', 'success')
            next_page = request.args.get('next')
            return redirect(next_page or url_for('dashboard'))
        else:
            flash('用户名或密码错误', 'error')
    
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('您已成功退出登录', 'info')
    return redirect(url_for('login'))


@app.route('/')
@login_required
def dashboard():
    """Main dashboard with statistics"""
    from datetime import date, timedelta
    conn = get_db_connection()
    
    # Get user's allowed sources for permission filtering
    allowed_sources = get_user_allowed_sources(current_user.id, current_user.is_admin())
    source_clause, source_params = build_source_filter_clause(allowed_sources)
    
    # Get filters
    quick_date = request.args.get('quick_date', 'this_month')
    source_filter = request.args.get('source', '')
    manager_filter = request.args.get('manager', '')
    
    # Get all managers
    all_managers = get_all_managers()
    
    # Validate source_filter against allowed sources
    if source_filter and allowed_sources is not None and source_filter not in allowed_sources:
        source_filter = ''  # Reset invalid filter
    
    # Get available sources (filtered by permissions and manager)
    conn = get_db_connection()
    
    # Base source query
    source_query = 'SELECT DISTINCT source FROM orders'
    source_params = []
    source_conditions = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(allowed_sources)
        else:
            source_conditions.append('1=0')
            
    if manager_filter:
        # Get sites managed by this manager
        manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
        manager_urls = [s['url'] for s in manager_sites]
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(manager_urls)
        else:
            source_conditions.append('1=0')
            
    if source_conditions:
        source_query += ' WHERE ' + ' AND '.join(source_conditions)
        
    source_query += ' ORDER BY source'
    sources = conn.execute(source_query, source_params).fetchall()
    
    # Process quick date filter
    today = date.today()
    date_from = ''
    date_to = today.isoformat()
    
    if quick_date == 'this_month':
        date_from = today.replace(day=1).isoformat()
    elif quick_date == 'last_month':
        first_of_this_month = today.replace(day=1)
        last_month_end = first_of_this_month - timedelta(days=1)
        last_month_start = last_month_end.replace(day=1)
        date_from = last_month_start.isoformat()
        date_to = last_month_end.isoformat()
    elif quick_date == 'last_quarter':
        date_from = (today - timedelta(days=90)).isoformat()
    elif quick_date == 'half_year':
        date_from = (today - timedelta(days=180)).isoformat()
    elif quick_date == 'one_year':
        date_from = (today - timedelta(days=365)).isoformat()
    elif quick_date == 'all':
        date_from = ''
        date_to = ''
    
    # Build filter conditions with permission check
    conditions = ['1=1']  # Base condition
    params = []
    
    # Add permission filter
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            conditions.append(f'source IN ({placeholders})')
            params.extend(allowed_sources)
        else:
            conditions.append('1=0')  # No access
            
    # Add manager filter
    if manager_filter:
        # Re-use manager_urls from above or fetch if not done
        if 'manager_urls' not in locals():
            manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
            manager_urls = [s['url'] for s in manager_sites]
        
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            conditions.append(f'source IN ({placeholders})')
            params.extend(manager_urls)
        else:
            conditions.append('1=0')
    
    if date_from and date_to:
        conditions.append(f"date_created >= ? AND date_created <= ?")
        params.extend([date_from, date_to + 'T23:59:59'])
    elif date_from:
        conditions.append(f"date_created >= ?")
        params.append(date_from)
    if source_filter:
        conditions.append("source = ?")
        params.append(source_filter)
    
    date_condition = 'WHERE ' + ' AND '.join(conditions)
    
    # Get overall statistics
    stats = {}
    
    # Total orders
    stats['total_orders'] = conn.execute(f'SELECT COUNT(*) FROM orders {date_condition}', params).fetchone()[0]
    
    # Total revenue by currency - add status filter
    revenue_conditions = conditions.copy()
    revenue_conditions.append('status NOT IN ("failed", "cancelled")')
    revenue_where = 'WHERE ' + ' AND '.join(revenue_conditions)
    
    # Valid orders count (for AOV calculation)
    stats['valid_orders'] = conn.execute(f'SELECT COUNT(*) FROM orders {revenue_where}', params).fetchone()[0]
    
    # Get revenue grouped by currency (total and net = total - shipping)
    revenue_by_currency_raw = conn.execute(f'''
        SELECT currency, SUM(total) as revenue, SUM(shipping_total) as shipping
        FROM orders {revenue_where}
        GROUP BY currency
    ''', params).fetchall()
    stats['total_revenue_by_currency'] = {row['currency']: row['revenue'] or 0 for row in revenue_by_currency_raw}
    stats['net_revenue_by_currency'] = {row['currency']: (row['revenue'] or 0) - (row['shipping'] or 0) for row in revenue_by_currency_raw}
    
    # Also keep a simple total for backward compatibility
    stats['total_revenue'] = sum(stats['total_revenue_by_currency'].values())
    stats['net_revenue'] = sum(stats['net_revenue_by_currency'].values())
    
    # Orders by status with currency
    status_data_raw = conn.execute(f'''
        SELECT status, currency, COUNT(*) as count, SUM(total) as revenue
        FROM orders {date_condition}
        GROUP BY status, currency
        ORDER BY count DESC
    ''', params).fetchall()
    # Group by status, with currency breakdown
    status_dict = {}
    for row in status_data_raw:
        status = row['status']
        if status not in status_dict:
            status_dict[status] = {'status': status, 'count': 0, 'revenue_by_currency': {}}
        status_dict[status]['count'] += row['count']
        currency = row['currency'] or 'N/A'
        if currency not in status_dict[status]['revenue_by_currency']:
            status_dict[status]['revenue_by_currency'][currency] = 0
        status_dict[status]['revenue_by_currency'][currency] += row['revenue'] or 0
    status_data = sorted(status_dict.values(), key=lambda x: x['count'], reverse=True)
    
    # Orders by source with currency
    source_data_raw = conn.execute(f'''
        SELECT source, currency, COUNT(*) as count, SUM(total) as revenue, SUM(shipping_total) as shipping
        FROM orders {date_condition}
        GROUP BY source, currency
    ''', params).fetchall()
    # Group by source
    source_dict = {}
    for row in source_data_raw:
        source = row['source']
        if source not in source_dict:
            source_dict[source] = {'source': source, 'count': 0, 'currency': row['currency'], 'revenue': 0, 'shipping': 0}
        source_dict[source]['count'] += row['count']
        source_dict[source]['revenue'] += row['revenue'] or 0
        source_dict[source]['shipping'] += row['shipping'] or 0
    # Calculate net revenue
    for source in source_dict.values():
        source['net_revenue'] = source['revenue'] - source['shipping']
    source_data = list(source_dict.values())
    
    # Recent orders - apply permission filter, optionally source filter
    recent_conditions = ['1=1']
    recent_params = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            recent_conditions.append(f'source IN ({placeholders})')
            recent_params.extend(allowed_sources)
        else:
            recent_conditions.append('1=0')
    
    if source_filter:
        recent_conditions.append('source = ?')
        recent_params.append(source_filter)

    if manager_filter:
        manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
        manager_urls = [s['url'] for s in manager_sites]
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            recent_conditions.append(f'source IN ({placeholders})')
            recent_params.extend(manager_urls)
        else:
            recent_conditions.append('1=0')
    
    recent_where = 'WHERE ' + ' AND '.join(recent_conditions)
    recent_orders = conn.execute(f'''
        SELECT id, number, status, total, shipping_total, currency, date_created, source, line_items, billing
        FROM orders
        {recent_where}
        ORDER BY date_created DESC
        LIMIT 10
    ''', recent_params).fetchall()
    
    # Process recent orders to get product count
    processed_orders = []
    for order in recent_orders:
        order_dict = dict(order)
        items = parse_json_field(order['line_items'])
        order_dict['product_count'] = sum(item.get('quantity', 0) for item in items) if isinstance(items, list) else 0
        
        # Parse customer info from billing
        billing = parse_json_field(order['billing'])
        order_dict['customer_name'] = f"{billing.get('first_name', '')} {billing.get('last_name', '')}".strip() if billing else ''
        order_dict['customer_email'] = billing.get('email', '') if billing else ''
        
        processed_orders.append(order_dict)
    
    # Trend data based on filter - daily for single month, monthly for longer periods
    trend_type = 'daily' if quick_date in ['this_month', 'last_month'] else 'monthly'
    
    if trend_type == 'daily':
        # Daily trend for this_month or last_month
        trend_data_raw = conn.execute(f'''
            SELECT strftime('%Y-%m-%d', date_created) as period, 
                   COUNT(*) as orders,
                   SUM(total) as revenue
            FROM orders {date_condition}
            GROUP BY period
            ORDER BY period
        ''', params).fetchall()
    else:
        # Monthly trend - use same date_condition or default
        trend_data_raw = conn.execute(f'''
            SELECT strftime('%Y-%m', date_created) as period, 
                   COUNT(*) as orders,
                   SUM(total) as revenue
            FROM orders {date_condition}
            GROUP BY period
            ORDER BY period
        ''', params).fetchall()
    trend_data = [dict(row) for row in trend_data_raw]
    
    # Get site managers mapping
    sites = conn.execute('SELECT url, manager FROM sites').fetchall()
    site_managers = {s['url']: s['manager'] or '' for s in sites}
    
    # Calculate CNY totals for revenue
    total_cny = 0
    net_cny = 0
    for currency, amount in stats['total_revenue_by_currency'].items():
        rate, _ = get_cny_rate(currency, date_to[:7] if date_to else (date_from[:7] if date_from else None))
        if rate:
            total_cny += amount * rate
            net_cny += stats['net_revenue_by_currency'].get(currency, 0) * rate
    stats['total_revenue_cny'] = round(total_cny, 2)
    stats['net_revenue_cny'] = round(net_cny, 2)
    
    # Add CNY conversion to source_data
    for source in source_data:
        currency = source.get('currency', 'PLN')
        # Get month from date filter or use current month
        month = date_to[:7] if date_to else (date_from[:7] if date_from else None)
        rate, _ = get_cny_rate(currency, month)
        source['rate_to_cny'] = rate if rate else None
        source['revenue_cny'] = round(source['revenue'] * rate, 2) if rate else None
        source['net_revenue_cny'] = round(source['net_revenue'] * rate, 2) if rate else None
    
    # Add CNY conversion to recent_orders
    for order in processed_orders:
        currency = order.get('currency', 'PLN')
        order_month = order['date_created'][:7] if order.get('date_created') else None
        rate, _ = get_cny_rate(currency, order_month)
        order['rate_to_cny'] = rate if rate else None
        order['total_cny'] = round(float(order['total'] or 0) * rate, 2) if rate else None
        order['net_total'] = float(order['total'] or 0) - float(order.get('shipping_total') or 0)
        order['net_total_cny'] = round(order['net_total'] * rate, 2) if rate else None
    
    # Get customer attributes for the recent orders
    customer_emails = list(set(o['customer_email'] for o in processed_orders if o.get('customer_email')))
    customer_attributes = {}
    
    if customer_emails:
        placeholders = ', '.join(['?' for _ in customer_emails])
        
        # Get manual quality settings
        manual_settings = conn.execute(f'''
            SELECT email, quality_tier FROM customer_settings 
            WHERE email IN ({placeholders})
        ''', customer_emails).fetchall()
        
        # Calculate attributes for each customer
        for email in customer_emails:
            if email not in customer_attributes:
                customer_attributes[email] = {}
            
            # Check manual setting
            manual_tier = 'auto'
            for row in manual_settings:
                if row['email'] == email:
                    manual_tier = row['quality_tier']
                    break
            
            # Get customer stats
            stats_row = conn.execute('''
                SELECT 
                    COUNT(*) as total_orders,
                    SUM(CASE WHEN status IN ('completed', 'processing') THEN 1 ELSE 0 END) as successful_orders,
                    SUM(CASE WHEN status IN ('completed', 'processing') THEN total ELSE 0 END) as total_spending,
                    MAX(date_created) as last_order_date,
                    MIN(date_created) as first_order_date
                FROM orders 
                WHERE billing LIKE ?
            ''', (f'%"{email}"%',)).fetchone()
            
            total_orders = stats_row['total_orders'] or 0
            successful_orders = stats_row['successful_orders'] or 0
            total_spending = stats_row['total_spending'] or 0
            
            customer_attributes[email]['order_count'] = total_orders
            customer_attributes[email]['is_new'] = (total_orders <= 1)
            
            tier = manual_tier
            if tier == 'auto':
                avg_days_between = 0
                if total_orders > 1 and stats_row['first_order_date'] and stats_row['last_order_date']:
                    from datetime import datetime
                    try:
                        first_date = datetime.fromisoformat(stats_row['first_order_date'][:19])
                        last_date = datetime.fromisoformat(stats_row['last_order_date'][:19])
                        days_span = (last_date - first_date).days
                        avg_days_between = days_span / (total_orders - 1)
                    except:
                        avg_days_between = 0
                
                tier = calculate_customer_tier(successful_orders, total_spending, avg_days_between)
            
            if tier == 'vip':
                customer_attributes[email]['quality'] = {'label': 'VIP', 'class': 'text-warning', 'icon': 'star-fill'}
            elif tier == 'good':
                customer_attributes[email]['quality'] = {'label': '优质', 'class': 'text-success', 'icon': 'gem'}
            elif tier == 'normal':
                customer_attributes[email]['quality'] = {'label': '普通', 'class': 'text-primary', 'icon': 'person-check'}
            elif tier == 'new':
                customer_attributes[email]['quality'] = {'label': '新客', 'class': 'text-info', 'icon': 'stars'}
            elif tier == 'bad':
                customer_attributes[email]['quality'] = {'label': '劣质', 'class': 'text-danger', 'icon': 'x-circle'}

    conn.close()
    
    # Calculate overall CNY rate for dashboard
    cny_rate = 0
    if stats.get('net_revenue'):
        cny_rate = round(stats.get('net_revenue_cny', 0) / stats['net_revenue'], 4)
        
    return render_template('dashboard.html',
                         stats=stats,
                         status_data=status_data,
                         source_data=source_data,
                         recent_orders=processed_orders,
                         trend_data=trend_data,
                         trend_type=trend_type,
                         quick_date=quick_date,
                         sources=sources,
                         source_filter=source_filter,
                         manager_filter=manager_filter,
                         all_managers=all_managers,
                         site_managers=site_managers,
                         customer_attributes=customer_attributes,
                         cny_rate=cny_rate)


@app.route('/orders')
@login_required
def orders():
    """Order list with filtering and summary statistics"""
    from datetime import date, timedelta
    conn = get_db_connection()
    
    # Get user's allowed sources for permission filtering
    allowed_sources = get_user_allowed_sources(current_user.id, current_user.is_admin())
    
    # Get filter parameters
    source_filter = request.args.get('source', '')
    status_filter = request.args.get('status', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    search = request.args.get('search', '')
    quick_date = request.args.get('quick_date', '')
    manager_filter = request.args.get('manager', '')
    page = int(request.args.get('page', 1))
    per_page = 20
    
    # Get all managers
    all_managers = get_all_managers()
    
    # Validate source_filter against allowed sources
    if source_filter and allowed_sources is not None and source_filter not in allowed_sources:
        source_filter = ''
    
    # Process quick date filter
    today = date.today()
    if quick_date == 'this_month':
        date_from = today.replace(day=1).isoformat()
        date_to = today.isoformat()
    elif quick_date == 'last_month':
        first_of_this_month = today.replace(day=1)
        last_month_end = first_of_this_month - timedelta(days=1)
        last_month_start = last_month_end.replace(day=1)
        date_from = last_month_start.isoformat()
        date_to = last_month_end.isoformat()
    elif quick_date == 'last_quarter':
        date_from = (today - timedelta(days=90)).isoformat()
        date_to = today.isoformat()
    elif quick_date == 'half_year':
        date_from = (today - timedelta(days=180)).isoformat()
        date_to = today.isoformat()
    elif quick_date == 'one_year':
        date_from = (today - timedelta(days=365)).isoformat()
        date_to = today.isoformat()
    
    # Build base conditions with permission filter
    conditions = []
    params = []
    
    # Add permission filter first
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            conditions.append(f'source IN ({placeholders})')
            params.extend(allowed_sources)
        else:
            conditions.append('1=0')  # No access
            
    # Add manager filter
    if manager_filter:
        manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
        manager_urls = [s['url'] for s in manager_sites]
        
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            conditions.append(f'source IN ({placeholders})')
            params.extend(manager_urls)
        else:
            conditions.append('1=0')
    
    if source_filter:
        conditions.append('source = ?')
        params.append(source_filter)
    if status_filter:
        conditions.append('status = ?')
        params.append(status_filter)
    if date_from:
        conditions.append('date_created >= ?')
        params.append(date_from)
    if date_to:
        conditions.append('date_created <= ?')
        params.append(date_to + 'T23:59:59')
    if search:
        conditions.append('(number LIKE ? OR id LIKE ?)')
        params.extend([f'%{search}%', f'%{search}%'])
    
    where_clause = ' WHERE ' + ' AND '.join(conditions) if conditions else ''
    
    # Summary statistics by source (with currency)
    stats_query = f'''
        SELECT source, currency,
            COUNT(*) as total_orders,
            SUM(total) as total_amount,
            SUM(shipping_total) as total_shipping,
            SUM(CASE WHEN status NOT IN ('failed','cancelled') THEN 1 ELSE 0 END) as success_orders,
            SUM(CASE WHEN status NOT IN ('failed','cancelled') THEN total ELSE 0 END) as success_amount,
            SUM(CASE WHEN status NOT IN ('failed','cancelled') THEN shipping_total ELSE 0 END) as success_shipping,
            SUM(CASE WHEN status='failed' THEN 1 ELSE 0 END) as failed_orders,
            SUM(CASE WHEN status='cancelled' THEN 1 ELSE 0 END) as cancelled_orders
        FROM orders {where_clause} GROUP BY source, currency
    '''
    summary_raw = conn.execute(stats_query, params).fetchall()
    summary_stats = [dict(row) for row in summary_raw]
    
    # Get product quantities
    items_query = f'SELECT line_items, status, source FROM orders {where_clause}'
    all_items = conn.execute(items_query, params).fetchall()
    
    source_products = {}
    for row in all_items:
        src = row['source']
        if src not in source_products:
            source_products[src] = {'total': 0, 'success': 0, 'failed': 0, 'cancelled': 0}
        items = parse_json_field(row['line_items'])
        qty = sum(i.get('quantity', 0) for i in items) if isinstance(items, list) else 0
        source_products[src]['total'] += qty
        if row['status'] == 'failed':
            source_products[src]['failed'] += qty
        elif row['status'] == 'cancelled':
            source_products[src]['cancelled'] += qty
        else:
            source_products[src]['success'] += qty
    
    for stat in summary_stats:
        src = stat['source']
        prods = source_products.get(src, {})
        stat['total_products'] = prods.get('total', 0)
        stat['success_products'] = prods.get('success', 0)
        stat['failed_products'] = prods.get('failed', 0)
        stat['cancelled_products'] = prods.get('cancelled', 0)
        stat['success_net_amount'] = (stat['success_amount'] or 0) - (stat['success_shipping'] or 0)
    
    totals = {
        'total_orders': sum(s['total_orders'] or 0 for s in summary_stats),
        'total_products': sum(s['total_products'] for s in summary_stats),
        'total_amount': sum(s['total_amount'] or 0 for s in summary_stats),
        'success_orders': sum(s['success_orders'] or 0 for s in summary_stats),
        'success_products': sum(s['success_products'] for s in summary_stats),
        'success_amount': sum(s['success_amount'] or 0 for s in summary_stats),
        'success_net_amount': sum(s['success_net_amount'] for s in summary_stats),
        'failed_orders': sum(s['failed_orders'] or 0 for s in summary_stats),
        'cancelled_orders': sum(s['cancelled_orders'] or 0 for s in summary_stats),
    }
    
    # Get all available months for pagination
    months_query = f'''
        SELECT DISTINCT strftime('%Y-%m', date_created) as month 
        FROM orders {where_clause} 
        ORDER BY month DESC
    '''
    available_months = [row['month'] for row in conn.execute(months_query, params).fetchall()]
    
    # Get current month from page parameter
    current_month = request.args.get('month', '')
    if not current_month and available_months:
        current_month = available_months[0]  # Default to most recent month
    
    # Get orders for the selected month
    if current_month:
        month_conditions = conditions.copy() if conditions else []
        month_params = params.copy()
        month_conditions.append("strftime('%Y-%m', date_created) = ?")
        month_params.append(current_month)
        month_where = ' WHERE ' + ' AND '.join(month_conditions)
        
        orders_query = f'SELECT * FROM orders {month_where} ORDER BY date_created DESC'
        orders_data = conn.execute(orders_query, month_params).fetchall()
    else:
        orders_data = []
    
    # Get total count for display
    count_query = f'SELECT COUNT(*) FROM orders {where_clause}'
    total = conn.execute(count_query, params).fetchone()[0]
    
    processed_orders = []
    # 嘻嘻嘻嘻嘻
    for order in orders_data:
        od = dict(order)
        items = parse_json_field(order['line_items'])
        if isinstance(items, list):
            od['product_count'] = sum(i.get('quantity', 0) for i in items)
            od['products'] = [{
                'name': i.get('name', ''),
                'quantity': i.get('quantity', 0),
                'total': float(i.get('total', 0))
            } for i in items]
        else:
            od['product_count'] = 0
            od['products'] = []
        billing = parse_json_field(order['billing'])
        od['customer_name'] = f"{billing.get('first_name', '')} {billing.get('last_name', '')}".strip()
        od['customer_email'] = billing.get('email', '')
        processed_orders.append(od)
    
    # Get available sources (filtered by permissions and manager)
    conn = get_db_connection()
    
    # Base source query
    source_query = 'SELECT DISTINCT source FROM orders'
    source_params = []
    source_conditions = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(allowed_sources)
        else:
            source_conditions.append('1=0')
            
    if manager_filter:
        # Get sites managed by this manager
        manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
        manager_urls = [s['url'] for s in manager_sites]
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(manager_urls)
        else:
            source_conditions.append('1=0')
            
    if source_conditions:
        source_query += ' WHERE ' + ' AND '.join(source_conditions)
        
    source_query += ' ORDER BY source'
    sources = conn.execute(source_query, source_params).fetchall()
    statuses = conn.execute('SELECT DISTINCT status FROM orders').fetchall()
    
    # Get site managers mapping (url -> manager)
    sites = conn.execute('SELECT url, manager FROM sites').fetchall()
    site_managers = {s['url']: s['manager'] or '' for s in sites}
    
    # Add CNY conversion to summary_stats
    totals_cny = 0
    for stat in summary_stats:
        currency = stat.get('currency', 'PLN')
        # Use current_month for rate lookup
        rate, _ = get_cny_rate(currency, current_month)
        stat['rate_to_cny'] = rate
        if rate:
            stat['total_amount_cny'] = round((stat['total_amount'] or 0) * rate, 2)
            stat['success_amount_cny'] = round((stat['success_amount'] or 0) * rate, 2)
            stat['success_net_amount_cny'] = round((stat['success_net_amount'] or 0) * rate, 2)
            totals_cny += stat['success_net_amount_cny']
        else:
            stat['total_amount_cny'] = None
            stat['success_amount_cny'] = None
            stat['success_net_amount_cny'] = None
    
    totals['success_net_amount_cny'] = round(totals_cny, 2)
    
    # Add CNY conversion to processed_orders
    for order in processed_orders:
        currency = order.get('currency', 'PLN')
        order_month = order['date_created'][:7] if order.get('date_created') else current_month
        rate, _ = get_cny_rate(currency, order_month)
        order['rate_to_cny'] = rate
        order['total_cny'] = round(float(order['total'] or 0) * rate, 2) if rate else None
        order['net_total'] = float(order['total'] or 0) - float(order.get('shipping_total') or 0)
        order['net_total_cny'] = round(order['net_total'] * rate, 2) if rate else None
    
    # Get customer attributes for the displayed orders
    customer_emails = list(set(o['customer_email'] for o in processed_orders if o.get('customer_email')))
    customer_attributes = {}
    
    if customer_emails:
        placeholders = ', '.join(['?' for _ in customer_emails])
        
        # 1. Get manual quality settings
        manual_settings = conn.execute(f'''
            SELECT email, quality_tier FROM customer_settings 
            WHERE email IN ({placeholders})
        ''', customer_emails).fetchall()
        
        # 2. Calculate attributes for each customer
        for email in customer_emails:
            if email not in customer_attributes:
                customer_attributes[email] = {}
            
            # Check manual setting
            manual_tier = 'auto'
            for row in manual_settings:
                if row['email'] == email:
                    manual_tier = row['quality_tier']
                    break
            
            # Get customer stats (order count and total spending)
            # Note: This is a bit expensive inside a loop, but for 20 items it's acceptable.
            # Optimized query to get count and total in one go
            stats = conn.execute('''
                SELECT 
                    COUNT(*) as total_orders,
                    SUM(CASE WHEN status IN ('completed', 'processing') THEN 1 ELSE 0 END) as successful_orders,
                    SUM(CASE WHEN status IN ('completed', 'processing') THEN total ELSE 0 END) as total_spending,
                    MAX(date_created) as last_order_date,
                    MIN(date_created) as first_order_date
                FROM orders 
                WHERE billing LIKE ?
            ''', (f'%"{email}"%',)).fetchone()
            
            total_orders = stats['total_orders'] or 0
            successful_orders = stats['successful_orders'] or 0
            total_spending = stats['total_spending'] or 0
            
            # Store order count for display
            customer_attributes[email]['order_count'] = total_orders
            customer_attributes[email]['is_new'] = (total_orders <= 1)
            
            # Determine Tier
            tier = manual_tier
            
            if tier == 'auto':
                # Calculate auto tier using shared logic
                avg_days_between = 0
                if total_orders > 1 and stats['first_order_date'] and stats['last_order_date']:
                    from datetime import datetime
                    try:
                        first_date = datetime.fromisoformat(stats['first_order_date'][:19])
                        last_date = datetime.fromisoformat(stats['last_order_date'][:19])
                        days_span = (last_date - first_date).days
                        avg_days_between = days_span / (total_orders - 1)
                    except:
                        avg_days_between = 0
                
                tier = calculate_customer_tier(successful_orders, total_spending, avg_days_between)
            
            # Set attributes based on tier
            if tier == 'vip':
                customer_attributes[email]['quality'] = {'label': 'VIP', 'class': 'text-warning', 'icon': 'star-fill'}
            elif tier == 'good':
                customer_attributes[email]['quality'] = {'label': '优质', 'class': 'text-success', 'icon': 'gem'}
            elif tier == 'normal':
                customer_attributes[email]['quality'] = {'label': '普通', 'class': 'text-primary', 'icon': 'person-check'}
            elif tier == 'new':
                customer_attributes[email]['quality'] = {'label': '新客', 'class': 'text-info', 'icon': 'stars'}
            elif tier == 'bad':
                customer_attributes[email]['quality'] = {'label': '劣质', 'class': 'text-danger', 'icon': 'x-circle'}

    conn.close()
    
    return render_template('orders.html',
                         orders=processed_orders,
                         sources=sources,
                         statuses=statuses,
                         summary_stats=summary_stats,
                         totals=totals,
                         site_managers=site_managers,
                         customer_attributes=customer_attributes,
                         all_managers=all_managers,
                         current_filters={
                             'source': source_filter,
                             'status': status_filter,
                             'date_from': date_from,
                             'date_to': date_to,
                             'search': search,
                             'quick_date': quick_date,
                             'manager': manager_filter
                         },
                         available_months=available_months,
                         current_month=current_month,
                         total=total)



@app.route('/monthly')
@login_required
def monthly():
    """Monthly statistics page"""
    conn = get_db_connection()
    
    # Get user's allowed sources for permission filtering
    allowed_sources = get_user_allowed_sources(current_user.id, current_user.is_admin())
    
    # Get source filter
    source_filter = request.args.get('source', '')
    manager_filter = request.args.get('manager', '')
    
    # Get all managers
    all_managers = get_all_managers()
    
    # Validate source_filter against allowed sources
    if source_filter and allowed_sources is not None and source_filter not in allowed_sources:
        source_filter = ''
    
    # Get available sources (filtered by permissions and manager)
    conn = get_db_connection()
    
    # Base source query
    source_query = 'SELECT DISTINCT source FROM orders'
    source_params = []
    source_conditions = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(allowed_sources)
        else:
            source_conditions.append('1=0')
            
    if manager_filter:
        # Get sites managed by this manager
        manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
        manager_urls = [s['url'] for s in manager_sites]
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(manager_urls)
        else:
            source_conditions.append('1=0')
            
    if source_conditions:
        source_query += ' WHERE ' + ' AND '.join(source_conditions)
        
    source_query += ' ORDER BY source'
    all_sources = conn.execute(source_query, source_params).fetchall()
    
    # Build query with permission filter and optional source filter
    conditions = []
    params = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            conditions.append(f'source IN ({placeholders})')
            params.extend(allowed_sources)
        else:
            conditions.append('1=0')
            
    # Add manager filter
    if manager_filter:
        # Re-use manager_urls from above or fetch if not done
        if 'manager_urls' not in locals():
            manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
            manager_urls = [s['url'] for s in manager_sites]
        
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            conditions.append(f'source IN ({placeholders})')
            params.extend(manager_urls)
        else:
            conditions.append('1=0')
    
    if source_filter:
        conditions.append('source = ?')
        params.append(source_filter)
    
    where_clause = 'WHERE ' + ' AND '.join(conditions) if conditions else ''
    
    query = f'''
        SELECT id, status, date_created, total, shipping_total, source, currency, line_items
        FROM orders
        {where_clause}
        ORDER BY date_created DESC
    '''
    
    df = pd.read_sql_query(query, conn, params=params if params else None)
    conn.close()
    
    if len(df) == 0:
        return render_template('monthly.html', monthly_stats=[], sources=all_sources, source_filter=source_filter)
    
    # Calculate product quantities
    def get_product_qty(line_items):
        try:
            items = json.loads(line_items) if line_items else []
            return sum(item.get('quantity', 0) for item in items)
        except:
            return 0
    
    df['product_qty'] = df['line_items'].apply(get_product_qty)
    df['month'] = pd.to_datetime(df['date_created']).dt.to_period('M')
    
    # Group by month, source, and currency
    rows = []
    for (month, source), gdf in df.groupby(['month', 'source']):
        # Get the currency for this source (should be the same for all orders from same source)
        currency = gdf['currency'].iloc[0] if len(gdf) > 0 else 'N/A'
        
        total_orders = len(gdf)
        total_products = gdf['product_qty'].sum()
        total_amount = gdf['total'].sum()
        
        success_mask = ~gdf['status'].isin(['failed', 'cancelled'])
        success_orders = int(success_mask.sum())
        success_amount = gdf.loc[success_mask, 'total'].sum()
        success_products = gdf.loc[success_mask, 'product_qty'].sum()
        success_shipping = gdf.loc[success_mask, 'shipping_total'].sum()
        success_net_amount = success_amount - success_shipping
        
        failed_mask = gdf['status'] == 'failed'
        failed_orders = int(failed_mask.sum())
        
        cancelled_mask = gdf['status'] == 'cancelled'
        cancelled_orders = int(cancelled_mask.sum())
        
        rows.append({
            'month': str(month),
            'source': source,
            'currency': currency,
            'total_orders': total_orders,
            'total_products': int(total_products),
            'total_amount': float(total_amount),
            'success_orders': success_orders,
            'success_products': int(success_products),
            'success_amount': float(success_amount),
            'success_net_amount': float(success_net_amount),
            'failed_orders': failed_orders,
            'cancelled_orders': cancelled_orders
        })
    
    # Get sort parameter
    sort_by = request.args.get('sort', 'month')  # default: sort by month
    
    # Sort based on parameter
    if sort_by == 'source':
        rows.sort(key=lambda x: (x['source'], x['month']), reverse=False)
    else:
        rows.sort(key=lambda x: (x['month'], x['source']), reverse=True)
    
    # Get site managers mapping
    conn2 = get_db_connection()
    sites = conn2.execute('SELECT url, manager FROM sites').fetchall()
    site_managers = {s['url']: s['manager'] or '' for s in sites}
    conn2.close()
    
    # Add CNY conversion to each row
    for row in rows:
        currency = row.get('currency', 'PLN')
        month = row.get('month', '')
        rate, _ = get_cny_rate(currency, month)
        row['rate_to_cny'] = rate
        row['success_net_amount_cny'] = round(row['success_net_amount'] * rate, 2) if rate else None
    
    return render_template('monthly.html', monthly_stats=rows, sources=all_sources, source_filter=source_filter, manager_filter=manager_filter, all_managers=all_managers, sort_by=sort_by, site_managers=site_managers)


@app.route('/customers')
@login_required
def customers():
    conn = get_db_connection()
    
    # Get user's allowed sources for permission filtering
    allowed_sources = get_user_allowed_sources(current_user.id, current_user.is_admin())
    
    # Get source filter
    source_filter = request.args.get('source', '')
    manager_filter = request.args.get('manager', '')
    
    # Get all managers
    all_managers = get_all_managers()
    
    # Validate source_filter against allowed sources
    if source_filter and allowed_sources is not None and source_filter not in allowed_sources:
        source_filter = ''
    
    # Get available sources (filtered by permissions and manager)
    conn = get_db_connection()
    
    # Base source query
    source_query = 'SELECT DISTINCT source FROM orders'
    source_params = []
    source_conditions = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(allowed_sources)
        else:
            source_conditions.append('1=0')
            
    if manager_filter:
        # Get sites managed by this manager
        manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
        manager_urls = [s['url'] for s in manager_sites]
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(manager_urls)
        else:
            source_conditions.append('1=0')
            
    if source_conditions:
        source_query += ' WHERE ' + ' AND '.join(source_conditions)
        
    source_query += ' ORDER BY source'
    all_sources = conn.execute(source_query, source_params).fetchall()
    
    # Build query with permission filter and optional source filter
    base_condition = "json_extract(billing, '$.email') IS NOT NULL AND json_extract(billing, '$.email') != ''"
    conditions = [base_condition]
    params = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            conditions.append(f'source IN ({placeholders})')
            params.extend(allowed_sources)
        else:
            conditions.append('1=0')
            
    # Add manager filter
    if manager_filter:
        # Re-use manager_urls from above or fetch if not done
        if 'manager_urls' not in locals():
            manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
            manager_urls = [s['url'] for s in manager_sites]
        
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            conditions.append(f'source IN ({placeholders})')
            params.extend(manager_urls)
        else:
            conditions.append('1=0')
    
    if source_filter:
        conditions.append('source = ?')
        params.append(source_filter)
    
    where_clause = 'WHERE ' + ' AND '.join(conditions)
    
    query = f'''
        SELECT 
            json_extract(billing, '$.email') as email,
            json_extract(billing, '$.first_name') || ' ' || json_extract(billing, '$.last_name') as name,
            json_extract(billing, '$.phone') as phone,
            COUNT(*) as total_orders,
            SUM(CASE WHEN status IN ('completed', 'processing') THEN 1 ELSE 0 END) as successful_orders,
            SUM(CASE WHEN status IN ('completed', 'processing') THEN total ELSE 0 END) as total_spent,
            MAX(date_created) as last_order_date,
            MIN(date_created) as first_order_date,
            GROUP_CONCAT(DISTINCT source) as sources
        FROM orders
        {where_clause}
        GROUP BY email
        ORDER BY total_spent DESC
    '''
    
    customers_data = conn.execute(query, params).fetchall()
    
    conn.close()
    
    customers_list = []
    total_customers = len(customers_data)
    total_revenue = 0
    repeat_customers = 0
    new_customers_month = 0
    new_customers_last_month = 0
    
    import datetime
    now = datetime.datetime.now()
    thirty_days_ago = (now - datetime.timedelta(days=30)).strftime('%Y-%m-%d')
    sixty_days_ago = (now - datetime.timedelta(days=60)).strftime('%Y-%m-%d')
    ninety_days_ago = (now - datetime.timedelta(days=90)).strftime('%Y-%m-%d')
    
    tier_counts = {'VIP': 0, '优质': 0, '普通': 0, '新客': 0}
    
    for row in customers_data:
        c = dict(row)
        c['total_spent'] = float(c['total_spent'] or 0)
        total_revenue += c['total_spent']
        
        if c['successful_orders'] > 1:
            repeat_customers += 1
            
        if c['first_order_date'] >= thirty_days_ago:
            new_customers_month += 1
        elif c['first_order_date'] >= sixty_days_ago:
            new_customers_last_month += 1
            
        # Calculate Tier (same logic as API)
        avg_days = 0
        if c['successful_orders'] > 1:
            first = datetime.datetime.fromisoformat(c['first_order_date'])
            last = datetime.datetime.fromisoformat(c['last_order_date'])
            days = (last - first).days or 1
            avg_days = days / c['successful_orders']
            
        score = min(c['successful_orders'] * 10, 30) + min(c['total_spent'] / 100, 40)
        score += 30 if avg_days > 0 and avg_days < 60 else (15 if avg_days < 120 else 0)
        score = min(score, 100)
        
        if score >= 80: tier = 'VIP'
        elif score >= 60: tier = '优质'
        elif score >= 40: tier = '普通'
        else: tier = '新客'
        
        c['tier'] = tier
        tier_counts[tier] += 1
        
        # Smart Actions
        actions = []
        if tier == 'VIP':
            actions.append({'type': 'success', 'icon': 'gift', 'text': '专属礼遇'})
            actions.append({'type': 'primary', 'icon': 'people', 'text': '邀请入群'})
        elif c['last_order_date'] < ninety_days_ago:
            actions.append({'type': 'warning', 'icon': 'ticket-perforated', 'text': '召回优惠券'})
        elif c['successful_orders'] == 1 and c['first_order_date'] >= thirty_days_ago:
            actions.append({'type': 'info', 'icon': 'book', 'text': '欢迎指南'})
            actions.append({'type': 'info', 'icon': 'bag-plus', 'text': '关联推荐'})
        elif c['successful_orders'] > 3:
             actions.append({'type': 'primary', 'icon': 'arrow-repeat', 'text': '订阅服务'})
        
        # Process sources
        source_str = c.get('sources', '') or ''
        sources = list(set([s.strip() for s in source_str.split(',') if s.strip()]))
        c['source'] = sources[0] if sources else 'Unknown'
        c['all_sources'] = sources

        c['actions'] = actions
        customers_list.append(c)
        
    # Calculate growth rate
    if new_customers_last_month > 0:
        growth_rate = ((new_customers_month - new_customers_last_month) / new_customers_last_month) * 100
    else:
        growth_rate = 100 if new_customers_month > 0 else 0

    stats = {
        'total_customers': total_customers,
        'avg_ltv': total_revenue / total_customers if total_customers > 0 else 0,
        'repeat_rate': (repeat_customers / total_customers * 100) if total_customers > 0 else 0,
        'new_customer_rate': (new_customers_month / total_customers * 100) if total_customers > 0 else 0,
        'new_customers_month': new_customers_month,
        'new_customers_last_month': new_customers_last_month,
        'growth_rate': growth_rate,
        'tier_counts': tier_counts
    }
    
    # Get site managers mapping
    conn2 = get_db_connection()
    sites = conn2.execute('SELECT url, manager FROM sites').fetchall()
    site_managers = {s['url']: s['manager'] or '' for s in sites}
    conn2.close()
    
    return render_template('customers.html', customers=customers_list, stats=stats, sources=all_sources, source_filter=source_filter, manager_filter=manager_filter, all_managers=all_managers, site_managers=site_managers)


@app.route('/api/chart-data')
@login_required
def chart_data():
    """API endpoint for chart data"""
    conn = get_db_connection()
    
    chart_type = request.args.get('type', 'monthly')
    
    if chart_type == 'monthly':
        data = conn.execute('''
            SELECT strftime('%Y-%m', date_created) as month,
                   source,
                   COUNT(*) as orders,
                   SUM(total) as revenue
            FROM orders
            WHERE date_created >= date('now', '-12 months')
            GROUP BY month, source
            ORDER BY month
        ''').fetchall()
        
        result = {}
        for row in data:
            month = row['month']
            if month not in result:
                result[month] = {}
            result[month][row['source']] = {
                'orders': row['orders'],
                'revenue': row['revenue']
            }
        
        conn.close()
        return jsonify(result)
    
    elif chart_type == 'status':
        data = conn.execute('''
            SELECT status, COUNT(*) as count
            FROM orders
            GROUP BY status
        ''').fetchall()
        
        conn.close()
        return jsonify([dict(row) for row in data])
    
    conn.close()
    return jsonify({})


@app.route('/api/order/<order_id>')
@login_required
def get_order_details(order_id):
    """API endpoint to get order details"""
    conn = get_db_connection()
    
    order = conn.execute('''
        SELECT * FROM orders WHERE id = ?
    ''', (order_id,)).fetchone()
    
    if not order:
        return jsonify({'error': 'Order not found'}), 404
    
    order_dict = dict(order)
    
    # Parse JSON fields
    order_dict['billing'] = parse_json_field(order['billing'])
    order_dict['shipping'] = parse_json_field(order['shipping'])
    order_dict['line_items'] = parse_json_field(order['line_items'])
    order_dict['shipping_lines'] = parse_json_field(order['shipping_lines'])
    order_dict['meta_data'] = parse_json_field(order['meta_data'])
    order_dict['fee_lines'] = parse_json_field(order['fee_lines'])
    order_dict['coupon_lines'] = parse_json_field(order['coupon_lines'])
    order_dict['coupon_lines'] = parse_json_field(order['coupon_lines'])
    order_dict['refunds'] = parse_json_field(order['refunds'])
    
    # Calculate customer total spending
    if order_dict['billing'] and order_dict['billing'].get('email'):
        email = order_dict['billing']['email']
        customer_stats = conn.execute('''
            SELECT COUNT(*) as count, SUM(total) as total
            FROM orders 
            WHERE billing LIKE ? AND status IN ('completed', 'processing')
        ''', (f'%"{email}"%',)).fetchone()
        
        order_dict['customer_stats'] = {
            'total_orders': customer_stats['count'] if customer_stats else 0,
            'total_spent': float(customer_stats['total'] or 0) if customer_stats else 0
        }
    
    # Get site manager for this order's source
    site_row = conn.execute('SELECT manager FROM sites WHERE url = ?', (order_dict.get('source', ''),)).fetchone()
    order_dict['site_manager'] = site_row['manager'] if site_row and site_row['manager'] else ''
    
    conn.close()
    
    return jsonify(order_dict)


@app.route('/api/customer/<email>')
@login_required
def get_customer_details(email):
    """API endpoint to get customer analysis data"""
    from urllib.parse import unquote
    email = unquote(email)
    
    conn = get_db_connection()
    
    # Get all orders from this customer
    orders = conn.execute('''
        SELECT id, number, status, total, currency, shipping_total, date_created, source, line_items, billing
        FROM orders
        WHERE billing LIKE ?
        ORDER BY date_created DESC
    ''', (f'%"{email}"%',)).fetchall()
    
    conn.close()
    
    if not orders:
        return jsonify({'error': 'Customer not found'}), 404
    
    # Process orders
    order_list = []
    all_products = {}
    site_spending = {}
    spending_by_currency = {}  # Track spending by currency
    total_spending = 0
    successful_orders = 0
    failed_orders = 0
    cancelled_orders = 0
    dates = []
    
    for order in orders:
        order_dict = dict(order)
        billing = parse_json_field(order['billing'])
        line_items = parse_json_field(order['line_items'])
        
        # Get customer name
        customer_name = f"{billing.get('first_name', '')} {billing.get('last_name', '')}".strip()
        customer_phone = billing.get('phone', '')
        
        # Calculate order products
        order_products = []
        order_qty = 0
        for item in (line_items or []):
            # Get full product name including flavor from meta_data
            full_name, flavor, meta_puffs = get_full_product_name(item)
            product_name = full_name or item.get('name', 'Unknown')
            qty = item.get('quantity', 1)
            total = float(item.get('total', 0))
            order_products.append({
                'name': product_name,
                'quantity': qty,
                'total': total
            })
            order_qty += qty
            
            # Aggregate products
            if product_name in all_products:
                all_products[product_name]['quantity'] += qty
                all_products[product_name]['total'] += total
            else:
                all_products[product_name] = {'quantity': qty, 'total': total}
        
        # Site spending
        source = order['source'] or 'Unknown'
        if source in site_spending:
            site_spending[source]['orders'] += 1
            site_spending[source]['amount'] += float(order['total'] or 0)
        else:
            site_spending[source] = {'orders': 1, 'amount': float(order['total'] or 0)}
        
        # Status counting
        status = order['status']
        currency = order['currency'] or 'N/A'
        if status in ['completed', 'processing']:
            successful_orders += 1
            order_total = float(order['total'] or 0)
            total_spending += order_total
            # Track by currency
            if currency not in spending_by_currency:
                spending_by_currency[currency] = 0
            spending_by_currency[currency] += order_total
        elif status == 'failed':
            failed_orders += 1
        elif status == 'cancelled':
            cancelled_orders += 1
        
        # Dates
        if order['date_created']:
            dates.append(order['date_created'][:10])
        
        order_list.append({
            'id': order['id'],
            'number': order['number'],
            'status': status,
            'total': float(order['total'] or 0),
            'currency': currency,
            'date_created': order['date_created'],
            'source': source.replace('https://www.', ''),
            'product_count': order_qty,
            'products': order_products
        })
    
    # Calculate frequency
    unique_dates = sorted(set(dates))
    if len(unique_dates) >= 2:
        from datetime import datetime
        first_date = datetime.fromisoformat(unique_dates[0])
        last_date = datetime.fromisoformat(unique_dates[-1])
        days_span = (last_date - first_date).days or 1
        avg_days_between = days_span / len(unique_dates)
    else:
        avg_days_between = 0
    
    # Sort products by quantity
    sorted_products = sorted(all_products.items(), key=lambda x: x[1]['quantity'], reverse=True)
    
    # Customer quality score (simple algorithm)
    quality_score = 0
    quality_score += min(successful_orders * 10, 30)  # Max 30 for orders
    quality_score += min(total_spending / 100, 40)    # Max 40 for spending
    quality_score += 30 if avg_days_between > 0 and avg_days_between < 60 else (15 if avg_days_between < 120 else 0)  # Frequency bonus
    quality_score = min(quality_score, 100)
    
    # Check for manual override
    conn = get_db_connection()
    manual_setting = conn.execute('SELECT quality_tier FROM customer_settings WHERE email = ?', (email,)).fetchone()
    conn.close()
    manual_tier = manual_setting['quality_tier'] if manual_setting else 'auto'
    
    # Determine customer tier
    if manual_tier != 'auto':
        if manual_tier == 'vip':
            customer_tier = {'level': 'VIP', 'color': '#f59e0b', 'icon': 'star-fill', 'manual': True}
        elif manual_tier == 'good':
            customer_tier = {'level': '优质', 'color': '#10b981', 'icon': 'gem', 'manual': True}
        elif manual_tier == 'normal':
            customer_tier = {'level': '普通', 'color': '#3b82f6', 'icon': 'person-check', 'manual': True}
        elif manual_tier == 'new':
            customer_tier = {'level': '新客', 'color': '#0dcaf0', 'icon': 'stars', 'manual': True}
        elif manual_tier == 'bad':
            customer_tier = {'level': '劣质', 'color': '#ef4444', 'icon': 'x-circle', 'manual': True}
        else:
            # Fallback to auto if unknown
            manual_tier = 'auto'
            
    if manual_tier == 'auto':
        tier = calculate_customer_tier(successful_orders, total_spending, avg_days_between)
        if tier == 'vip':
            customer_tier = {'level': 'VIP', 'color': '#f59e0b', 'icon': 'star-fill', 'manual': False}
        elif tier == 'good':
            customer_tier = {'level': '优质', 'color': '#10b981', 'icon': 'gem', 'manual': False}
        elif tier == 'normal':
            customer_tier = {'level': '普通', 'color': '#3b82f6', 'icon': 'person-check', 'manual': False}
        else:
            customer_tier = {'level': '新客', 'color': '#0dcaf0', 'icon': 'stars', 'manual': False}
            
    customer_tier['value'] = manual_tier
    
    # Calculate CNY total for customer spending
    from datetime import datetime
    current_month = datetime.now().strftime('%Y-%m')
    spending_cny = 0
    for currency, amount in spending_by_currency.items():
        rate, _ = get_cny_rate(currency, current_month)
        if rate:
            spending_cny += amount * rate
    
    result = {
        'email': email,
        'name': customer_name,
        'phone': customer_phone,
        'total_orders': len(order_list),
        'successful_orders': successful_orders,
        'failed_orders': failed_orders,
        'cancelled_orders': cancelled_orders,
        'total_spending': total_spending,
        'spending_by_currency': spending_by_currency,  # New: currency breakdown
        'spending_cny': round(spending_cny, 2),  # New: CNY total
        'avg_order_value': total_spending / successful_orders if successful_orders > 0 else 0,
        'first_order_date': unique_dates[0] if unique_dates else None,
        'last_order_date': unique_dates[-1] if unique_dates else None,
        'avg_days_between_orders': round(avg_days_between, 1),
        'quality_score': round(quality_score),
        'customer_tier': customer_tier,
        'site_spending': [{'site': k.replace('https://www.', ''), 'orders': v['orders'], 'amount': v['amount']} for k, v in site_spending.items()],
        'top_products': [{'name': k, 'quantity': v['quantity'], 'total': v['total']} for k, v in sorted_products[:10]],
        'orders': order_list[:20]  # Limit to 20 most recent
    }
    
    return jsonify(result)


@app.route('/api/customer/quality', methods=['POST'])
@login_required
def update_customer_quality():
    """Update customer quality tier manually"""
    data = request.json
    email = data.get('email')
    quality = data.get('quality')
    
    if not email or not quality:
        return jsonify({'success': False, 'error': 'Missing required fields'}), 400
        
    conn = get_db_connection()
    try:
        conn.execute('''
            INSERT INTO customer_settings (email, quality_tier, updated_at) 
            VALUES (?, ?, datetime('now'))
            ON CONFLICT(email) DO UPDATE SET 
                quality_tier = excluded.quality_tier,
                updated_at = excluded.updated_at
        ''', (email, quality))
        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        conn.close()


# Status translation
STATUS_LABELS = {
    'pending': '待处理',
    'processing': '处理中',
    'on-hold': '已发货',
    'completed': '已完成',
    'cancelled': '已取消',
    'refunded': '已退款',
    'failed': '失败'
}

# Currency symbols mapping
CURRENCY_SYMBOLS = {
    'PLN': 'zł',      # Polish Złoty
    'AUD': 'A$',      # Australian Dollar
    'AED': 'د.إ',     # UAE Dirham
    'USD': '$',       # US Dollar
    'EUR': '€',       # Euro
    'GBP': '£',       # British Pound
}


def format_amount_with_currency(amount, currency):
    """Format amount with currency code, e.g., '145.00 PLN'"""
    try:
        return f"{float(amount):,.2f} {currency}"
    except:
        return f"{amount} {currency}"


def get_currency_symbol(currency):
    """Get symbol for a currency code"""
    return CURRENCY_SYMBOLS.get(currency, currency)


@app.template_filter('status_label')
def status_label_filter(status):
    return STATUS_LABELS.get(status, status)


@app.template_filter('format_currency')
def format_currency_filter(value, currency=None):
    """Format currency value, optionally with currency code"""
    try:
        formatted = f"{float(value):,.2f}"
        if currency:
            return f"{formatted} {currency}"
        return formatted
    except:
        return value


@app.template_filter('format_date')
def format_date_filter(value):
    try:
        dt = datetime.fromisoformat(value.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d %H:%M')
    except:
        return value



# -----------------------------
# Data Synchronization Features
# -----------------------------

def init_sites_table():
    """Initialize sites table"""
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS sites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            url TEXT NOT NULL,
            consumer_key TEXT NOT NULL,
            consumer_secret TEXT NOT NULL,
            manager TEXT,
            last_sync TEXT
        )
    ''')
    # Add manager column if not exists (for existing databases)
    try:
        conn.execute('ALTER TABLE sites ADD COLUMN manager TEXT')
    except:
        pass  # Column already exists
    conn.commit()
    conn.close()


def init_sync_logs_table():
    """Initialize sync_logs table for storing synchronization history"""
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS sync_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id INTEGER,
            site_url TEXT,
            status TEXT,
            message TEXT,
            new_orders INTEGER DEFAULT 0,
            updated_orders INTEGER DEFAULT 0,
            sync_time TEXT,
            duration_seconds INTEGER
        )
    ''')
    conn.commit()
    conn.close()


def init_settings_table():
    """Initialize settings table for storing app configuration"""
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    # Insert default autosync settings if not exist
    conn.execute('''
        INSERT OR IGNORE INTO settings (key, value) VALUES ('autosync_enabled', 'false')
    ''')
    conn.execute('''
        INSERT OR IGNORE INTO settings (key, value) VALUES ('autosync_interval', '900')
    ''')
    conn.execute('''
        INSERT OR IGNORE INTO settings (key, value) VALUES ('source_display_mode', 'full')
    ''')
    conn.commit()
    conn.close()


@app.context_processor
def inject_settings():
    """Inject global settings into all templates"""
    conn = get_db_connection()
    mode = conn.execute("SELECT value FROM settings WHERE key = 'source_display_mode'").fetchone()
    conn.close()
    return dict(source_display_mode=mode['value'] if mode else 'full')


def save_sync_log(site_id, site_url, status, message, new_orders=0, updated_orders=0, duration_seconds=0):
    """Save a sync log entry to the database"""
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO sync_logs (site_id, site_url, status, message, new_orders, updated_orders, sync_time, duration_seconds)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (site_id, site_url, status, message, new_orders, updated_orders, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), duration_seconds))
    conn.commit()
    conn.close()

def init_users_table():
    """Initialize users table"""
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            name TEXT,
            role TEXT DEFAULT 'user',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS user_site_permissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            site_id INTEGER,
            UNIQUE(user_id, site_id)
        )
    ''')
    conn.commit()
    
    # Migrate default admin user if not exists
    existing = conn.execute('SELECT id FROM users WHERE username = ?', ('admin',)).fetchone()
    if not existing:
        conn.execute('''
            INSERT INTO users (username, password_hash, name, role)
            VALUES (?, ?, ?, ?)
        ''', ('admin', generate_password_hash('admin123'), '管理员', 'admin'))
        conn.commit()
    
    conn.close()


def init_product_tables():
    """Initialize product analysis tables (brands, series, product_mappings)"""
    conn = get_db_connection()
    
    # Brands table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS brands (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            aliases TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Series table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS series (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            brand_id INTEGER,
            name TEXT NOT NULL,
            aliases TEXT,
            UNIQUE(brand_id, name),
            FOREIGN KEY (brand_id) REFERENCES brands(id)
        )
    ''')
    
    # Product mappings table
    conn.execute('''
        CREATE TABLE IF NOT EXISTS product_mappings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            raw_name TEXT NOT NULL,
            brand_id INTEGER,
            series_id INTEGER,
            puff_count INTEGER,
            flavor TEXT,
            is_manual INTEGER DEFAULT 0,
            source TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(raw_name, source),
            FOREIGN KEY (brand_id) REFERENCES brands(id),
            FOREIGN KEY (series_id) REFERENCES series(id)
        )
    ''')
    
    conn.commit()
    
    # Insert default brands if table is empty
    existing = conn.execute('SELECT COUNT(*) FROM brands').fetchone()[0]
    if existing == 0:
        default_brands = [
            ('IGET', '["iget", "I-GET"]'),
            ('FUMO', '["Fumo"]'),
            ('Crystal Blind', '["CRYSTAL BLIND", "Crystal-Blind"]'),
            ('POD SALT', '["Pod Salt", "PODSALT"]'),
            ('Waka', '["WAKA"]'),
            ('UMIN', '["Umin"]'),
            ('Isgo', '["ISGO", "Isgo bar"]'),
            ('FIZZY', '["Fizzy", "FIZZY TWINS"]'),
            ('FUMOT', '["Fumot"]'),
            ('Esin', '["ESIN", "Esin Vape"]'),
            ('XCOR', '["Xcor", "X-COR"]'),
            ('ELF BAR', '["Elf Bar", "ELFBAR"]'),
            ('Lost Mary', '["LOST MARY", "LostMary"]'),
            ('Geek Bar', '["GEEK BAR", "GeekBar"]'),
        ]
        conn.executemany('INSERT OR IGNORE INTO brands (name, aliases) VALUES (?, ?)', default_brands)
        conn.commit()
    
    conn.close()


import re

def parse_product_name(name, brands_cache=None):
    """
    Parse product name to extract brand, series, puff count, and flavor.
    
    Examples:
    - "IGET ONE 12000 puffs - Mixed Berries" → brand: IGET, puffs: 12000, flavor: Mixed Berries
    - "Crystal Blind 25000 Puffs" → brand: Crystal Blind, puffs: 25000
    - "FUMO king 6000 puffs Disposable Vape 20mg" → brand: FUMO, series: king, puffs: 6000
    """
    if not name:
        return {'brand': None, 'series': None, 'puffs': None, 'flavor': None, 'normalized': None}
    
    result = {'brand': None, 'series': None, 'puffs': None, 'flavor': None, 'normalized': None}
    
    # 1. Extract puff count
    puff_match = re.search(r'(\d+)\s*puffs?', name, re.IGNORECASE)
    if puff_match:
        result['puffs'] = int(puff_match.group(1))
    
    # 2. Get brands from cache or database
    if brands_cache is None:
        conn = get_db_connection()
        brands_rows = conn.execute('SELECT id, name, aliases FROM brands').fetchall()
        conn.close()
        brands_cache = []
        for row in brands_rows:
            brand_name = row['name']
            aliases = []
            if row['aliases']:
                try:
                    aliases = json.loads(row['aliases'])
                except:
                    pass
            brands_cache.append({
                'id': row['id'],
                'name': brand_name,
                'aliases': aliases,
                'patterns': [brand_name.upper()] + [a.upper() for a in aliases]
            })
    
    # 3. Match brand (longest match first)
    name_upper = name.upper()
    matched_brand = None
    matched_len = 0
    
    for brand in brands_cache:
        for pattern in brand['patterns']:
            if pattern in name_upper and len(pattern) > matched_len:
                matched_brand = brand
                matched_len = len(pattern)
    
    if matched_brand:
        result['brand'] = matched_brand['name']
        result['brand_id'] = matched_brand['id']
    
    # 4. Extract flavor (usually after separator)
    flavor = None
    for sep in [' - ', ' – ', ' | ', ' / ']:
        if sep in name:
            parts = name.split(sep)
            if len(parts) > 1:
                flavor = parts[-1].strip()
                # Don't treat product type/specs as flavor
                if any(x in flavor.lower() for x in ['puff', 'disposable', 'vape', 'mg', 'ml']):
                    flavor = None
                break
    
    result['flavor'] = flavor
    
    # 5. Build normalized name
    parts = []
    if result['brand']:
        parts.append(result['brand'])
    if result['puffs']:
        parts.append(f"{result['puffs']} Puffs")
    if result['flavor']:
        parts.append(result['flavor'])
    
    result['normalized'] = ' - '.join(parts) if parts else name
    
    return result


# Initialize tables on startup
with app.app_context():
    init_sites_table()
    init_sync_logs_table()
    init_settings_table()
    init_users_table()
    init_product_tables()

@app.route('/settings')
@login_required
@admin_required
def settings():
    """Settings page for site management - Admin only"""
    conn = get_db_connection()
    sites = conn.execute('SELECT * FROM sites').fetchall()
    
    # Get exchange rates
    exchange_rates = conn.execute('''
        SELECT id, year_month, currency, rate_to_cny, updated_at
        FROM exchange_rates
        ORDER BY year_month DESC, currency
    ''').fetchall()
    
    # Get distinct currencies from orders
    currencies = conn.execute('SELECT DISTINCT currency FROM orders WHERE currency IS NOT NULL').fetchall()
    currency_list = [c['currency'] for c in currencies if c['currency']]
    
    conn.close()
    return render_template('settings.html', 
                          sites=sites, 
                          exchange_rates=exchange_rates,
                          currencies=currency_list)


@app.route('/api/exchange-rates', methods=['GET', 'POST'])
@login_required
@admin_required
def exchange_rates_api():
    """API for managing exchange rates"""
    conn = get_db_connection()
    
    if request.method == 'GET':
        rates = conn.execute('''
            SELECT id, year_month, currency, rate_to_cny, updated_at
            FROM exchange_rates
            ORDER BY year_month DESC, currency
        ''').fetchall()
        conn.close()
        return jsonify([dict(r) for r in rates])
    
    elif request.method == 'POST':
        data = request.get_json()
        year_month = data.get('year_month', '')
        currency = data.get('currency', '').upper()
        rate_to_cny = data.get('rate_to_cny')
        
        if not year_month or not currency or rate_to_cny is None:
            conn.close()
            return jsonify({'error': '请填写所有字段'}), 400
        
        try:
            rate_to_cny = float(rate_to_cny)
            if rate_to_cny <= 0:
                raise ValueError("Rate must be positive")
        except (ValueError, TypeError):
            conn.close()
            return jsonify({'error': '汇率必须是正数'}), 400
        
        try:
            conn.execute('''
                INSERT OR REPLACE INTO exchange_rates (year_month, currency, rate_to_cny, updated_at)
                VALUES (?, ?, ?, datetime('now'))
            ''', (year_month, currency, rate_to_cny))
            conn.commit()
            conn.close()
            return jsonify({'success': True, 'message': '汇率保存成功'})
        except Exception as e:
            conn.close()
            return jsonify({'error': str(e)}), 500


@app.route('/api/exchange-rates/<int:rate_id>', methods=['DELETE'])
@login_required
@admin_required
def delete_exchange_rate(rate_id):
    """Delete an exchange rate"""
    conn = get_db_connection()
    conn.execute('DELETE FROM exchange_rates WHERE id = ?', (rate_id,))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '汇率已删除'})


@app.route('/api/sites/import-from-script', methods=['POST'])
@login_required
def import_sites_from_script():
    """从 1.wooorders_sqlite.py 的硬编码配置导入站点到数据库"""
    import ast
    import re
    
    script_path = '1.wooorders_sqlite.py'
    
    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 提取 HARDCODED_SITES 列表
        pattern = r'HARDCODED_SITES\s*=\s*(\[[\s\S]*?\n\])'
        match = re.search(pattern, content)
        
        if not match:
            # 尝试旧格式 sites = [...]
            pattern = r'^sites\s*=\s*(\[[\s\S]*?\n\])'
            match = re.search(pattern, content, re.MULTILINE)
        
        if not match:
            return jsonify({'error': '无法在脚本中找到站点配置'}), 400
        
        # 解析 Python 列表
        sites_str = match.group(1)
        sites_list = ast.literal_eval(sites_str)
        
        conn = get_db_connection()
        imported = 0
        skipped = 0
        
        for site in sites_list:
            url = site.get('url', '')
            ck = site.get('ck', '')
            cs = site.get('cs', '')
            
            if not all([url, ck, cs]):
                continue
            
            # 检查是否已存在
            existing = conn.execute('SELECT id FROM sites WHERE url = ?', (url,)).fetchone()
            if existing:
                skipped += 1
                continue
            
            conn.execute('INSERT INTO sites (url, consumer_key, consumer_secret) VALUES (?, ?, ?)',
                        (url, ck, cs))
            imported += 1
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'success': True, 
            'imported': imported, 
            'skipped': skipped,
            'message': f'成功导入 {imported} 个站点，跳过 {skipped} 个已存在的站点'
        })
        
    except FileNotFoundError:
        return jsonify({'error': f'找不到脚本文件: {script_path}'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/sites', methods=['POST'])
@login_required
def add_site():
    """Add a new WooCommerce site"""
    data = request.json
    url = data.get('url')
    ck = data.get('consumer_key')
    cs = data.get('consumer_secret')
    manager = data.get('manager', '')
    
    if not all([url, ck, cs]):
        return jsonify({'error': 'Missing required fields'}), 400
        
    # Remove trailing slash from URL
    if url.endswith('/'):
        url = url[:-1]
        
    conn = get_db_connection()
    try:
        conn.execute('INSERT INTO sites (url, consumer_key, consumer_secret, manager) VALUES (?, ?, ?, ?)',
                     (url, ck, cs, manager))
        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()



@app.route('/api/sites/<int:site_id>', methods=['PUT'])
@login_required
def update_site(site_id):
    """Update a WooCommerce site"""
    data = request.json
    url = data.get('url')
    ck = data.get('consumer_key')
    cs = data.get('consumer_secret')
    manager = data.get('manager', '')
    
    if not all([url, ck, cs]):
        return jsonify({'error': 'Missing required fields'}), 400
        
    # Remove trailing slash from URL
    if url.endswith('/'):
        url = url[:-1]
        
    conn = get_db_connection()
    try:
        conn.execute('UPDATE sites SET url = ?, consumer_key = ?, consumer_secret = ?, manager = ? WHERE id = ?',
                     (url, ck, cs, manager, site_id))
        conn.commit()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()

@app.route('/api/sites/<int:site_id>', methods=['DELETE'])
@login_required
def delete_site(site_id):
    """Delete a WooCommerce site"""
    print(f"Received delete request for site_id: {site_id}") # Debug log
    conn = get_db_connection()
    try:
        conn.execute('DELETE FROM sites WHERE id = ?', (site_id,))
        conn.commit()
        print(f"Successfully deleted site_id: {site_id}") # Debug log
        return jsonify({'success': True})
    except Exception as e:
        print(f"Error deleting site: {e}") # Debug log
        return jsonify({'error': str(e)}), 500
    finally:
        conn.close()

# Global sync status storage
# Format: {site_id: {'status': 'idle'|'running'|'success'|'error', 'progress': 0, 'message': '', 'logs': []}}
SYNC_STATUS = {}

@app.route('/api/sync/status/<int:site_id>')
@login_required
def get_sync_status(site_id):
    """Get synchronization status for a site"""
    status = SYNC_STATUS.get(site_id, {'status': 'idle', 'message': '', 'logs': []})
    return jsonify(status)

@app.route('/api/sync', methods=['POST'])
@login_required
def sync_data():
    """Trigger data synchronization"""
    import sync_utils
    import threading
    
    site_id = request.json.get('site_id')
    if not site_id:
        return jsonify({'error': 'Missing site_id'}), 400
        
    site_id = int(site_id)
    
    # Initialize status
    SYNC_STATUS[site_id] = {
        'status': 'running',
        'message': 'Starting synchronization...',
        'logs': [f"[{datetime.now().strftime('%H:%M:%S')}] Job started"]
    }
    
    def run_sync(app_context, site_id):
        with app_context:
            sync_start_time = datetime.now()
            try:
                conn = get_db_connection()
                site = conn.execute('SELECT * FROM sites WHERE id = ?', (site_id,)).fetchone()
                conn.close()
                
                if not site:
                    SYNC_STATUS[site_id]['status'] = 'error'
                    SYNC_STATUS[site_id]['message'] = 'Site not found'
                    return

                def progress_callback(msg):
                    timestamp = datetime.now().strftime('%H:%M:%S')
                    log_entry = f"[{timestamp}] {msg}"
                    SYNC_STATUS[site_id]['message'] = msg
                    SYNC_STATUS[site_id]['logs'].append(log_entry)
                    print(log_entry) # Keep console logging for debug

                result = sync_utils.sync_site(
                    site['url'], 
                    site['consumer_key'], 
                    site['consumer_secret'],
                    progress_callback
                )
                
                duration = int((datetime.now() - sync_start_time).total_seconds())
                
                if result['status'] == 'success':
                    # Update last sync time
                    conn = get_db_connection()
                    conn.execute('UPDATE sites SET last_sync = ? WHERE id = ?', 
                                 (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), site_id))
                    conn.commit()
                    conn.close()
                    
                    SYNC_STATUS[site_id]['status'] = 'success'
                    SYNC_STATUS[site_id]['message'] = 'Synchronization completed successfully'
                    SYNC_STATUS[site_id]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Finished: New {result['new_orders']}, Updated {result['updated_orders']}")
                    
                    # Save sync log to database
                    save_sync_log(site_id, site['url'], 'success', 
                                  f"New: {result['new_orders']}, Updated: {result['updated_orders']}", 
                                  result['new_orders'], result['updated_orders'], duration)
                else:
                    SYNC_STATUS[site_id]['status'] = 'error'
                    SYNC_STATUS[site_id]['message'] = result.get('message', 'Unknown error')
                    SYNC_STATUS[site_id]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {result.get('message')}")
                    
                    # Save error log to database
                    save_sync_log(site_id, site['url'], 'error', result.get('message', 'Unknown error'), 0, 0, duration)
                    
            except Exception as e:
                duration = int((datetime.now() - sync_start_time).total_seconds())
                SYNC_STATUS[site_id]['status'] = 'error'
                SYNC_STATUS[site_id]['message'] = str(e)
                SYNC_STATUS[site_id]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Critical Error: {str(e)}")
                
                # Save error log to database
                try:
                    save_sync_log(site_id, '', 'error', str(e), 0, 0, duration)
                except:
                    pass

    # Run sync in background thread
    thread = threading.Thread(target=run_sync, args=(app.app_context(), site_id))
    thread.start()
    
    return jsonify({'success': True, 'message': 'Synchronization started'})

@app.route('/api/sync/all', methods=['POST'])
@login_required
def sync_all_data():
    """Trigger data synchronization for ALL sites"""
    import sync_utils
    import threading
    
    # Use a special ID for "all sites" sync status
    ALL_SITES_ID = 999999
    
    # Initialize status
    SYNC_STATUS[ALL_SITES_ID] = {
        'status': 'running',
        'message': 'Starting global synchronization...',
        'logs': [f"[{datetime.now().strftime('%H:%M:%S')}] Global sync job started"]
    }
    
    def run_sync_all(app_context):
        with app_context:
            try:
                conn = get_db_connection()
                sites = conn.execute('SELECT * FROM sites').fetchall()
                conn.close()
                
                if not sites:
                    SYNC_STATUS[ALL_SITES_ID]['status'] = 'error'
                    SYNC_STATUS[ALL_SITES_ID]['message'] = 'No sites found to sync'
                    return

                total_sites = len(sites)
                SYNC_STATUS[ALL_SITES_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Found {total_sites} sites to sync")

                for index, site in enumerate(sites):
                    site_id = site['id']
                    site_url = site['url']
                    current_step = index + 1
                    
                    msg = f"Syncing site {current_step}/{total_sites}: {site_url}"
                    SYNC_STATUS[ALL_SITES_ID]['message'] = msg
                    SYNC_STATUS[ALL_SITES_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] --- Starting {site_url} ---")
                    
                    def progress_callback(msg):
                        timestamp = datetime.now().strftime('%H:%M:%S')
                        # Prefix log with site info
                        log_entry = f"[{timestamp}] [{site_url}] {msg}"
                        SYNC_STATUS[ALL_SITES_ID]['logs'].append(log_entry)
                        # Only update main message if it's significant, otherwise keep "Syncing site X/Y"
                        # Actually, let's update the message to show detail
                        SYNC_STATUS[ALL_SITES_ID]['message'] = f"[{current_step}/{total_sites}] {site_url}: {msg}"

                    result = sync_utils.sync_site(
                        site['url'], 
                        site['consumer_key'], 
                        site['consumer_secret'],
                        progress_callback
                    )
                    
                    if result['status'] == 'success':
                        # Update last sync time
                        conn = get_db_connection()
                        conn.execute('UPDATE sites SET last_sync = ? WHERE id = ?', 
                                     (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), site_id))
                        conn.commit()
                        conn.close()
                        SYNC_STATUS[ALL_SITES_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] {site_url} Completed")
                    else:
                        SYNC_STATUS[ALL_SITES_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] {site_url} Failed: {result.get('message')}")
                
                SYNC_STATUS[ALL_SITES_ID]['status'] = 'success'
                SYNC_STATUS[ALL_SITES_ID]['message'] = 'All sites synchronized successfully'
                SYNC_STATUS[ALL_SITES_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Global sync finished")
                    
            except Exception as e:
                SYNC_STATUS[ALL_SITES_ID]['status'] = 'error'
                SYNC_STATUS[ALL_SITES_ID]['message'] = str(e)
                SYNC_STATUS[ALL_SITES_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Critical Error: {str(e)}")

    # Run sync in background thread
    thread = threading.Thread(target=run_sync_all, args=(app.app_context(),))
    thread.start()
    
    return jsonify({'success': True, 'message': 'Global synchronization started', 'sync_id': ALL_SITES_ID})


@app.route('/api/settings/autosync', methods=['GET'])
@login_required
def get_autosync_status():
    """Get autosync status from database and verify cron status"""
    conn = get_db_connection()
    enabled_row = conn.execute("SELECT value FROM settings WHERE key = 'autosync_enabled'").fetchone()
    interval_row = conn.execute("SELECT value FROM settings WHERE key = 'autosync_interval'").fetchone()
    conn.close()
    
    enabled = enabled_row['value'] == 'true' if enabled_row else False
    interval = int(interval_row['value']) if interval_row else 900
    
    # Verify cron job exists if enabled
    if enabled:
        import subprocess
        try:
            result = subprocess.run(['crontab', '-l'], capture_output=True, text=True)
            if 'auto_sync.py' not in result.stdout:
                enabled = False  # Cron was removed externally
        except:
            pass
    
    return jsonify({
        'enabled': enabled,
        'interval': interval
    })

@app.route('/api/settings/autosync', methods=['POST'])
@login_required
def set_autosync_status():
    """Set autosync status, save to database and manage cron job"""
    import subprocess
    
    data = request.json
    conn = get_db_connection()
    
    # Get current values
    enabled_row = conn.execute("SELECT value FROM settings WHERE key = 'autosync_enabled'").fetchone()
    interval_row = conn.execute("SELECT value FROM settings WHERE key = 'autosync_interval'").fetchone()
    
    current_enabled = enabled_row['value'] == 'true' if enabled_row else False
    current_interval = int(interval_row['value']) if interval_row else 900
    
    new_enabled = bool(data.get('enabled', current_enabled))
    new_interval = current_interval
    
    if 'interval' in data:
        interval = int(data['interval'])
        if 300 <= interval <= 86400:  # 5 min to 24 hours
            new_interval = interval
    
    # Save to database
    conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('autosync_enabled', ?)", 
                 ('true' if new_enabled else 'false',))
    conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('autosync_interval', ?)", 
                 (str(new_interval),))
    conn.commit()
    conn.close()
    
    # Manage cron job
    try:
        # Get existing crontab
        result = subprocess.run(['crontab', '-l'], capture_output=True, text=True)
        existing_crontab = result.stdout if result.returncode == 0 else ''
        
        # Remove any existing auto_sync.py entries
        lines = [line for line in existing_crontab.split('\n') 
                 if line.strip() and 'auto_sync.py' not in line]
        
        if new_enabled:
            # Convert interval (seconds) to cron expression
            interval_mins = new_interval // 60
            
            if interval_mins < 60:
                # Every X minutes
                cron_schedule = f"*/{interval_mins} * * * *"
            elif interval_mins < 1440:
                # Every X hours
                hours = interval_mins // 60
                cron_schedule = f"0 */{hours} * * *"
            else:
                # Once a day at 6am
                cron_schedule = "0 6 * * *"
            
            # Add new cron job
            script_path = '/www/wwwroot/woo-analysis'
            cron_line = f"{cron_schedule} cd {script_path} && {script_path}/venv/bin/python auto_sync.py >> {script_path}/auto_sync.log 2>&1"
            lines.append(cron_line)
        
        # Write new crontab
        new_crontab = '\n'.join(lines) + '\n' if lines else ''
        process = subprocess.Popen(['crontab', '-'], stdin=subprocess.PIPE, text=True)
        process.communicate(input=new_crontab)
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'enabled': new_enabled, 'interval': new_interval}), 500
    
    return jsonify({'success': True, 'enabled': new_enabled, 'interval': new_interval})


@app.route('/api/sync/logs')
@login_required
def get_sync_logs():
    """Get synchronization logs"""
    site_id = request.args.get('site_id', '')
    limit = int(request.args.get('limit', 50))
    
    conn = get_db_connection()
    
    if site_id:
        logs = conn.execute('''
            SELECT * FROM sync_logs 
            WHERE site_id = ? 
            ORDER BY sync_time DESC 
            LIMIT ?
        ''', (site_id, limit)).fetchall()
    else:
        logs = conn.execute('''
            SELECT * FROM sync_logs 
            ORDER BY sync_time DESC 
            LIMIT ?
        ''', (limit,)).fetchall()
    
    conn.close()
    
    return jsonify([dict(row) for row in logs])


@app.route('/api/sync/summary')
@login_required
def get_sync_summary():
    """Get sync summary for all sites"""
    conn = get_db_connection()
    
    # Get all sites with their latest sync info
    sites = conn.execute('SELECT * FROM sites').fetchall()
    
    summary = []
    for site in sites:
        # Get latest log for this site
        latest_log = conn.execute('''
            SELECT * FROM sync_logs 
            WHERE site_id = ? 
            ORDER BY sync_time DESC 
            LIMIT 1
        ''', (site['id'],)).fetchone()
        
        # Get stats for last 7 days
        stats = conn.execute('''
            SELECT 
                COUNT(*) as total_syncs,
                SUM(CASE WHEN status = 'success' THEN 1 ELSE 0 END) as success_count,
                SUM(new_orders) as total_new_orders,
                SUM(updated_orders) as total_updated_orders,
                AVG(duration_seconds) as avg_duration
            FROM sync_logs 
            WHERE site_id = ? AND sync_time >= datetime('now', '-7 days')
        ''', (site['id'],)).fetchone()
        
        summary.append({
            'site_id': site['id'],
            'url': site['url'],
            'last_sync': site['last_sync'],
            'latest_log': dict(latest_log) if latest_log else None,
            'stats_7_days': dict(stats) if stats else None
        })
    
    conn.close()
    
    return jsonify(summary)


@app.route('/api/sync/deep', methods=['POST'])
@login_required
def trigger_deep_sync():
    """Trigger deep sync using 1.wooorders_sqlite.py script"""
    import subprocess
    import threading
    
    DEEP_SYNC_ID = 888888
    
    SYNC_STATUS[DEEP_SYNC_ID] = {
        'status': 'running',
        'message': '正在启动深度同步...',
        'logs': [f"[{datetime.now().strftime('%H:%M:%S')}] Deep sync job started"]
    }
    
    def run_deep_sync(app_context):
        with app_context:
            try:
                script_path = '/www/wwwroot/woo-analysis/1.wooorders_sqlite.py'
                venv_python = '/www/wwwroot/woo-analysis/venv/bin/python'
                
                SYNC_STATUS[DEEP_SYNC_ID]['message'] = '正在执行深度同步脚本...'
                SYNC_STATUS[DEEP_SYNC_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Running {script_path}")
                
                result = subprocess.run(
                    [venv_python, script_path],
                    cwd='/www/wwwroot/woo-analysis',
                    capture_output=True,
                    text=True,
                    timeout=3600  # 1 hour timeout
                )
                
                if result.returncode == 0:
                    SYNC_STATUS[DEEP_SYNC_ID]['status'] = 'success'
                    SYNC_STATUS[DEEP_SYNC_ID]['message'] = '深度同步完成'
                    SYNC_STATUS[DEEP_SYNC_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Completed successfully")
                    # Add last few lines of output
                    output_lines = result.stdout.strip().split('\n')[-10:]
                    for line in output_lines:
                        SYNC_STATUS[DEEP_SYNC_ID]['logs'].append(line)
                else:
                    SYNC_STATUS[DEEP_SYNC_ID]['status'] = 'error'
                    SYNC_STATUS[DEEP_SYNC_ID]['message'] = '深度同步失败'
                    SYNC_STATUS[DEEP_SYNC_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {result.stderr[:500]}")
                    
            except subprocess.TimeoutExpired:
                SYNC_STATUS[DEEP_SYNC_ID]['status'] = 'error'
                SYNC_STATUS[DEEP_SYNC_ID]['message'] = '深度同步超时'
            except Exception as e:
                SYNC_STATUS[DEEP_SYNC_ID]['status'] = 'error'
                SYNC_STATUS[DEEP_SYNC_ID]['message'] = str(e)
    
    thread = threading.Thread(target=run_deep_sync, args=(app.app_context(),))
    thread.start()
    
    return jsonify({'success': True, 'sync_id': DEEP_SYNC_ID, 'message': 'Deep sync started'})


@app.route('/api/cron/status')
@login_required
def get_cron_status():
    """Get current cron job status for deep sync"""
    import subprocess
    
    try:
        result = subprocess.run(['crontab', '-l'], capture_output=True, text=True)
        crontab_content = result.stdout
        
        # Check if our deep sync job exists
        job_pattern = '1.wooorders_sqlite.py'
        has_job = job_pattern in crontab_content
        
        # Extract schedule if exists
        schedule = None
        if has_job:
            for line in crontab_content.split('\n'):
                if job_pattern in line and not line.startswith('#'):
                    parts = line.split()
                    if len(parts) >= 5:
                        schedule = ' '.join(parts[:5])
                    break
        
        return jsonify({
            'enabled': has_job,
            'schedule': schedule,
            'raw': crontab_content if has_job else None
        })
    except Exception as e:
        return jsonify({'enabled': False, 'error': str(e)})


@app.route('/api/cron/setup', methods=['POST'])
@login_required
def setup_cron():
    """Setup cron job for deep sync"""
    import subprocess
    
    data = request.json
    hour = int(data.get('hour', 3))  # Default 3 AM
    minute = int(data.get('minute', 0))
    
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return jsonify({'error': 'Invalid time'}), 400
    
    try:
        # Get existing crontab
        result = subprocess.run(['crontab', '-l'], capture_output=True, text=True)
        existing_crontab = result.stdout if result.returncode == 0 else ''
        
        # Remove existing deep sync job if any
        lines = [line for line in existing_crontab.split('\n') 
                 if '1.wooorders_sqlite.py' not in line and line.strip()]
        
        # Add new job
        new_job = f"{minute} {hour} * * * cd /www/wwwroot/woo-analysis && /www/wwwroot/woo-analysis/venv/bin/python 1.wooorders_sqlite.py >> /www/wwwroot/woo-analysis/deep_sync.log 2>&1"
        lines.append(new_job)
        
        # Write new crontab
        new_crontab = '\n'.join(lines) + '\n'
        process = subprocess.Popen(['crontab', '-'], stdin=subprocess.PIPE, text=True)
        process.communicate(input=new_crontab)
        
        if process.returncode == 0:
            return jsonify({'success': True, 'message': f'Cron job set for {hour:02d}:{minute:02d}'})
        else:
            return jsonify({'error': 'Failed to set crontab'}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/cron/remove', methods=['DELETE'])
@login_required
def remove_cron():
    """Remove cron job for deep sync"""
    import subprocess
    
    try:
        result = subprocess.run(['crontab', '-l'], capture_output=True, text=True)
        existing_crontab = result.stdout if result.returncode == 0 else ''
        
        # Remove deep sync job
        lines = [line for line in existing_crontab.split('\n') 
                 if '1.wooorders_sqlite.py' not in line and line.strip()]
        
        new_crontab = '\n'.join(lines) + '\n' if lines else ''
        process = subprocess.Popen(['crontab', '-'], stdin=subprocess.PIPE, text=True)
        process.communicate(input=new_crontab)
        
        return jsonify({'success': True, 'message': 'Cron job removed'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ============== USER MANAGEMENT API ==============


@app.route('/users')
@login_required
@admin_required
def users_page():
    """User management page"""
    return render_template('users.html')


@app.route('/api/users')
@login_required
@admin_required
def get_users():
    """Get all users"""
    conn = get_db_connection()
    users = conn.execute('SELECT id, username, name, role, created_at FROM users').fetchall()
    conn.close()
    return jsonify([dict(row) for row in users])


@app.route('/api/users', methods=['POST'])
@login_required
@admin_required
def add_user():
    """Add a new user"""
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '')
    name = data.get('name', '').strip()
    role = data.get('role', 'user')
    
    if not username or not password:
        return jsonify({'error': 'Username and password required'}), 400
    
    if role not in ['admin', 'user']:
        role = 'user'
    
    conn = get_db_connection()
    try:
        conn.execute('''
            INSERT INTO users (username, password_hash, name, role)
            VALUES (?, ?, ?, ?)
        ''', (username, generate_password_hash(password), name, role))
        conn.commit()
        return jsonify({'success': True})
    except sqlite3.IntegrityError:
        return jsonify({'error': 'Username already exists'}), 400
    finally:
        conn.close()


@app.route('/api/users/<int:user_id>', methods=['PUT'])
@login_required
@admin_required
def update_user(user_id):
    """Update a user"""
    data = request.json
    name = data.get('name', '').strip()
    role = data.get('role', 'user')
    password = data.get('password', '')
    
    if role not in ['admin', 'user']:
        role = 'user'
    
    conn = get_db_connection()
    try:
        if password:
            conn.execute('UPDATE users SET name = ?, role = ?, password_hash = ? WHERE id = ?',
                        (name, role, generate_password_hash(password), user_id))
        else:
            conn.execute('UPDATE users SET name = ?, role = ? WHERE id = ?',
                        (name, role, user_id))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


@app.route('/api/users/<int:user_id>', methods=['DELETE'])
@login_required
@admin_required
def delete_user(user_id):
    """Delete a user"""
    # Prevent deleting yourself
    if str(user_id) == str(current_user.id):
        return jsonify({'error': 'Cannot delete yourself'}), 400
    
    conn = get_db_connection()
    try:
        conn.execute('DELETE FROM user_site_permissions WHERE user_id = ?', (user_id,))
        conn.execute('DELETE FROM users WHERE id = ?', (user_id,))
        conn.commit()
        return jsonify({'success': True})
    finally:
        conn.close()


@app.route('/api/users/change-password', methods=['POST'])
@login_required
def change_own_password():
    """Change current user's password"""
    data = request.json
    old_password = data.get('old_password', '')
    new_password = data.get('new_password', '')
    
    if not old_password or not new_password:
        return jsonify({'error': 'Both old and new password required'}), 400
    
    conn = get_db_connection()
    user = conn.execute('SELECT * FROM users WHERE id = ?', (current_user.id,)).fetchone()
    
    if not check_password_hash(user['password_hash'], old_password):
        conn.close()
        return jsonify({'error': 'Incorrect old password'}), 400
    
    conn.execute('UPDATE users SET password_hash = ? WHERE id = ?',
                (generate_password_hash(new_password), current_user.id))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True, 'message': 'Password changed successfully'})


@app.route('/api/users/<int:user_id>/permissions')
@login_required
@admin_required
def get_user_permissions(user_id):
    """Get user's site permissions"""
    conn = get_db_connection()
    permissions = conn.execute('''
        SELECT site_id FROM user_site_permissions WHERE user_id = ?
    ''', (user_id,)).fetchall()
    sites = conn.execute('SELECT id, url FROM sites').fetchall()
    conn.close()
    
    return jsonify({
        'allowed_sites': [p['site_id'] for p in permissions],
        'all_sites': [dict(s) for s in sites]
    })


@app.route('/api/users/<int:user_id>/permissions', methods=['PUT'])
@login_required
@admin_required
def update_user_permissions(user_id):
    """Update user's site permissions"""
    data = request.json
    site_ids = data.get('site_ids', [])
    
    conn = get_db_connection()
    # Clear existing permissions
    conn.execute('DELETE FROM user_site_permissions WHERE user_id = ?', (user_id,))
    
    # Add new permissions
    for site_id in site_ids:
        conn.execute('INSERT INTO user_site_permissions (user_id, site_id) VALUES (?, ?)',
                    (user_id, site_id))
    
    conn.commit()
    conn.close()
    
    return jsonify({'success': True})


# ============== PRODUCT ANALYSIS ==============

@app.route('/products')
@login_required
def products():
    """Product analysis page"""
    from datetime import date, timedelta
    conn = get_db_connection()
    
    # Get user's allowed sources for permission filtering
    allowed_sources = get_user_allowed_sources(current_user.id, current_user.is_admin())
    
    # Get filter parameters
    source_filter = request.args.get('source', '')
    brand_filter = request.args.get('brand', '')
    puff_filter = request.args.get('puffs', '')
    quick_date = request.args.get('quick_date', 'this_month')
    manager_filter = request.args.get('manager', '')
    
    # Get all managers
    all_managers = get_all_managers()
    
    # Validate source_filter against allowed sources
    if source_filter and allowed_sources is not None and source_filter not in allowed_sources:
        source_filter = ''
    
    # Process quick date filter
    today = date.today()
    date_from = ''
    date_to = today.isoformat()
    
    if quick_date == 'this_month':
        date_from = today.replace(day=1).isoformat()
    elif quick_date == 'last_month':
        first_of_this_month = today.replace(day=1)
        last_month_end = first_of_this_month - timedelta(days=1)
        last_month_start = last_month_end.replace(day=1)
        date_from = last_month_start.isoformat()
        date_to = last_month_end.isoformat()
    elif quick_date == 'last_quarter':
        date_from = (today - timedelta(days=90)).isoformat()
    elif quick_date == 'half_year':
        date_from = (today - timedelta(days=180)).isoformat()
    elif quick_date == 'one_year':
        date_from = (today - timedelta(days=365)).isoformat()
    elif quick_date == 'all':
        date_from = ''
        date_to = ''
    
    # Build filter conditions with permission check
    conditions = ["status NOT IN ('failed', 'cancelled')"]
    params = []
    
    # Add permission filter
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            conditions.append(f'source IN ({placeholders})')
            params.extend(allowed_sources)
        else:
            conditions.append('1=0')
            
    # Add manager filter
    if manager_filter:
        # Re-use manager_urls from above or fetch if not done
        if 'manager_urls' not in locals():
            manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
            manager_urls = [s['url'] for s in manager_sites]
        
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            conditions.append(f'source IN ({placeholders})')
            params.extend(manager_urls)
        else:
            conditions.append('1=0')
    
    if date_from and date_to:
        conditions.append("date_created >= ? AND date_created <= ?")
        params.extend([date_from, date_to + 'T23:59:59'])
    elif date_from:
        conditions.append("date_created >= ?")
        params.append(date_from)
    
    if source_filter:
        conditions.append("source = ?")
        params.append(source_filter)
    
    where_clause = 'WHERE ' + ' AND '.join(conditions)
    
    # Get all orders with line items (including currency and shipping)
    orders = conn.execute(f'''
        SELECT id, line_items, source, currency, total, shipping_total, date_created
        FROM orders {where_clause}
    ''', params).fetchall()
    
    # Get brands list for parsing and filtering
    brands = conn.execute('SELECT id, name, aliases FROM brands ORDER BY name').fetchall()
    brands_list = [dict(b) for b in brands]
    
    # Build brands cache for parsing
    brands_cache = []
    for row in brands:
        brand_name = row['name']
        aliases = []
        if row['aliases']:
            try:
                aliases = json.loads(row['aliases'])
            except:
                pass
        brands_cache.append({
            'id': row['id'],
            'name': brand_name,
            'aliases': aliases,
            'patterns': [brand_name.upper()] + [a.upper() for a in aliases]
        })
    
    # Load manual product mappings
    mappings_rows = conn.execute('''
        SELECT pm.raw_name, pm.puff_count, pm.flavor, b.name as brand_name
        FROM product_mappings pm
        LEFT JOIN brands b ON pm.brand_id = b.id
        WHERE pm.is_manual = 1
    ''').fetchall()
    manual_mappings = {}
    for m in mappings_rows:
        manual_mappings[m['raw_name']] = {
            'brand': m['brand_name'],
            'puffs': m['puff_count'],
            'flavor': m['flavor']
        }
    
    # Aggregate product data
    product_stats = {}
    brand_stats = {}
    puff_stats = {}
    
    for order in orders:
        items = parse_json_field(order['line_items'])
        if not isinstance(items, list):
            continue
        
        # Calculate order items sum for shipping pro-rating
        order_items_sum = sum(float(i.get('total', 0)) for i in items)
        order_shipping = float(order['shipping_total'] or 0)
        
        for item in items:
            product_name = item.get('name', '')
            quantity = item.get('quantity', 0)
            total = float(item.get('total', 0))
            
            # Pro-rate shipping
            item_shipping = 0
            if order_items_sum > 0:
                item_shipping = order_shipping * (total / order_items_sum)
            
            gross_total = total + item_shipping
            
            # Extract flavor from WooCommerce variation meta_data
            meta_flavor = extract_flavor_from_meta(item)
            
            # Check for manual mapping first (using full name with flavor)
            full_name, _, meta_puffs = get_full_product_name(item)
            if full_name in manual_mappings:
                mapping = manual_mappings[full_name]
                brand = mapping.get('brand') or 'Unknown'
                # Use meta_puffs if mapping doesn't have puffs
                puffs = mapping.get('puffs') or meta_puffs
                flavor = mapping.get('flavor') or meta_flavor or ''
            elif product_name in manual_mappings:
                mapping = manual_mappings[product_name]
                brand = mapping.get('brand') or 'Unknown'
                # Use meta_puffs if mapping doesn't have puffs
                puffs = mapping.get('puffs') or meta_puffs
                # Use meta_flavor if mapping doesn't have flavor
                flavor = mapping.get('flavor') or meta_flavor or ''
            else:
                # Parse product name automatically
                parsed = parse_product_name(product_name, brands_cache)
                brand = parsed.get('brand') or 'Unknown'
                # Prefer meta_puffs from WooCommerce variation data, then parsed puffs
                puffs = meta_puffs or parsed.get('puffs')
                # Prefer meta_flavor from WooCommerce variation data
                flavor = meta_flavor or parsed.get('flavor') or ''
            
            # Apply brand filter
            if brand_filter and brand != brand_filter:
                continue
            
            # Apply puff filter
            if puff_filter and puffs and str(puffs) != puff_filter:
                continue
            
            # Product level stats (brand + puffs + flavor)
            product_key = f"{brand}|{puffs or 'N/A'}|{flavor or 'No Flavor'}"
            if product_key not in product_stats:
                product_stats[product_key] = {
                    'name': product_name,  # Store one sample name for mapping
                    'brand': brand,
                    'puffs': puffs,
                    'flavor': flavor,
                    'quantity': 0,
                    'revenue_by_currency': {},
                    'gross_revenue_by_currency': {},
                    'order_count': 0
                }
            product_stats[product_key]['quantity'] += quantity
            # Track revenue by currency
            currency = order['currency'] or 'N/A'
            if currency not in product_stats[product_key]['revenue_by_currency']:
                product_stats[product_key]['revenue_by_currency'][currency] = 0
                product_stats[product_key]['gross_revenue_by_currency'][currency] = 0
            product_stats[product_key]['revenue_by_currency'][currency] += total
            product_stats[product_key]['gross_revenue_by_currency'][currency] += gross_total
            product_stats[product_key]['order_count'] += 1
            
            # Brand level stats
            if brand not in brand_stats:
                brand_stats[brand] = {'quantity': 0, 'revenue_by_currency': {}, 'gross_revenue_by_currency': {}, 'order_count': 0}
            brand_stats[brand]['quantity'] += quantity
            if currency not in brand_stats[brand]['revenue_by_currency']:
                brand_stats[brand]['revenue_by_currency'][currency] = 0
                brand_stats[brand]['gross_revenue_by_currency'][currency] = 0
            brand_stats[brand]['revenue_by_currency'][currency] += total
            brand_stats[brand]['gross_revenue_by_currency'][currency] += gross_total
            brand_stats[brand]['order_count'] += 1
            
            # Puff level stats
            puff_key = str(puffs) if puffs else 'Unknown'
            if puff_key not in puff_stats:
                puff_stats[puff_key] = {'quantity': 0, 'revenue_by_currency': {}, 'gross_revenue_by_currency': {}}
            puff_stats[puff_key]['quantity'] += quantity
            if currency not in puff_stats[puff_key]['revenue_by_currency']:
                puff_stats[puff_key]['revenue_by_currency'][currency] = 0
                puff_stats[puff_key]['gross_revenue_by_currency'][currency] = 0
            puff_stats[puff_key]['revenue_by_currency'][currency] += total
            puff_stats[puff_key]['gross_revenue_by_currency'][currency] += gross_total
    
    # Sort products by quantity (top sellers)
    top_products = sorted(product_stats.values(), key=lambda x: x['quantity'], reverse=True)[:50]
    
    # Sort brands by quantity
    brand_ranking = sorted(
        [{'name': k, **v} for k, v in brand_stats.items()],
        key=lambda x: x['quantity'],
        reverse=True
    )
    
    # Sort puffs by quantity
    puff_ranking = sorted(
        [{'puffs': k, **v} for k, v in puff_stats.items()],
        key=lambda x: x['quantity'],
        reverse=True
    )
    
    # Get available sources (filtered by permissions and manager)
    conn = get_db_connection()
    
    # Base source query
    source_query = 'SELECT DISTINCT source FROM orders'
    source_params = []
    source_conditions = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(allowed_sources)
        else:
            source_conditions.append('1=0')
            
    if manager_filter:
        # Get sites managed by this manager
        manager_sites = conn.execute('SELECT url FROM sites WHERE manager = ?', (manager_filter,)).fetchall()
        manager_urls = [s['url'] for s in manager_sites]
        if manager_urls:
            placeholders = ', '.join(['?' for _ in manager_urls])
            source_conditions.append(f'source IN ({placeholders})')
            source_params.extend(manager_urls)
        else:
            source_conditions.append('1=0')
            
    if source_conditions:
        source_query += ' WHERE ' + ' AND '.join(source_conditions)
        
    source_query += ' ORDER BY source'
    sources = conn.execute(source_query, source_params).fetchall()
    
    # Get available puff counts
    puff_options = sorted([p for p in puff_stats.keys() if p != 'Unknown'], key=lambda x: int(x) if x.isdigit() else 0)
    
    # Get site managers mapping
    sites = conn.execute('SELECT url, manager FROM sites').fetchall()
    site_managers = {s['url']: s['manager'] or '' for s in sites}
    
    # Calculate totals with revenue by currency
    total_revenue_by_currency = {}
    total_gross_revenue_by_currency = {}
    for p in product_stats.values():
        for currency, amount in p.get('revenue_by_currency', {}).items():
            if currency not in total_revenue_by_currency:
                total_revenue_by_currency[currency] = 0
            total_revenue_by_currency[currency] += amount
            
        for currency, amount in p.get('gross_revenue_by_currency', {}).items():
            if currency not in total_gross_revenue_by_currency:
                total_gross_revenue_by_currency[currency] = 0
            total_gross_revenue_by_currency[currency] += amount
    
    # Calculate CNY total for products
    total_revenue_cny = 0
    total_gross_revenue_cny = 0
    from datetime import datetime
    current_month = datetime.now().strftime('%Y-%m')
    
    for currency, amount in total_revenue_by_currency.items():
        rate, _ = get_cny_rate(currency, current_month)
        if rate:
            total_revenue_cny += amount * rate
            
    for currency, amount in total_gross_revenue_by_currency.items():
        rate, _ = get_cny_rate(currency, current_month)
        if rate:
            total_gross_revenue_cny += amount * rate
    
    totals = {
        'total_quantity': sum(p['quantity'] for p in product_stats.values()),
        'total_revenue_by_currency': total_revenue_by_currency,
        'total_gross_revenue_by_currency': total_gross_revenue_by_currency,
        'total_revenue_cny': round(total_revenue_cny, 2),
        'total_gross_revenue_cny': round(total_gross_revenue_cny, 2),
        'brand_count': len(brand_stats),
        'product_count': len(product_stats)
    }
    
    conn.close()
    
    return render_template('products.html',
                          top_products=top_products,
                          brand_ranking=brand_ranking,
                          puff_ranking=puff_ranking,
                          totals=totals,
                          brands=brands_list,
                          sources=sources,
                          puff_options=puff_options,
                          site_managers=site_managers,
                          current_filters={
                              'source': source_filter,
                              'brand': brand_filter,
                              'puffs': puff_filter,
                              'quick_date': quick_date,
                              'manager': manager_filter
                          },
                          all_managers=all_managers)


@app.route('/api/brands')
@login_required
def get_brands():
    """Get all brands"""
    conn = get_db_connection()
    brands = conn.execute('SELECT * FROM brands ORDER BY name').fetchall()
    conn.close()
    
    result = []
    for b in brands:
        aliases = []
        if b['aliases']:
            try:
                aliases = json.loads(b['aliases'])
            except:
                pass
        result.append({
            'id': b['id'],
            'name': b['name'],
            'aliases': aliases,
            'created_at': b['created_at']
        })
    
    return jsonify(result)


@app.route('/api/brands', methods=['POST'])
@login_required
@admin_required
def add_brand():
    """Add a new brand"""
    data = request.json
    name = data.get('name', '').strip()
    aliases = data.get('aliases', [])
    
    if not name:
        return jsonify({'error': 'Brand name is required'}), 400
    
    conn = get_db_connection()
    try:
        conn.execute('INSERT INTO brands (name, aliases) VALUES (?, ?)',
                    (name, json.dumps(aliases) if aliases else None))
        conn.commit()
        brand_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
        conn.close()
        return jsonify({'success': True, 'id': brand_id})
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({'error': 'Brand already exists'}), 400


@app.route('/api/brands/<int:brand_id>', methods=['PUT'])
@login_required
@admin_required
def update_brand(brand_id):
    """Update a brand"""
    data = request.json
    name = data.get('name', '').strip()
    aliases = data.get('aliases', [])
    
    if not name:
        return jsonify({'error': 'Brand name is required'}), 400
    
    conn = get_db_connection()
    conn.execute('UPDATE brands SET name = ?, aliases = ? WHERE id = ?',
                (name, json.dumps(aliases) if aliases else None, brand_id))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True})


@app.route('/api/brands/<int:brand_id>', methods=['DELETE'])
@login_required
@admin_required
def delete_brand(brand_id):
    """Delete a brand"""
    conn = get_db_connection()
    conn.execute('DELETE FROM brands WHERE id = ?', (brand_id,))
    conn.execute('UPDATE product_mappings SET brand_id = NULL WHERE brand_id = ?', (brand_id,))
    conn.commit()
    conn.close()
    
    return jsonify({'success': True})


@app.route('/api/products/stats')
@login_required
def get_product_stats():
    """Get product statistics API"""
    from datetime import date, timedelta
    
    conn = get_db_connection()
    allowed_sources = get_user_allowed_sources(current_user.id, current_user.is_admin())
    
    # Get parameters
    source = request.args.get('source', '')
    days = int(request.args.get('days', 30))
    
    # Build conditions
    conditions = ["status NOT IN ('failed', 'cancelled')"]
    params = []
    
    if allowed_sources is not None:
        if allowed_sources:
            placeholders = ', '.join(['?' for _ in allowed_sources])
            conditions.append(f'source IN ({placeholders})')
            params.extend(allowed_sources)
        else:
            conditions.append('1=0')
    
    if source:
        conditions.append('source = ?')
        params.append(source)
    
    date_from = (date.today() - timedelta(days=days)).isoformat()
    conditions.append('date_created >= ?')
    params.append(date_from)
    
    where_clause = 'WHERE ' + ' AND '.join(conditions)
    
    orders = conn.execute(f'''
        SELECT line_items FROM orders {where_clause}
    ''', params).fetchall()
    
    # Get brands cache
    brands_rows = conn.execute('SELECT id, name, aliases FROM brands').fetchall()
    brands_cache = []
    for row in brands_rows:
        brand_name = row['name']
        aliases = []
        if row['aliases']:
            try:
                aliases = json.loads(row['aliases'])
            except:
                pass
        brands_cache.append({
            'id': row['id'],
            'name': brand_name,
            'aliases': aliases,
            'patterns': [brand_name.upper()] + [a.upper() for a in aliases]
        })
    
    conn.close()
    
    # Aggregate
    stats = {}
    for order in orders:
        items = parse_json_field(order['line_items'])
        if not isinstance(items, list):
            continue
        for item in items:
            name = item.get('name', '')
            qty = item.get('quantity', 0)
            total = float(item.get('total', 0))
            
            # Extract flavor from meta_data
            meta_flavor = extract_flavor_from_meta(item)
            
            parsed = parse_product_name(name, brands_cache)
            # Use meta_flavor if available, otherwise use parsed flavor
            flavor = meta_flavor or parsed.get('flavor') or ''
            
            # Create key with flavor for proper aggregation
            base_key = parsed.get('normalized') or name
            key = f"{base_key} - {flavor}" if flavor and flavor.upper() not in base_key.upper() else base_key
            
            if key not in stats:
                stats[key] = {
                    'name': key,
                    'brand': parsed.get('brand'),
                    'puffs': parsed.get('puffs'),
                    'flavor': flavor,
                    'quantity': 0,
                    'revenue': 0
                }
            stats[key]['quantity'] += qty
            stats[key]['revenue'] += total
    
    result = sorted(stats.values(), key=lambda x: x['quantity'], reverse=True)[:100]
    return jsonify(result)


@app.route('/api/products/unknown')
@login_required
def get_unknown_products():
    """Get products that could not be mapped to any brand"""
    from datetime import date, timedelta
    
    conn = get_db_connection()
    
    # Get parameters
    days = int(request.args.get('days', 90))
    limit = int(request.args.get('limit', 50))
    
    # Get orders from last N days
    date_from = (date.today() - timedelta(days=days)).isoformat()
    
    orders = conn.execute('''
        SELECT line_items, source FROM orders 
        WHERE date_created >= ? AND status NOT IN ('failed', 'cancelled')
    ''', (date_from,)).fetchall()
    
    # Get brands cache
    brands_rows = conn.execute('SELECT id, name, aliases FROM brands').fetchall()
    brands_cache = []
    for row in brands_rows:
        brand_name = row['name']
        aliases = []
        if row['aliases']:
            try:
                aliases = json.loads(row['aliases'])
            except:
                pass
        brands_cache.append({
            'id': row['id'],
            'name': brand_name,
            'aliases': aliases,
            'patterns': [brand_name.upper()] + [a.upper() for a in aliases]
        })
    
    conn.close()
    
    # Find unknown products
    unknown_products = {}
    
    for order in orders:
        items = parse_json_field(order['line_items'])
        if not isinstance(items, list):
            continue
        
        for item in items:
            name = item.get('name', '')
            if not name:
                continue
            
            quantity = item.get('quantity', 0)
            
            # Get full name with flavor from meta_data
            full_name, meta_flavor, meta_puffs = get_full_product_name(item)
            
            # Parse and check if brand is found
            parsed = parse_product_name(name, brands_cache)
            
            if parsed.get('brand') is None:
                # This product is unknown
                # Use a simplified key (remove flavor for grouping)
                key = name.split(' - ')[0].strip() if ' - ' in name else name
                
                if key not in unknown_products:
                    unknown_products[key] = {
                        'name': key,
                        'sample_full_name': full_name,  # Include flavor in sample
                        'puffs': parsed.get('puffs'),
                        'quantity': 0,
                        'sources': set()
                    }
                unknown_products[key]['quantity'] += quantity
                unknown_products[key]['sources'].add(order['source'])
    
    # Convert sets to lists and sort by quantity
    result = []
    for key, product in unknown_products.items():
        product['sources'] = list(product['sources'])
        result.append(product)
    
    result = sorted(result, key=lambda x: x['quantity'], reverse=True)[:limit]
    
    return jsonify(result)


@app.route('/api/products/mapping', methods=['GET', 'POST'])
@login_required
def product_mapping():
    """Get or save product mapping"""
    conn = get_db_connection()
    
    if request.method == 'GET':
        # Get mapping for a specific product name
        raw_name = request.args.get('name', '')
        if not raw_name:
            conn.close()
            return jsonify({'error': 'Name is required'}), 400
        
        mapping = conn.execute('''
            SELECT pm.*, b.name as brand_name 
            FROM product_mappings pm
            LEFT JOIN brands b ON pm.brand_id = b.id
            WHERE pm.raw_name = ?
        ''', (raw_name,)).fetchone()
        
        conn.close()
        
        if mapping:
            return jsonify({
                'id': mapping['id'],
                'raw_name': mapping['raw_name'],
                'brand_id': mapping['brand_id'],
                'brand_name': mapping['brand_name'],
                'puff_count': mapping['puff_count'],
                'flavor': mapping['flavor'],
                'is_manual': mapping['is_manual']
            })
        else:
            return jsonify({'exists': False})
    
    else:  # POST - Save mapping
        data = request.get_json()
        raw_name = data.get('raw_name', '').strip()
        
        if not raw_name:
            conn.close()
            return jsonify({'success': False, 'error': 'Product name is required'})
        
        brand_id = data.get('brand_id')
        puff_count = data.get('puff_count')
        flavor = data.get('flavor')
        
        # Convert to proper types
        if brand_id and str(brand_id).strip():
            try:
                brand_id = int(brand_id)
            except ValueError:
                brand_id = None
        else:
            brand_id = None
            
        if puff_count and str(puff_count).strip():
            try:
                puff_count = int(puff_count)
            except ValueError:
                puff_count = None
        else:
            puff_count = None
            
        if flavor:
            flavor = str(flavor).strip()
        else:
            flavor = ''
        
        try:
            # Check if mapping exists
            existing = conn.execute(
                'SELECT id FROM product_mappings WHERE raw_name = ?', 
                (raw_name,)
            ).fetchone()
            
            if existing:
                # Update existing
                conn.execute('''
                    UPDATE product_mappings 
                    SET brand_id = ?, puff_count = ?, flavor = ?, is_manual = 1
                    WHERE raw_name = ?
                ''', (brand_id, puff_count, flavor, raw_name))
            else:
                # Insert new
                conn.execute('''
                    INSERT INTO product_mappings (raw_name, brand_id, puff_count, flavor, is_manual)
                    VALUES (?, ?, ?, ?, 1)
                ''', (raw_name, brand_id, puff_count, flavor))
            
            conn.commit()
            conn.close()
            return jsonify({'success': True})
        except Exception as e:
            conn.close()
            return jsonify({'success': False, 'error': str(e)})


@app.route('/api/products/samples')
@login_required
def get_product_samples():
    """Get sample orders containing a specific product"""
    product_name = request.args.get('name', '')
    if not product_name:
        return jsonify({'error': 'Product name is required'}), 400
    
    conn = get_db_connection()
    
    # Search for orders containing this product in line_items
    # We need to search JSON, so we'll filter in Python
    orders = conn.execute('''
        SELECT id, number, source, date_created, line_items 
        FROM orders 
        WHERE status NOT IN ('failed', 'cancelled')
        ORDER BY date_created DESC
        LIMIT 500
    ''').fetchall()
    
    conn.close()
    
    # Find orders with matching product name
    results = []
    sources_map = {}
    
    for order in orders:
        items = parse_json_field(order['line_items'])
        if not isinstance(items, list):
            continue
        
        for item in items:
            if item.get('name', '') == product_name:
                # Extract domain from source for display
                source = order['source']
                source_display = source.replace('https://www.', '').replace('https://', '').split('/')[0]
                
                # Get manager
                manager = conn.execute('SELECT manager FROM sites WHERE url = ?', (source,)).fetchone()
                manager_name = manager['manager'] if manager else ''
                
                sources_map[source_display] = manager_name
                
                results.append({
                    'order_number': order['number'],
                    'source': source_display,
                    'manager': manager_name,
                    'date': order['date_created'][:10] if order['date_created'] else ''
                })
                break  # One match per order is enough
        
        if len(results) >= 10:
            break
    
    sources_list = [{'site': k, 'manager': v} for k, v in sources_map.items()]
    
    return jsonify({
        'sources': sources_list,
        'orders': results,
        'total_found': len(results)
    })



@app.route('/api/settings', methods=['GET', 'POST'])
@login_required
def settings_api():
    """API endpoint for general settings"""
    conn = get_db_connection()
    
    if request.method == 'GET':
        settings = conn.execute("SELECT key, value FROM settings").fetchall()
        conn.close()
        return jsonify({row['key']: row['value'] for row in settings})
        
    elif request.method == 'POST':
        data = request.json
        try:
            for key, value in data.items():
                # Allow specific keys only for security
                if key in ['source_display_mode', 'autosync_enabled', 'autosync_interval']:
                    conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value)))
            conn.commit()
            conn.close()
            return jsonify({'success': True})
        except Exception as e:
            conn.close()
            return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
