"""
WooCommerce Order Analysis Web Dashboard
Flask application with user authentication and data visualization
"""
import sqlite3
import json
from datetime import datetime
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session
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
    def __init__(self, user_id, username, name):
        self.id = user_id
        self.username = username
        self.name = name


@login_manager.user_loader
def load_user(user_id):
    for username, user_data in USERS.items():
        if user_data['id'] == user_id:
            return User(user_data['id'], username, user_data['name'])
    return None


def get_db_connection():
    """Create database connection"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def parse_json_field(value):
    """Safely parse JSON field"""
    if not value:
        return {}
    try:
        return json.loads(value)
    except:
        return {}


# ============== ROUTES ==============

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        
        if username in USERS and check_password_hash(USERS[username]['password'], password):
            user_data = USERS[username]
            user = User(user_data['id'], username, user_data['name'])
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
    
    # Get filters
    quick_date = request.args.get('quick_date', 'this_month')
    source_filter = request.args.get('source', '')
    
    # Get available sources for dropdown
    sources = conn.execute('SELECT DISTINCT source FROM orders ORDER BY source').fetchall()
    
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
    
    # Build filter conditions
    conditions = []
    if date_from and date_to:
        conditions.append(f"date_created >= '{date_from}' AND date_created <= '{date_to}T23:59:59'")
    elif date_from:
        conditions.append(f"date_created >= '{date_from}'")
    if source_filter:
        conditions.append(f"source = '{source_filter}'")
    
    date_condition = 'WHERE ' + ' AND '.join(conditions) if conditions else ''
    
    # Get overall statistics
    stats = {}
    
    # Total orders
    stats['total_orders'] = conn.execute(f'SELECT COUNT(*) FROM orders {date_condition}').fetchone()[0]
    
    # Total revenue
    revenue_where = date_condition.replace('WHERE', 'WHERE status NOT IN ("failed", "cancelled") AND') if date_condition else 'WHERE status NOT IN ("failed", "cancelled")'
    result = conn.execute(f'SELECT SUM(total) FROM orders {revenue_where}').fetchone()[0]
    stats['total_revenue'] = result or 0
    
    # Orders by status
    status_data_raw = conn.execute(f'''
        SELECT status, COUNT(*) as count, SUM(total) as revenue
        FROM orders {date_condition}
        GROUP BY status
        ORDER BY count DESC
    ''').fetchall()
    status_data = [dict(row) for row in status_data_raw]
    
    # Orders by source
    source_data_raw = conn.execute(f'''
        SELECT source, COUNT(*) as count, SUM(total) as revenue
        FROM orders {date_condition}
        GROUP BY source
    ''').fetchall()
    source_data = [dict(row) for row in source_data_raw]
    
    # Recent orders - only filter by source, ignore date filter
    if source_filter:
        recent_orders = conn.execute(f'''
            SELECT id, number, status, total, date_created, source, line_items
            FROM orders
            WHERE source = '{source_filter}'
            ORDER BY date_created DESC
            LIMIT 10
        ''').fetchall()
    else:
        recent_orders = conn.execute('''
            SELECT id, number, status, total, date_created, source, line_items
            FROM orders
            ORDER BY date_created DESC
            LIMIT 10
        ''').fetchall()
    
    # Process recent orders to get product count
    processed_orders = []
    for order in recent_orders:
        order_dict = dict(order)
        items = parse_json_field(order['line_items'])
        order_dict['product_count'] = sum(item.get('quantity', 0) for item in items) if isinstance(items, list) else 0
        processed_orders.append(order_dict)
    
    # Trend data based on filter - daily for single month, monthly for longer periods
    trend_type = 'daily' if quick_date in ['this_month', 'last_month'] else 'monthly'
    
    if trend_type == 'daily' and date_condition:
        # Daily trend for this_month or last_month
        trend_data_raw = conn.execute(f'''
            SELECT strftime('%Y-%m-%d', date_created) as period, 
                   COUNT(*) as orders,
                   SUM(total) as revenue
            FROM orders {date_condition}
            GROUP BY period
            ORDER BY period
        ''').fetchall()
    elif date_condition:
        # Monthly trend for other filters
        trend_data_raw = conn.execute(f'''
            SELECT strftime('%Y-%m', date_created) as period, 
                   COUNT(*) as orders,
                   SUM(total) as revenue
            FROM orders {date_condition}
            GROUP BY period
            ORDER BY period
        ''').fetchall()
    else:
        # Default to last 6 months if no filter
        trend_data_raw = conn.execute('''
            SELECT strftime('%Y-%m', date_created) as period, 
                   COUNT(*) as orders,
                   SUM(total) as revenue
            FROM orders
            WHERE date_created >= date('now', '-6 months')
            GROUP BY period
            ORDER BY period
        ''').fetchall()
    trend_data = [dict(row) for row in trend_data_raw]
    
    conn.close()
    
    return render_template('dashboard.html',
                         stats=stats,
                         status_data=status_data,
                         source_data=source_data,
                         recent_orders=processed_orders,
                         trend_data=trend_data,
                         trend_type=trend_type,
                         quick_date=quick_date,
                         sources=sources,
                         source_filter=source_filter)


@app.route('/orders')
@login_required
def orders():
    """Order list with filtering and summary statistics"""
    from datetime import date, timedelta
    conn = get_db_connection()
    
    # Get filter parameters
    source_filter = request.args.get('source', '')
    status_filter = request.args.get('status', '')
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')
    search = request.args.get('search', '')
    quick_date = request.args.get('quick_date', '')
    page = int(request.args.get('page', 1))
    per_page = 20
    
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
    
    # Build base conditions
    conditions = []
    params = []
    
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
    
    # Summary statistics by source
    stats_query = f'''
        SELECT source,
            COUNT(*) as total_orders,
            SUM(total) as total_amount,
            SUM(shipping_total) as total_shipping,
            SUM(CASE WHEN status NOT IN ('failed','cancelled') THEN 1 ELSE 0 END) as success_orders,
            SUM(CASE WHEN status NOT IN ('failed','cancelled') THEN total ELSE 0 END) as success_amount,
            SUM(CASE WHEN status NOT IN ('failed','cancelled') THEN shipping_total ELSE 0 END) as success_shipping,
            SUM(CASE WHEN status='failed' THEN 1 ELSE 0 END) as failed_orders,
            SUM(CASE WHEN status='cancelled' THEN 1 ELSE 0 END) as cancelled_orders
        FROM orders {where_clause} GROUP BY source
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
    
    sources = conn.execute('SELECT DISTINCT source FROM orders').fetchall()
    statuses = conn.execute('SELECT DISTINCT status FROM orders').fetchall()
    conn.close()
    
    return render_template('orders.html',
                         orders=processed_orders,
                         sources=sources,
                         statuses=statuses,
                         summary_stats=summary_stats,
                         totals=totals,
                         current_filters={
                             'source': source_filter,
                             'status': status_filter,
                             'date_from': date_from,
                             'date_to': date_to,
                             'search': search,
                             'quick_date': quick_date
                         },
                         available_months=available_months,
                         current_month=current_month,
                         total=total)



@app.route('/monthly')
@login_required
def monthly():
    """Monthly statistics page"""
    conn = get_db_connection()
    
    # Get source filter
    source_filter = request.args.get('source', '')
    
    # Get all sources for dropdown
    all_sources = conn.execute('SELECT DISTINCT source FROM orders ORDER BY source').fetchall()
    
    # Build query with optional source filter
    if source_filter:
        query = f'''
            SELECT id, status, date_created, total, shipping_total, source, line_items
            FROM orders
            WHERE source = '{source_filter}'
            ORDER BY date_created DESC
        '''
    else:
        query = '''
            SELECT id, status, date_created, total, shipping_total, source, line_items
            FROM orders
            ORDER BY date_created DESC
        '''
    
    df = pd.read_sql_query(query, conn)
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
    
    # Group by month and source
    rows = []
    for (month, source), gdf in df.groupby(['month', 'source']):
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
    
    return render_template('monthly.html', monthly_stats=rows, sources=all_sources, source_filter=source_filter, sort_by=sort_by)


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


# Status translation
STATUS_LABELS = {
    'pending': '待处理',
    'processing': '处理中',
    'on-hold': '暂停',
    'completed': '已完成',
    'cancelled': '已取消',
    'refunded': '已退款',
    'failed': '失败'
}


@app.template_filter('status_label')
def status_label_filter(status):
    return STATUS_LABELS.get(status, status)


@app.template_filter('format_currency')
def format_currency_filter(value):
    try:
        return f"{float(value):,.2f}"
    except:
        return value


@app.template_filter('format_date')
def format_date_filter(value):
    try:
        dt = datetime.fromisoformat(value.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d %H:%M')
    except:
        return value


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
