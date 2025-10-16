import sqlite3

# 连接到SQLite数据库
conn = sqlite3.connect('woocommerce_orders.db')
cursor = conn.cursor()

# 检查总订单数
cursor.execute('SELECT COUNT(*) FROM orders')
total_orders = cursor.fetchone()[0]
print(f'数据库中总订单数: {total_orders}')

# 按站点统计订单数（使用source字段）
cursor.execute('SELECT source, COUNT(*) FROM orders GROUP BY source')
print('\n按站点统计:')
for row in cursor.fetchall():
    print(f'{row[0]}: {row[1]} 个订单')

# 检查最新的几个订单
cursor.execute('SELECT id, source, number, status, total, date_created FROM orders ORDER BY updated_at DESC LIMIT 5')
print('\n最新的5个订单:')
for row in cursor.fetchall():
    print(f'订单号: {row[2]}, 站点: {row[1]}, 状态: {row[3]}, 总额: {row[4]}, 创建时间: {row[5]}')

conn.close()