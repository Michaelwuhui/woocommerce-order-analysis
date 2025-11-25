import sqlite3

# 连接到数据库
conn = sqlite3.connect('woocommerce_orders.db')
cursor = conn.cursor()

# 查询总订单数
cursor.execute('SELECT COUNT(*) FROM orders')
total_orders = cursor.fetchone()[0]
print(f'总订单数: {total_orders}')

# 查询各站点订单数
cursor.execute('SELECT source, COUNT(*) FROM orders GROUP BY source')
print('各站点订单数:')
for row in cursor.fetchall():
    print(f'  {row[0]}: {row[1]}')

# 查询最新订单日期
cursor.execute('SELECT MAX(date_created) FROM orders')
latest_date = cursor.fetchone()[0]
print(f'最新订单日期: {latest_date}')

# 查询最近5个订单
cursor.execute('SELECT id, date_created, status, total, source FROM orders ORDER BY date_created DESC LIMIT 5')
print('\n最近5个订单:')
for row in cursor.fetchall():
    print(f'  订单ID: {row[0]}, 日期: {row[1]}, 状态: {row[2]}, 总额: {row[3]}, 来源: {row[4]}')

conn.close()