#!/usr/bin/env python3
"""
测试打印API返回的数据
"""

import sys
sys.path.insert(0, '/www/wwwroot/woo-analysis')

from app import app
import json

with app.test_client() as client:
    # 需要先登录
    with client.session_transaction() as sess:
        sess['user_id'] = 1  # 假设admin用户ID是1
    
    # 测试pending list API
    response = client.get('/api/shipping/print/pending')
    
    if response.status_code == 200:
        data = response.get_json()
        print("=" * 60)
        print(f"API返回订单数量: {data['count']}")
        print("=" * 60)
        
        # 显示前3个订单的详细信息
        for i, order in enumerate(data['orders'][:3]):
            print(f"\n订单 #{i+1}: {order['order_number']}")
            print(f"  shipping_method: {order.get('shipping_method', 'N/A')}")
            print(f"  customer_inpost_id: {order.get('customer_inpost_id', 'N/A')}")
            print(f"  customer_address: {order.get('customer_address', 'N/A')[:50]}...")
            print(f"  customer_address_2: {order.get('customer_address_2', 'N/A')}")
            print(f"  customer_email: {order.get('customer_email', 'N/A')}")
    else:
        print(f"API请求失败: {response.status_code}")
        print(f"响应: {response.data}")
