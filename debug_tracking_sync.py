#!/usr/bin/env python3
"""
调试脚本：测试订单tracking同步
"""

import sys
import json

# 添加Flask应用路径
sys.path.insert(0, '/www/wwwroot/woo-analysis')

from app import app, get_db_connection

def debug_order_tracking_sync(order_number):
    """调试订单tracking同步"""
    
    with app.app_context():
        conn = get_db_connection()
        
        # 查找订单
        order = conn.execute('''
            SELECT o.*, sl.tracking_number, sl.carrier_slug
            FROM orders o
            LEFT JOIN shipping_logs sl ON o.id = sl.order_id
            WHERE o.number = ?
        ''', (order_number,)).fetchone()
        
        if not order:
            print(f"❌ 订单 #{order_number} 不存在")
            return
        
        print(f"\n{'='*60}")
        print(f"订单信息")
        print(f"{'='*60}")
        print(f"订单ID: {order['id']}")
        print(f"订单号: {order['number']}")
        print(f"来源: {order['source']}")
        print(f"运单号: {order['tracking_number'] or '未设置'}")
        print(f"物流商: {order['carrier_slug'] or '未设置'}")
        
        if not order['tracking_number']:
            print("\n⚠️  该订单没有运单号，无法同步")
            return
        
        # 获取站点信息
        site = conn.execute('SELECT * FROM sites WHERE url = ?', (order['source'],)).fetchone()
        conn.close()
        
        if not site:
            print(f"\n❌ 找不到站点配置: {order['source']}")
            return
        
        print(f"\n{'='*60}")
        print(f"站点配置")
        print(f"{'='*60}")
        print(f"站点URL: {site['url']}")
        print(f"Consumer Key: {site['consumer_key'][:20]}...")
        print(f"Consumer Secret: {'*' * 20}")
        
        # 构建API请求
        tracking_url_api = f"{site['url']}/wp-json/woo-orders-tracking/v1/tracking/set"
        tracking_payload = {
            'order_id': order['number'],
            'tracking_data': [{
                'tracking_number': order['tracking_number'],
                'carrier_slug': order['carrier_slug']
            }],
            'send_email': False
        }
        
        print(f"\n{'='*60}")
        print(f"API请求")
        print(f"{'='*60}")
        print(f"端点: {tracking_url_api}")
        print(f"方法: POST")
        print(f"Payload:")
        print(json.dumps(tracking_payload, indent=2, ensure_ascii=False))
        
        # 发送请求
        print(f"\n{'='*60}")
        print(f"发送API请求...")
        print(f"{'='*60}")
        
        import requests
        try:
            response = requests.post(
                tracking_url_api,
                json=tracking_payload,
                auth=(site['consumer_key'], site['consumer_secret']),
                timeout=30
            )
            
            print(f"\n✅ 请求已发送")
            print(f"\n{'='*60}")
            print(f"响应信息")
            print(f"{'='*60}")
            print(f"状态码: {response.status_code}")
            print(f"响应头:")
            for key, value in response.headers.items():
                if key.lower() in ['content-type', 'x-wp-nonce', 'link']:
                    print(f"  {key}: {value}")
            
            print(f"\n响应体:")
            try:
                response_json = response.json()
                print(json.dumps(response_json, indent=2, ensure_ascii=False))
            except:
                print(response.text[:500])
            
            if response.status_code in [200, 201]:
                print(f"\n✅ 同步成功！")
            else:
                print(f"\n❌ 同步失败")
                
        except Exception as e:
            print(f"\n❌ 请求异常: {str(e)}")
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    order_number = '14820'
    if len(sys.argv) > 1:
        order_number = sys.argv[1]
    
    debug_order_tracking_sync(order_number)
