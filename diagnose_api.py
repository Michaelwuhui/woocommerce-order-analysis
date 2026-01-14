import sys
import time
sys.path.insert(0, '/www/wwwroot/woo-analysis')
from app import app, get_db_connection
from woocommerce import API
import requests as req

def run_diagnosis():
    with app.app_context():
        conn = get_db_connection()
        # 获取站点信息 (buchmistrz.pl)
        site = conn.execute("SELECT * FROM sites WHERE url LIKE '%buchmistrz.pl%'").fetchone()
        if not site:
            print("错误: 找不到 buchmistrz.pl 站点")
            return

        order_id = 44951 # 目标订单
        print(f"=== 开始诊断订单 #{order_id} (站点: {site['url']}) ===")

        wcapi = API(
            url=site['url'],
            consumer_key=site['consumer_key'],
            consumer_secret=site['consumer_secret'],
            version="wc/v3",
            timeout=30
        )

        # 1. 测试读取
        print("\n[1/3] 测试读取订单...")
        try:
            resp = wcapi.get(f"orders/{order_id}")
            if resp.status_code == 200:
                print(f"✅ 读取成功! 状态: {resp.json().get('status')}")
            else:
                print(f"❌ 读取失败: {resp.status_code} - {resp.text}")
                return
        except Exception as e:
            print(f"❌ 读取连接错误: {e}")
            return

        # 2. 测试写入纯文本 (无特殊字符)
        print("\n[2/3] 测试写入纯文本备注 (Hello World)...")
        try:
            data = {
                "note": "诊断测试: 纯文本消息 (此消息可删除)",
                "customer_note": False
            }
            start_time = time.time()
            resp = wcapi.post(f"orders/{order_id}/notes", data=data)
            duration = time.time() - start_time
            
            if resp.status_code in [200, 201]:
                print(f"✅ 纯文本写入成功! (耗时: {duration:.2f}s)")
                # 清理
                note_id = resp.json().get('id')
                if note_id:
                     wcapi.delete(f"orders/{order_id}/notes/{note_id}", params={"force": True})
            else:
                print(f"❌ 纯文本写入失败: {resp.status_code}")
                print(f"响应: {resp.text}")
        except Exception as e:
            print(f"❌ 纯文本写入连接错误: {e}")

        # 3. 测试写入带 HTML 链接的内容 (模拟发货)
        print("\n[3/3] 测试写入带 HTML 链接的内容 (模拟发货)...")
        try:
            html_content = 'Order shipped. Track here: <a href="https://inpost.pl/track?number=123">123</a>'
            data = {
                "note": html_content,
                "customer_note": False
            }
            start_time = time.time()
            resp = wcapi.post(f"orders/{order_id}/notes", data=data)
            duration = time.time() - start_time
            
            if resp.status_code in [200, 201]:
                print(f"✅ HTML 内容写入成功! (耗时: {duration:.2f}s)")
                # 清理
                note_id = resp.json().get('id')
                if note_id:
                     wcapi.delete(f"orders/{order_id}/notes/{note_id}", params={"force": True})
            else:
                print(f"❌ HTML 内容写入失败: {resp.status_code}")
                # 尝试解析是否被防火墙拦截
                if "403" in str(resp.status_code) or "Security" in resp.text:
                    print("⚠️  疑似被防火墙拦截！")
                print(f"响应片段: {resp.text[:200]}...")
        except Exception as e:
            print(f"❌ HTML 内容写入连接错误: {e}")
            print("⚠️  典型的防火墙拦截特征：连接直接被重置或超时")

if __name__ == "__main__":
    run_diagnosis()
