
import os
import re

file_path = r'c:\Users\Administrator\Documents\GitHub\woocommerce-order-analysis\app.py'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add Global Variables
# We'll add them after imports or near SYNC_STATUS
if "SYNC_STATUS = {}" in content:
    content = content.replace("SYNC_STATUS = {}", "SYNC_STATUS = {}\nAUTO_SYNC_ENABLED = False\nAUTO_SYNC_INTERVAL = 900  # 15 minutes")

# 2. Add API Endpoints
api_code = """
@app.route('/api/settings/autosync', methods=['GET'])
@login_required
def get_autosync_status():
    return jsonify({
        'enabled': AUTO_SYNC_ENABLED,
        'interval': AUTO_SYNC_INTERVAL
    })

@app.route('/api/settings/autosync', methods=['POST'])
@login_required
def set_autosync_status():
    global AUTO_SYNC_ENABLED
    data = request.json
    if 'enabled' in data:
        AUTO_SYNC_ENABLED = bool(data['enabled'])
    return jsonify({'success': True, 'enabled': AUTO_SYNC_ENABLED})
"""

# Insert before if __name__ == '__main__':
if "if __name__ == '__main__':" in content:
    content = content.replace("if __name__ == '__main__':", api_code + "\nif __name__ == '__main__':")

# 3. Add Background Thread Loop
thread_code = """
def auto_sync_loop(app_context):
    import time
    import sync_utils
    
    print("Auto-sync thread started")
    while True:
        try:
            if AUTO_SYNC_ENABLED:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Auto-sync triggering...")
                # We reuse the logic from sync_all_data but we need to call it directly or extract it.
                # Since sync_all_data is a route, we can't call it easily.
                # Let's extract the inner logic of sync_all_data or just hit the endpoint internally?
                # Hitting endpoint requires a request context or running server.
                # Better to duplicate the logic or extract it.
                
                # To avoid code duplication, let's just do a simplified version here.
                # Or better, we can invoke the same function if we refactor.
                # For now, I will copy the core logic since I can't easily refactor the route in this script.
                
                with app_context:
                    # Check if a sync is already running?
                    # For simplicity, we just run it.
                    
                    # We need to use the ALL_SITES_ID = 999999
                    ALL_SITES_ID = 999999
                    
                    # Only run if not already running
                    current_status = SYNC_STATUS.get(ALL_SITES_ID, {}).get('status')
                    if current_status == 'running':
                        print("Auto-sync skipped: Sync already in progress")
                    else:
                        # Initialize status
                        SYNC_STATUS[ALL_SITES_ID] = {
                            'status': 'running',
                            'message': 'Auto-sync started...',
                            'logs': [f"[{datetime.now().strftime('%H:%M:%S')}] Auto-sync job started"]
                        }
                        
                        try:
                            conn = get_db_connection()
                            sites = conn.execute('SELECT * FROM sites').fetchall()
                            conn.close()
                            
                            if sites:
                                total_sites = len(sites)
                                for index, site in enumerate(sites):
                                    site_url = site['url']
                                    current_step = index + 1
                                    
                                    SYNC_STATUS[ALL_SITES_ID]['message'] = f"Auto-syncing {current_step}/{total_sites}: {site_url}"
                                    
                                    def progress_callback(msg):
                                        pass # Mute detailed logs for auto-sync to save memory? Or keep them?
                                        # Let's keep them but minimal
                                    
                                    result = sync_utils.sync_site(
                                        site['url'], 
                                        site['consumer_key'], 
                                        site['consumer_secret'],
                                        progress_callback
                                    )
                                    
                                    if result['status'] == 'success':
                                        conn = get_db_connection()
                                        conn.execute('UPDATE sites SET last_sync = ? WHERE id = ?', 
                                                     (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), site['id']))
                                        conn.commit()
                                        conn.close()
                                
                                SYNC_STATUS[ALL_SITES_ID]['status'] = 'success'
                                SYNC_STATUS[ALL_SITES_ID]['message'] = 'Auto-sync completed'
                                SYNC_STATUS[ALL_SITES_ID]['logs'].append(f"[{datetime.now().strftime('%H:%M:%S')}] Auto-sync finished")
                                
                        except Exception as e:
                            print(f"Auto-sync error: {e}")
                            SYNC_STATUS[ALL_SITES_ID]['status'] = 'error'
            
            # Sleep for interval
            time.sleep(AUTO_SYNC_INTERVAL)
            
        except Exception as e:
            print(f"Auto-sync loop error: {e}")
            time.sleep(60) # Sleep on error
"""

# Insert thread loop definition before main
content = content.replace("if __name__ == '__main__':", thread_code + "\nif __name__ == '__main__':")

# 4. Start Thread in Main
# We need to start it only once.
main_replacement = """if __name__ == '__main__':
    # Start auto-sync thread
    import threading
    # Daemon thread so it dies when main thread dies
    sync_thread = threading.Thread(target=auto_sync_loop, args=(app.app_context(),), daemon=True)
    sync_thread.start()
    
    app.run(debug=True, host='0.0.0.0', port=5000)"""

content = content.replace("if __name__ == '__main__':\n    app.run(debug=True, host='0.0.0.0', port=5000)", main_replacement)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Successfully added auto-sync backend logic")
