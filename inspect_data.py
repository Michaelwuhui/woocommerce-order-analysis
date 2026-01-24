import sqlite3

def normalize_domain(url):
    if not url: return ""
    return url.lower().replace('https://', '').replace('http://', '').replace('www.', '').split('/')[0]

conn = sqlite3.connect('woocommerce_orders.db')
conn.row_factory = sqlite3.Row

print("--- Sites ---")
sites = conn.execute('SELECT url, mask_id FROM sites').fetchall()
site_mask_map = {}
for site in sites:
    url = site['url']
    normalized = normalize_domain(url)
    print(f"Original: {repr(url)}, Normalized: {repr(normalized)}")
    if site['mask_id']:
         site_mask_map[normalized] = site['mask_id']

print("\n--- Orders Sources ---")
orders = conn.execute('SELECT DISTINCT source FROM orders').fetchall()
for order in orders:
    source = order['source']
    norm_source = normalize_domain(source)
    site_name = None
    
    # Simulate app logic
    if norm_source in site_mask_map:
        site_name = norm_source
        match_type = "Direct Match"
    else:
        for name in site_mask_map.keys():
            if name in norm_source:
                site_name = name
                match_type = "Partial Match"
                break
    
    if not site_name:
        site_name = norm_source
        match_type = "Fallback"
        
    print(f"Source: {repr(source)} -> Norm: {repr(norm_source)} -> SiteName: {repr(site_name)} ({match_type})")

conn.close()
