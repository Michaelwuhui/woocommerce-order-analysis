#!/usr/bin/env python3
"""One-off repair: re-host product description/short_description images that
leaked from another of our own sites (e.g. strefajednorazowek.pl) into the
target master's media library, then rewrite the copy.

Same root cause as the clone bug fixed in app.py (_migrate_inline_images_after_clone):
WooCommerce only sideloads the gallery (`images` array), never inline content
images, and WC consumer keys can't use /wp/v2/media. So we reuse WC's own
sideloading via the product's images array (PUT #1), read back the new URLs,
then save the rewritten HTML + restore the original gallery (PUT #2).

Usage:
  venv/bin/python repair_inline_images.py            # dry-run, master id 2
  venv/bin/python repair_inline_images.py --apply     # actually fix
  venv/bin/python repair_inline_images.py --master-id 3 --apply
  venv/bin/python repair_inline_images.py --apply --only 11345,9946
"""
import argparse
import re
import sqlite3
import sys
from urllib.parse import urlparse

import requests as req

DB = '/www/wwwroot/woo-analysis/woocommerce_orders.db'
WC_HEADERS = {
    'User-Agent': 'WooCommerce API Client-Python/3.0.0',
    'Content-Type': 'application/json',
    'Accept': 'application/json',
}
ASSET_RE = re.compile(
    r"""(?:https?:)?//[^/"'\s)]+/[^"'\s)]+?\.(jpe?g|png|gif|webp|avif|svg|bmp|mp4|webm|mov|m4v|ogg|ogv|mp3|wav|pdf)(?:\?[^"'\s)]*)?""",
    re.IGNORECASE,
)
IMAGE_EXT = {'jpg', 'jpeg', 'png', 'gif', 'webp', 'avif', 'svg', 'bmp'}


def family_domain(host):
    """Last two labels, e.g. master.vapego.pl -> vapego.pl (good enough for
    our .pl/.com domains)."""
    parts = (host or '').lower().split('.')
    return '.'.join(parts[-2:]) if len(parts) >= 2 else host


def extract_leaked(html, fam):
    """Find leaked media (hosted off our target network, under wp-content/uploads).

    Returns (img_order, img_groups, other_order, other_groups). Images can be
    re-hosted via WC's image sideloading; 'other' (mp4/webm/audio/pdf) cannot
    (WC rejects non-image uploads) — those need wp/v2/media + an app password,
    so we only report them."""
    img_order, img_groups, oth_order, oth_groups = [], {}, [], {}
    if not html:
        return img_order, img_groups, oth_order, oth_groups
    for m in ASSET_RE.finditer(html):
        tok = m.group(0)
        ext = m.group(1).lower()
        absu = ('https:' + tok) if tok.startswith('//') else tok
        host = urlparse(absu).netloc.lower()
        if host.endswith(fam):
            continue  # already on our target network — fine
        if '/wp-content/uploads/' not in urlparse(absu).path.lower():
            continue  # not a WP media leak — leave intentional embeds alone
        key = re.sub(r'^https?:', '', absu)
        order, groups = ((img_order, img_groups) if ext in IMAGE_EXT
                         else (oth_order, oth_groups))
        if key not in groups:
            groups[key] = {'submit': absu.replace('&amp;', '&'), 'tokens': set()}
            order.append(key)
        groups[key]['tokens'].add(tok)
    return img_order, img_groups, oth_order, oth_groups


def wc_get(url, ck, cs, **kw):
    return req.get(url, auth=(ck, cs), timeout=60, headers=WC_HEADERS, **kw)


def parse(resp):
    if resp.status_code not in (200, 201):
        try:
            return None, f'HTTP {resp.status_code}: {resp.json().get("message")}'
        except Exception:
            return None, f'HTTP {resp.status_code}: {resp.text[:200]}'
    try:
        return resp.json(), None
    except Exception:
        return None, f'非JSON响应: {resp.text[:200]}'


def url_alive(u):
    try:
        r = req.get(u, timeout=20, stream=True,
                    headers={'User-Agent': WC_HEADERS['User-Agent']})
        r.close()
        return r.status_code == 200
    except Exception:
        return False


def repair_product(base, ck, cs, p, fam, apply):
    pid = p['id']
    desc = p.get('description', '') or ''
    short = p.get('short_description', '') or ''
    order, groups, oth_order, _oth_groups = extract_leaked(desc + '\n' + short, fam)
    if not order and not oth_order:
        return None  # not affected

    info = {'id': pid, 'name': p.get('name', '')[:45], 'status': p.get('status'),
            'leaked': len(order), 'video': len(oth_order), 'fixed': 0, 'dead': 0,
            'note': ''}
    if not apply:
        info['note'] = 'dry-run'
        return info
    if not order:
        info['note'] = '仅有视频/媒体泄漏，WC无法迁移(需应用密码)'
        return info

    # Drop dead source links (can't re-host a deleted image).
    live_keys = []
    for k in order:
        if url_alive(groups[k]['submit']):
            live_keys.append(k)
        else:
            info['dead'] += 1
    if not live_keys:
        info['note'] = '所有泄漏图片在源站已失效，无法重新托管'
        return info

    content_urls = [groups[k]['submit'] for k in live_keys]
    gallery_ids = [im.get('id') for im in (p.get('images') or []) if im.get('id')]
    gcount = len(gallery_ids)
    # Safety: on a PUBLISHED product with no gallery, appending content images
    # would make the first one the featured image for the brief PUT#1→PUT#2
    # window. Skip rather than risk a wrong featured image on a live product.
    if gcount == 0 and p.get('status') == 'publish':
        info['note'] = '已发布且无相册，跳过(避免误设主图)；建议手动或转草稿后再修'
        return info
    put_url = f'{base}/wp-json/wc/v3/products/{pid}'

    # PUT #1 — append leaked images so WC sideloads them.
    put1 = {'images': [{'id': g} for g in gallery_ids] + [{'src': u} for u in content_urls]}
    data1, err = parse(req.put(put_url, auth=(ck, cs), json=put1, timeout=180, headers=WC_HEADERS))
    if err:
        info['note'] = f'PUT#1 失败: {err}'
        return info
    resp_imgs = data1.get('images') or []
    new_for_content = resp_imgs[gcount:]

    url_map = {}
    if len(new_for_content) == len(content_urls):
        for k, im in zip(live_keys, new_for_content):
            if im.get('src'):
                url_map[k] = im['src']
    else:
        def basekey(u):
            n = urlparse(u).path.rsplit('/', 1)[-1]
            n = re.sub(r'\.(?:jpe?g|png|gif|webp|avif|svg|bmp)$', '', n, flags=re.I)
            return re.sub(r'-\d+x\d+$', '', n).lower()
        by_base = {}
        for im in resp_imgs:
            if im.get('src'):
                by_base.setdefault(basekey(im['src']), im['src'])
        for k in live_keys:
            ns = by_base.get(basekey(groups[k]['submit']))
            if ns:
                url_map[k] = ns

    if not url_map:
        info['note'] = 'PUT#1 后未匹配到新地址'
        return info

    def rewrite(h):
        for k, ns in url_map.items():
            for tok in groups[k]['tokens']:
                h = h.replace(tok, ns)
        return h

    put2 = {'description': rewrite(desc), 'short_description': rewrite(short),
            'images': [{'id': g} for g in gallery_ids]}
    _, err = parse(req.put(put_url, auth=(ck, cs), json=put2, timeout=180, headers=WC_HEADERS))
    if err:
        info['note'] = f'图片已上传但回写失败: {err}'
        return info

    info['fixed'] = len(url_map)
    info['note'] = 'OK'
    return info


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--master-id', type=int, default=2)
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--only', default='', help='逗号分隔的产品ID，仅处理这些')
    args = ap.parse_args()

    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    m = conn.execute('SELECT url, consumer_key, consumer_secret FROM product_masters WHERE id=?',
                     (args.master_id,)).fetchone()
    conn.close()
    if not m:
        print(f'master id {args.master_id} 不存在'); sys.exit(1)
    base = m['url'].rstrip('/')
    ck, cs = m['consumer_key'], m['consumer_secret']
    fam = family_domain(urlparse(base).netloc)
    only = {int(x) for x in args.only.split(',') if x.strip().isdigit()}

    print(f"目标 master: {base}  (家族域: {fam})  模式: {'APPLY 写入' if args.apply else 'DRY-RUN 只读'}")
    print('=' * 78)

    affected, page = [], 1
    while True:
        data, err = parse(wc_get(f'{base}/wp-json/wc/v3/products',
                                 ck, cs, params={'per_page': 100, 'page': page,
                                                 'status': 'any'}))
        if err:
            print('拉取产品失败:', err); break
        if not data:
            break
        for p in data:
            if only and p['id'] not in only:
                continue
            r = repair_product(base, ck, cs, p, fam, args.apply)
            if r:
                affected.append(r)
                print(f"  [{r['status']:>7}] ID {r['id']:<6} 图片{r['leaked']:>2} "
                      f"视频{r['video']:>2} 修复{r['fixed']:>2} 失效{r['dead']:>2}  "
                      f"{r['note']:<14} {r['name']}")
        if len(data) < 100:
            break
        page += 1

    print('=' * 78)
    tot_leaked = sum(a['leaked'] for a in affected)
    tot_video = sum(a['video'] for a in affected)
    tot_fixed = sum(a['fixed'] for a in affected)
    tot_dead = sum(a['dead'] for a in affected)
    print(f"受影响产品: {len(affected)} 个 | 泄漏图片: {tot_leaked} | 已修复: {tot_fixed} | "
          f"源站失效: {tot_dead} | 泄漏视频/媒体(无法自动迁移): {tot_video}")
    if not args.apply and affected:
        print("\n这是 DRY-RUN。确认无误后加 --apply 实际修复。")


if __name__ == '__main__':
    main()
