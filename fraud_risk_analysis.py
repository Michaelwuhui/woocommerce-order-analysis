#!/usr/bin/env python3
"""
退货/拒收风险 — 第一阶段离线验证脚本 (完全只读，不修改 app.py，不写数据库)

目的:验证"跨站点身份历史风险分"这个信号到底准不准、值不值得做进系统。
做三件事:
  1. 电话归一化覆盖率报告 (身份能不能可靠地跨 18 站点合并)
  2. 高风险身份排行榜 (今天就能据此行动的名单)
  3. 时点回测 (point-in-time): 对每一单只用它"之前"的历史算分，
     再看真正变坏的订单里有多少能被提前标出来 —— 这是信号有没有用的硬指标。

用法:  venv/bin/python fraud_risk_analysis.py
输出:  控制台摘要 + exports/fraud_risk_identities_<时间>.csv
"""
import sqlite3
import json
import re
import csv
import os
from collections import defaultdict
from datetime import datetime

DB_FILE = 'woocommerce_orders.db'

# ---- 与 app.py 完全一致的状态口径 ----
EXCLUDED = {'checkout-draft', 'trash'}
GOOD = {'completed', 'shipped', 'delivered', 'partial-shipped'}
SOFT_BAD = {'cancelled', 'failed'}          # 软坏单(没发出去/没成交)
HARD_BAD_STATUS = {'cheat'}                 # 人工标记的作弊
# is_undelivered / is_problem_return 来自专门字段，单独处理

# ---- 评分权重 (放在顶部方便调) ----
W_UNDELIVERED   = 35   # 每次历史拒收(真金白银的运费损失)，最强信号
W_PROBLEM_RET   = 45   # 每次历史问题退货(确认欺诈)
W_CHEAT         = 50   # 历史被标记 cheat
W_BADRATE       = 25   # 历史坏单率高(需≥MIN_ORDERS_FOR_RATE 单)
W_CROSSSITE     = 15   # 跨多站点下单
W_NEW_COD_BIG   = 15   # 新客 + COD + 大额
W_VELOCITY      = 10   # 短时间内多单
CAP_REPEAT      = 75   # 重复类信号(拒收/退货)累计上限
MIN_ORDERS_FOR_RATE = 3
HIGH_BADRATE    = 0.5
CROSSSITE_MIN   = 3
VELOCITY_HOURS  = 48
VELOCITY_MIN    = 3
HIGH_THRESHOLD  = 60
MED_THRESHOLD   = 30


def normalize_phone(raw):
    """归一化电话为身份键。波兰号(绝大多数)统一成 9 位裸号；国际号保留全数字。
    返回 None 表示无法作为可靠身份键。"""
    if not raw:
        return None
    d = re.sub(r'\D', '', str(raw))
    if d.startswith('0048'):
        d = d[4:]
    if len(d) == 11 and d.startswith('48'):  # +48 / 48 前缀的波兰号
        d = d[2:]
    if len(d) < 7:        # 太短 = 垃圾数据(如把邮箱填进电话)
        return None
    return d


def identity_key(billing_json):
    """优先用归一化电话，回退到邮箱。返回 (key, kind) 或 (None, None)。"""
    try:
        b = json.loads(billing_json or '{}')
    except Exception:
        return None, None
    ph = normalize_phone(b.get('phone'))
    if ph:
        return f"tel:{ph}", 'phone'
    em = (b.get('email') or '').strip().lower()
    if em and '@' in em:
        return f"mail:{em}", 'email'
    return None, None


def parse_dt(s):
    if not s:
        return None
    s = s.replace('T', ' ').strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try:
            return datetime.strptime(s[:19], fmt)
        except Exception:
            continue
    return None


def classify(o):
    """把一单归类为 good / soft_bad / undelivered / problem_return / cheat / neutral。"""
    if o['is_problem_return'] == 1:
        return 'problem_return'
    if o['is_undelivered'] == 1:
        return 'undelivered'
    st = (o['status'] or '').lower()
    if st in HARD_BAD_STATUS:
        return 'cheat'
    if st in SOFT_BAD:
        return 'soft_bad'
    if st in GOOD:
        return 'good'
    return 'neutral'


def main():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    rows = conn.execute("""
        SELECT id, number, status, total, currency, source,
               billing, customer_ip_address, date_created, date_created_gmt,
               is_undelivered, is_problem_return, product_loss_amount, shipping_loss_amount,
               payment_method
        FROM orders
        WHERE status NOT IN ('checkout-draft', 'trash')
    """).fetchall()
    conn.close()

    # ---- 构建身份 -> 订单列表(按时间排序) ----
    identities = defaultdict(list)
    no_key = 0
    high_total_samples = []
    for o in rows:
        key, kind = identity_key(o['billing'])
        if key is None:
            no_key += 1
            continue
        dt = parse_dt(o['date_created_gmt'] or o['date_created'])
        identities[key].append({
            'row': o, 'dt': dt, 'kind': kind,
            'cls': classify(o),
            'cod': (o['payment_method'] or '') == 'cod',
            'total': float(o['total'] or 0),
            'site': o['source'],
        })
        if (o['status'] or '') in GOOD:
            high_total_samples.append(float(o['total'] or 0))

    for k in identities:
        identities[k].sort(key=lambda x: (x['dt'] or datetime.min))

    high_total_samples.sort()
    p75 = high_total_samples[int(len(high_total_samples) * 0.75)] if high_total_samples else 200.0

    # ============ 1. 归一化覆盖率 ============
    total_linkable = sum(len(v) for v in identities.values())
    phone_ids = sum(1 for k in identities if k.startswith('tel:'))
    email_ids = sum(1 for k in identities if k.startswith('mail:'))
    repeat_ids = {k: v for k, v in identities.items() if len(v) >= 2}
    print("=" * 64)
    print("1. 身份归一化覆盖率")
    print("=" * 64)
    print(f"  非草稿订单总数        : {len(rows)}")
    print(f"  无法关联身份(无电话/邮箱): {no_key}  ({100*no_key//max(len(rows),1)}%)")
    print(f"  可关联订单            : {total_linkable}")
    print(f"  独立身份数            : {len(identities)}  (电话 {phone_ids} / 邮箱 {email_ids})")
    print(f"  回头客身份(>=2单)     : {len(repeat_ids)}  覆盖 {sum(len(v) for v in repeat_ids.values())} 单")
    print(f"  '大额'阈值(好单75分位): {p75:.0f}")

    # ============ 2 & 3. 时点评分 + 回测 ============
    def score_order(prior, this):
        """只用 prior(该身份在本单之前的订单)给 this 算分。返回 (score, reasons)。"""
        s = 0
        reasons = []
        und = sum(1 for p in prior if p['cls'] == 'undelivered')
        prob = sum(1 for p in prior if p['cls'] == 'problem_return')
        cheat = sum(1 for p in prior if p['cls'] == 'cheat')
        bad = sum(1 for p in prior if p['cls'] in ('undelivered', 'problem_return', 'cheat', 'soft_bad'))
        n = len(prior)
        sites = set(p['site'] for p in prior) | {this['site']}

        repeat_pts = 0
        if und:
            repeat_pts += W_UNDELIVERED * und
            reasons.append(f"历史拒收{und}次")
        if prob:
            repeat_pts += W_PROBLEM_RET * prob
            reasons.append(f"历史问题退货{prob}次")
        if cheat:
            repeat_pts += W_CHEAT * cheat
            reasons.append(f"曾被标记cheat{cheat}次")
        s += min(repeat_pts, CAP_REPEAT)

        if n >= MIN_ORDERS_FOR_RATE and bad / n >= HIGH_BADRATE:
            s += W_BADRATE
            reasons.append(f"历史坏单率{100*bad//n}%({bad}/{n}单)")

        if len(sites) >= CROSSSITE_MIN:
            s += W_CROSSSITE
            reasons.append(f"跨{len(sites)}个站点")

        if n == 0 and this['cod'] and this['total'] >= p75:
            s += W_NEW_COD_BIG
            reasons.append(f"新客+COD+大额({this['total']:.0f})")

        if this['dt']:
            recent = [p for p in prior if p['dt'] and 0 <= (this['dt'] - p['dt']).total_seconds() <= VELOCITY_HOURS*3600]
            if len(recent) + 1 >= VELOCITY_MIN:
                s += W_VELOCITY
                reasons.append(f"{VELOCITY_HOURS}h内{len(recent)+1}单")

        return min(s, 100), reasons

    # 时点回测：对每个身份的每一单，用其之前的历史算分
    scored = []          # (score, reasons, order, identity_key, this_cls)
    for key, orders in identities.items():
        for i, this in enumerate(orders):
            prior = orders[:i]
            sc, rs = score_order(prior, this)
            scored.append((sc, rs, this['row'], key, this['cls']))

    def bucket(s):
        return 'HIGH' if s >= HIGH_THRESHOLD else ('MED' if s >= MED_THRESHOLD else 'LOW')

    # 回测指标：真正变坏的订单(undelivered/problem_return/cheat)里，有多少被提前打了高/中分
    bad_orders = [x for x in scored if x[4] in ('undelivered', 'problem_return', 'cheat')]
    flagged_high = sum(1 for x in bad_orders if x[0] >= HIGH_THRESHOLD)
    flagged_med = sum(1 for x in bad_orders if x[0] >= MED_THRESHOLD)
    all_high = [x for x in scored if x[0] >= HIGH_THRESHOLD]
    print()
    print("=" * 64)
    print("3. 时点回测 (只用历史算分，看能否提前命中坏单)")
    print("=" * 64)
    print(f"  真正变坏的订单(拒收/问题退货/cheat): {len(bad_orders)}")
    if bad_orders:
        print(f"    其中被提前标为 HIGH(>={HIGH_THRESHOLD}) : {flagged_high}  ({100*flagged_high//len(bad_orders)}% 召回)")
        print(f"    其中被提前标为 MED+ (>={MED_THRESHOLD})  : {flagged_med}  ({100*flagged_med//len(bad_orders)}% 召回)")
    print(f"  被标为 HIGH 的订单总数         : {len(all_high)}")
    if all_high:
        precision_bad = sum(1 for x in all_high if x[4] in ('undelivered', 'problem_return', 'cheat', 'soft_bad'))
        print(f"    其中确实是坏单(含软坏单)     : {precision_bad}  ({100*precision_bad//len(all_high)}% 命中率)")
    print("  说明: 时点评分只能抓'已经坑过你的回头客'。首次作案无历史，本就抓不到——这是诚实的下限。")

    # ============ 2. 高风险身份排行榜(用全量历史) ============
    print()
    print("=" * 64)
    print(f"2. 高风险身份排行榜 (今天就能行动的名单, 阈值 HIGH>={HIGH_THRESHOLD})")
    print("=" * 64)
    ident_rows = []
    for key, orders in identities.items():
        und = sum(1 for o in orders if o['cls'] == 'undelivered')
        prob = sum(1 for o in orders if o['cls'] == 'problem_return')
        cheat = sum(1 for o in orders if o['cls'] == 'cheat')
        soft = sum(1 for o in orders if o['cls'] == 'soft_bad')
        good = sum(1 for o in orders if o['cls'] == 'good')
        bad = und + prob + cheat + soft
        n = len(orders)
        sites = sorted(set(o['site'] for o in orders))
        loss = sum(float(o['row']['product_loss_amount'] or 0) + float(o['row']['shipping_loss_amount'] or 0) for o in orders)
        # 用"最后一单视角"打个代表分(把最后一单当作待评估单)
        rep_score, rep_reasons = score_order(orders[:-1], orders[-1]) if n >= 1 else (0, [])
        ident_rows.append({
            'identity': key, 'orders': n, 'good': good, 'bad': bad,
            'undelivered': und, 'problem_return': prob, 'cheat': cheat, 'soft_bad': soft,
            'bad_rate': round(bad / n, 2) if n else 0,
            'sites': len(sites), 'loss_recorded': round(loss, 2),
            'score': rep_score, 'reasons': '; '.join(rep_reasons),
            'sample_numbers': ','.join(str(o['row']['number']) for o in orders[:6]),
        })
    ident_rows.sort(key=lambda r: (-r['score'], -r['bad'], -r['bad_rate']))
    top = [r for r in ident_rows if r['score'] >= MED_THRESHOLD]
    print(f"  风险分 >= {MED_THRESHOLD} 的身份: {len(top)} 个")
    print(f"  {'身份(脱敏)':<20} {'分':>4} {'单':>3} {'坏':>3} {'拒收':>4} {'退货':>4} {'站点':>4}  理由")
    for r in top[:25]:
        masked = r['identity'][:4] + '***' + r['identity'][-3:]
        print(f"  {masked:<20} {r['score']:>4} {r['orders']:>3} {r['bad']:>3} "
              f"{r['undelivered']:>4} {r['problem_return']:>4} {r['sites']:>4}  {r['reasons'][:46]}")

    # ============ 导出 CSV ============
    os.makedirs('exports', exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    out = f'exports/fraud_risk_identities_{ts}.csv'
    with open(out, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=['identity', 'score', 'orders', 'good', 'bad', 'bad_rate',
                                          'undelivered', 'problem_return', 'cheat', 'soft_bad',
                                          'sites', 'loss_recorded', 'reasons', 'sample_numbers'])
        w.writeheader()
        for r in sorted(ident_rows, key=lambda r: -r['score']):
            w.writerow(r)
    print()
    print(f"完整身份风险表已导出: {out}  (可用 Excel 打开, 共 {len(ident_rows)} 行)")


if __name__ == '__main__':
    main()
