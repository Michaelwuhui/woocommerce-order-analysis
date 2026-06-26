"""
inv_common.py — 进销存(库存)模块共享基础设施。

设计原则(对齐已锁定架构):
  - 库存真源 = 本系统。所有库存/批次/出入库都在 inv_* 新表记账。
  - inv_* 表只读关联现有表(sites / warehouses / brands / series / orders),
    绝不修改现有表的列含义,绝不手删现有数据。
  - 所有库存/金额写操作必须写审计到 inv_movements(操作人 + 时间 + 前后数量)。

本模块只放"被多个 inv_* 模块复用"的东西:数据库连接、审计流水写入、
权限小工具。具体业务逻辑放到各自的 inv_xxx.py。

故意不 import app.py —— 避免与 22000 行单体应用产生循环依赖。
连接复用同一个 DB_FILE(相对路径,跟随当前工作目录,与 app.py 一致)。
"""

import os
import sqlite3
from functools import wraps

# 与 app.py 保持一致:相对路径,跟随进程工作目录。
# 允许用环境变量覆盖(自测时可指向副本),默认生产/副本同名文件。
DB_FILE = os.environ.get('INV_DB_FILE', 'woocommerce_orders.db')


def get_conn():
    """返回一个 row_factory=Row 的 sqlite 连接。调用方负责 close()。

    与 app.get_db_connection() 行为一致,但不依赖 app.py。
    """
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    # 外键约束:inv_* 之间的 ON DELETE CASCADE 需要它生效。
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


# ───────────────────────────── 审计流水 ─────────────────────────────
# 任何改变库存(on_hand)或预留(reserved)的写操作,都必须通过这里落一条
# inv_movements。这是库存账本的唯一真相来源:库存表 inv_stock 只是流水的
# 物化汇总,可由 inv_movements 完全重算。

# 合法的流水类型(movement_type)。新增类型时在此登记,便于报表归类。
MOVEMENT_TYPES = {
    'purchase_in',   # 采购入库
    'sale_out',      # 销售出库(订单发货扣减)
    'adjust',        # 手工盘点调整(可正可负)
    'reserve',       # 预留(下单占用,on_hand 不变、reserved 增加)
    'release',       # 释放预留(取消/拆单调整)
    'transfer_out',  # 调拨出库
    'transfer_in',   # 调拨入库
    'return_in',     # 退货入库
    'init',          # 期初建账
}


def record_movement(conn, *, warehouse_id, sku_id, movement_type,
                    qty_delta=0, reserved_delta=0,
                    batch_id=None, ref_type=None, ref_id=None, order_id=None,
                    operator_id=None, operator_name=None, note=None,
                    apply_stock=True):
    """写一条出入库流水,并(默认)同步更新 inv_stock 的现存/预留。

    这是库存写操作的唯一入口。它会:
      1. 读取该 (warehouse, sku) 当前的 on_hand / reserved(前值);
      2. 应用 qty_delta(改 on_hand)与 reserved_delta(改 reserved);
      3. 写入带前后值的 inv_movements 审计记录;
      4. (apply_stock=True 时)upsert inv_stock。

    参数:
      qty_delta:      对 on_hand(现存)的有符号增量。入库为正,出库为负。
      reserved_delta: 对 reserved(预留)的有符号增量。
      apply_stock:    False 时只记流水不动库存(用于重算/对账场景)。

    返回: 新建 movement 的 id。
    注意: 调用方负责 conn.commit() —— 以便与同事务的其它写操作一起提交/回滚。
    """
    if movement_type not in MOVEMENT_TYPES:
        raise ValueError(f'未知的 movement_type: {movement_type}')

    row = conn.execute(
        'SELECT on_hand, reserved FROM inv_stock WHERE warehouse_id=? AND sku_id=?',
        (warehouse_id, sku_id)
    ).fetchone()
    on_hand_before = row['on_hand'] if row else 0
    reserved_before = row['reserved'] if row else 0
    on_hand_after = on_hand_before + (qty_delta or 0)
    reserved_after = reserved_before + (reserved_delta or 0)

    cur = conn.execute('''
        INSERT INTO inv_movements
            (warehouse_id, sku_id, batch_id, movement_type,
             qty_delta, reserved_delta,
             qty_before, qty_after, reserved_before, reserved_after,
             ref_type, ref_id, order_id, operator_id, operator_name, note)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ''', (warehouse_id, sku_id, batch_id, movement_type,
          qty_delta or 0, reserved_delta or 0,
          on_hand_before, on_hand_after, reserved_before, reserved_after,
          ref_type, ref_id, order_id, operator_id, operator_name, note))
    movement_id = cur.lastrowid

    if apply_stock:
        if row:
            conn.execute(
                'UPDATE inv_stock SET on_hand=?, reserved=?, updated_at=CURRENT_TIMESTAMP '
                'WHERE warehouse_id=? AND sku_id=?',
                (on_hand_after, reserved_after, warehouse_id, sku_id))
        else:
            conn.execute(
                'INSERT INTO inv_stock (warehouse_id, sku_id, on_hand, reserved) '
                'VALUES (?,?,?,?)',
                (warehouse_id, sku_id, on_hand_after, reserved_after))

    return movement_id


def current_operator():
    """返回 (user_id, user_name) 用于审计。在请求上下文外返回 (None, None)。"""
    try:
        from flask_login import current_user
        if getattr(current_user, 'is_authenticated', False):
            return current_user.id, (current_user.name or current_user.username)
    except Exception:
        pass
    return None, None


# ───────────────────────────── 权限小工具 ─────────────────────────────
# 库存模块整体读权限 = can_view_inventory;写权限 = can_manage_inventory。
# 仓库主数据(增删仓/标合伙人/市场路由)属管理员级,用 inv_admin_required。
# 详细的角色视图在模块 9 收口;此处先提供基础闸门,字段在迁移中加到 users。

def _user_flag(flag):
    """读取当前用户某个 can_* 标志(admin 用户名永远放行)。"""
    from flask_login import current_user
    if not getattr(current_user, 'is_authenticated', False):
        return False
    if getattr(current_user, 'username', None) == 'admin':
        return True
    try:
        conn = get_conn()
        u = conn.execute(f'SELECT {flag} FROM users WHERE id=?',
                         (current_user.id,)).fetchone()
        conn.close()
        return bool(u and u[flag] == 1)
    except Exception:
        return False


def can_view_inventory():
    return _user_flag('can_view_inventory')


def can_manage_inventory():
    return _user_flag('can_manage_inventory')


def _deny(message='您没有权限访问库存管理。', code=403, json=False):
    from flask import jsonify, render_template_string
    if json:
        return jsonify({'error': message}), code
    return render_template_string('''
<!DOCTYPE html><html lang="zh"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0"><title>权限不足</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css" rel="stylesheet">
<style>body{background:#0f172a;color:#e2e8f0;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;}
.box{background:#1e293b;border:1px solid #334155;border-radius:16px;padding:48px;text-align:center;max-width:440px;}
.icon{font-size:56px;color:#f59e0b;margin-bottom:20px;}h1{font-size:22px;font-weight:600;margin-bottom:12px;}
p{color:#94a3b8;margin-bottom:24px;}a{background:#3b82f6;color:#fff;padding:10px 24px;border-radius:8px;text-decoration:none;}</style>
</head><body><div class="box"><div class="icon"><i class="bi bi-shield-exclamation"></i></div>
<h1>权限不足</h1><p>{{ msg }}</p><a href="/"><i class="bi bi-house"></i> 返回首页</a></div></body></html>
''', msg=message), code


def inv_view_required(f):
    """页面级:需要库存查看权限(can_view_inventory 或更高)。"""
    @wraps(f)
    def wrapper(*args, **kwargs):
        from flask_login import current_user
        from flask import redirect, url_for
        if not getattr(current_user, 'is_authenticated', False):
            return redirect(url_for('login'))
        if not (can_view_inventory() or can_manage_inventory()):
            return _deny('库存管理需要授权,请联系管理员。')
        return f(*args, **kwargs)
    return wrapper


def inv_manage_required(f):
    """写操作级:需要库存管理权限(can_manage_inventory)。返回 JSON 403。"""
    @wraps(f)
    def wrapper(*args, **kwargs):
        from flask_login import current_user
        if not getattr(current_user, 'is_authenticated', False):
            return _deny('请先登录。', code=401, json=True)
        if not can_manage_inventory():
            return _deny('您没有库存写入权限。', json=True)
        return f(*args, **kwargs)
    return wrapper


def inv_admin_required(f):
    """主数据级(仓库/市场路由):仅超级管理员。返回 JSON 403。"""
    @wraps(f)
    def wrapper(*args, **kwargs):
        from flask_login import current_user
        if not getattr(current_user, 'is_authenticated', False):
            return _deny('请先登录。', code=401, json=True)
        is_admin = (getattr(current_user, 'username', None) == 'admin'
                    or (hasattr(current_user, 'is_admin') and current_user.is_admin()))
        if not is_admin:
            return _deny('仓库主数据仅限管理员管理。', json=True)
        return f(*args, **kwargs)
    return wrapper
