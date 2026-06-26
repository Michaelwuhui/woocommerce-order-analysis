"""
inv_migrations.py — 进销存模块的可回滚数据库迁移。

为什么自己写一个迷你迁移器而不用 Alembic:
  - 现有项目是 sqlite 单库 + 零迁移框架,引入 Alembic 成本过高。
  - 我们只需要:有序应用 / 按序回滚 / 记录已应用版本 / 改动前自动备份。

约束(对齐用户锁定的硬约束):
  - 只新建 inv_* 表;绝不 ALTER 现有表的列含义、绝不删现有表的列或数据。
  - 唯一对现有表的"加列"是给 users 加 can_view_inventory / can_manage_inventory
    两个权限开关(纯新增布尔列,不影响任何现有列含义)—— 这是可接受的加法。
  - 每个迁移提供 up(conn) 与 down(conn),down 必须能干净回滚 up。
  - up/down 前自动把 DB 复制成 *.db.pre<version> 快照(.gitignore 已忽略 *.db.pre*)。

用法:
  venv/bin/python inv_migrations.py status     # 查看已应用/待应用
  venv/bin/python inv_migrations.py up          # 应用所有待应用迁移
  venv/bin/python inv_migrations.py up 001       # 应用到(含)001
  venv/bin/python inv_migrations.py down         # 回滚最近一个迁移
  venv/bin/python inv_migrations.py down 001     # 回滚回到(含)001 之前 -> 即回滚 001 及之后全部
"""

import sys
import shutil
import datetime
import sqlite3

from inv_common import DB_FILE, get_conn


# ───────────────────────── 迁移注册表 ─────────────────────────
# 每个迁移是 (version, name, up_fn, down_fn)。version 字符串保证字典序 = 应用序。

def _ensure_meta(conn):
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_schema_migrations (
            version    TEXT PRIMARY KEY,
            name       TEXT,
            applied_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()


def _applied(conn):
    _ensure_meta(conn)
    return {r['version'] for r in
            conn.execute('SELECT version FROM inv_schema_migrations').fetchall()}


# ───────────────── 001: 全套 inv_ 核心表 ─────────────────

def up_001(conn):
    """创建全套库存核心表 + 给 users 加两个权限开关。幂等(IF NOT EXISTS)。"""

    # ── SKU 主档:库存最小核算单位 ──
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_skus (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            sku_code        TEXT NOT NULL UNIQUE,
            name            TEXT NOT NULL,
            brand_id        INTEGER,          -- 软关联 brands.id(只读)
            series_id       INTEGER,          -- 软关联 series.id(只读)
            puff_count      INTEGER,
            flavor          TEXT,
            barcode         TEXT,
            unit            TEXT DEFAULT 'pcs',
            shelf_life_days INTEGER,          -- 默认保质期天数(批次未填到期日时回退用)
            is_active       INTEGER DEFAULT 1,
            notes           TEXT,
            created_at      TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at      TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_skus_brand ON inv_skus(brand_id, series_id)')

    # ── 站点 WC 商品 ↔ SKU 映射 ──
    # 一个 WC 商品(可含变体)对应一个 SKU;bundle 用 qty_per_item 表达"一件=N支"。
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_site_sku_map (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id         INTEGER NOT NULL,  -- 软关联 sites.id
            wc_product_id   INTEGER,           -- WooCommerce 商品 id
            wc_variation_id INTEGER DEFAULT 0, -- 变体 id(无变体=0)
            wc_sku          TEXT,              -- WC 自带 sku 字符串(若有)
            raw_name        TEXT,              -- line item 名称(兜底匹配)
            sku_id          INTEGER NOT NULL,  -- -> inv_skus.id
            qty_per_item    INTEGER DEFAULT 1, -- 每件 WC 商品折合多少个 SKU 单位
            is_active       INTEGER DEFAULT 1,
            created_at      TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at      TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(site_id, wc_product_id, wc_variation_id),
            FOREIGN KEY (sku_id) REFERENCES inv_skus(id) ON DELETE CASCADE
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_ssm_site ON inv_site_sku_map(site_id)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_ssm_sku ON inv_site_sku_map(sku_id)')

    # ── 仓库扩展属性(不动现有 warehouses 表) ──
    # ownership_type: 'self' 自营 | 'partner' 合伙人。partner_name 记合伙人名(如 金毅金谷)。
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_warehouse_ext (
            warehouse_id   INTEGER PRIMARY KEY,   -- 软关联 warehouses.id
            ownership_type TEXT NOT NULL DEFAULT 'self',
            partner_name   TEXT,
            partner_id     INTEGER,               -- 可选软关联 partners.id(未来对账打通用)
            region         TEXT,                  -- 备注性质的地理大区(可空)
            is_fulfillment INTEGER DEFAULT 1,     -- 是否参与自动分仓(0=只记账不参与发货路由)
            notes          TEXT,
            created_at     TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at     TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # ── 市场 → 仓库优先级(未来新市场只加这张表的行) ──
    # market_code 通常是目的市场国家码(CZ/HU/PL/AU...)。priority 越小越优先。
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_market_warehouses (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            market_code  TEXT NOT NULL,
            warehouse_id INTEGER NOT NULL,        -- 软关联 warehouses.id
            priority     INTEGER NOT NULL DEFAULT 100,
            is_active    INTEGER DEFAULT 1,
            notes        TEXT,
            created_at   TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at   TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(market_code, warehouse_id)
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_mw_market ON inv_market_warehouses(market_code, priority)')

    # ── 库存台账:仓 × SKU 的现存/预留(可用 = on_hand - reserved,不落库实时算) ──
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_stock (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            warehouse_id INTEGER NOT NULL,
            sku_id       INTEGER NOT NULL,
            on_hand      INTEGER NOT NULL DEFAULT 0,
            reserved     INTEGER NOT NULL DEFAULT 0,
            updated_at   TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(warehouse_id, sku_id),
            FOREIGN KEY (sku_id) REFERENCES inv_skus(id) ON DELETE CASCADE
        )
    ''')

    # ── 供应商 ──
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_suppliers (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            name       TEXT NOT NULL,
            contact    TEXT,
            phone      TEXT,
            email      TEXT,
            address    TEXT,
            currency   TEXT DEFAULT 'CNY',
            is_active  INTEGER DEFAULT 1,
            notes      TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # ── 采购单 + 明细 ──
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_purchase_orders (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            po_no         TEXT UNIQUE,
            supplier_id   INTEGER,
            warehouse_id  INTEGER NOT NULL,
            status        TEXT NOT NULL DEFAULT 'draft',  -- draft/received/cancelled
            order_date    TEXT,
            received_date TEXT,
            currency      TEXT DEFAULT 'CNY',
            total_amount  REAL DEFAULT 0,
            operator_id   INTEGER,
            operator_name TEXT,
            note          TEXT,
            created_at    TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at    TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_purchase_order_items (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            po_id           INTEGER NOT NULL,
            sku_id          INTEGER NOT NULL,
            qty             INTEGER NOT NULL DEFAULT 0,
            unit_cost       REAL DEFAULT 0,
            batch_no        TEXT,
            production_date TEXT,
            expiry_date     TEXT,
            received_qty    INTEGER DEFAULT 0,
            note            TEXT,
            FOREIGN KEY (po_id) REFERENCES inv_purchase_orders(id) ON DELETE CASCADE,
            FOREIGN KEY (sku_id) REFERENCES inv_skus(id)
        )
    ''')

    # ── 批次(生产日/到期日/单位成本/剩余量),FEFO 先到期先发 ──
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_batches (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            warehouse_id      INTEGER NOT NULL,
            sku_id            INTEGER NOT NULL,
            batch_no          TEXT,
            production_date   TEXT,
            expiry_date       TEXT,
            unit_cost         REAL DEFAULT 0,
            cost_currency     TEXT DEFAULT 'CNY',
            qty_received      INTEGER NOT NULL DEFAULT 0,
            qty_remaining     INTEGER NOT NULL DEFAULT 0,
            purchase_order_id INTEGER,
            created_at        TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at        TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (sku_id) REFERENCES inv_skus(id) ON DELETE CASCADE
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_batches_fefo ON inv_batches(warehouse_id, sku_id, expiry_date)')

    # ── 出入库流水(append-only 审计,只增不改) ──
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_movements (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            ts              TEXT DEFAULT CURRENT_TIMESTAMP,
            warehouse_id    INTEGER NOT NULL,
            sku_id          INTEGER NOT NULL,
            batch_id        INTEGER,
            movement_type   TEXT NOT NULL,
            qty_delta       INTEGER NOT NULL DEFAULT 0,   -- 对 on_hand 的有符号增量
            reserved_delta  INTEGER NOT NULL DEFAULT 0,   -- 对 reserved 的有符号增量
            qty_before      INTEGER,
            qty_after       INTEGER,
            reserved_before INTEGER,
            reserved_after  INTEGER,
            ref_type        TEXT,          -- order/po/fulfillment/manual...
            ref_id          TEXT,
            order_id        TEXT,          -- 软关联 orders.id(跨站撞号代理键),可空
            operator_id     INTEGER,
            operator_name   TEXT,
            note            TEXT
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_mov_wh_sku ON inv_movements(warehouse_id, sku_id, ts)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_mov_order ON inv_movements(order_id)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_mov_ref ON inv_movements(ref_type, ref_id)')

    # ── 分单(fulfillment)+ 明细 ──
    # 一个订单可拆成多个 fulfillment(不同仓发货)。
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_fulfillments (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            order_id        TEXT NOT NULL,    -- 软关联 orders.id(<site_id>-<woo_id>)
            source          TEXT,             -- 站点 url(冗余,便于查询)
            warehouse_id    INTEGER NOT NULL,
            status          TEXT NOT NULL DEFAULT 'planned',  -- planned/reserved/shipped/cancelled
            is_split        INTEGER DEFAULT 0,
            split_index     INTEGER DEFAULT 1,
            split_total     INTEGER DEFAULT 1,
            tracking_number TEXT,
            carrier_slug    TEXT,
            operator_id     INTEGER,
            operator_name   TEXT,
            note            TEXT,
            created_at      TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at      TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_ful_order ON inv_fulfillments(order_id)')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_fulfillment_items (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            fulfillment_id INTEGER NOT NULL,
            sku_id         INTEGER NOT NULL,
            qty            INTEGER NOT NULL DEFAULT 0,
            batch_id       INTEGER,
            note           TEXT,
            FOREIGN KEY (fulfillment_id) REFERENCES inv_fulfillments(id) ON DELETE CASCADE,
            FOREIGN KEY (sku_id) REFERENCES inv_skus(id)
        )
    ''')

    # ── 给 users 加两个库存权限开关(纯新增布尔列,不改现有列含义) ──
    for col in ('can_view_inventory', 'can_manage_inventory'):
        try:
            conn.execute(f'ALTER TABLE users ADD COLUMN {col} INTEGER DEFAULT 0')
        except sqlite3.OperationalError:
            pass  # 已存在

    conn.commit()


def down_001(conn):
    """回滚 001:删除所有 inv_* 核心表。

    注意:不删 users 上新增的两个布尔列 —— sqlite 早期不支持 DROP COLUMN,
    且删列有改动现有表结构风险;这两个纯新增列留着无害(默认 0),符合
    "禁止手删表的列或数据" 的硬约束。
    """
    for t in (
        'inv_fulfillment_items', 'inv_fulfillments',
        'inv_movements', 'inv_batches',
        'inv_purchase_order_items', 'inv_purchase_orders', 'inv_suppliers',
        'inv_stock', 'inv_market_warehouses', 'inv_warehouse_ext',
        'inv_site_sku_map', 'inv_skus',
    ):
        conn.execute(f'DROP TABLE IF EXISTS {t}')
    conn.commit()


# ───────────────── 002: 种子数据(匈牙利仓 + 波兰合伙人 + CZ/HU 市场) ─────────────────

def up_002(conn):
    """业务种子:
       1) 建匈牙利(HU)自营仓(若不存在);
       2) 标波兰仓为合伙人(金毅金谷)仓,匈牙利仓为自营;
       3) 配置市场路由:CZ -> [HU 优先, PL 次之];HU -> [HU];PL -> [PL];
       全部幂等。
    """
    # 1) 匈牙利自营仓
    hu = conn.execute("SELECT id FROM warehouses WHERE code='HU' OR country='HU'").fetchone()
    if not hu:
        conn.execute(
            'INSERT INTO warehouses (name, code, country, default_currency, notes) '
            "VALUES ('匈牙利仓库', 'HU', 'HU', 'HUF', '本系统自营仓(库存真账本)')")
        hu_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
    else:
        hu_id = hu['id']

    # 波兰仓:取 code='PL' 或 country='PL' 的第一个(现有 id=1 波兰仓库)
    pl = conn.execute("SELECT id FROM warehouses WHERE code='PL' OR country='PL' ORDER BY id LIMIT 1").fetchone()
    pl_id = pl['id'] if pl else None

    # 2) 仓库扩展:HU 自营,PL 合伙人(金毅金谷)
    conn.execute('''INSERT INTO inv_warehouse_ext (warehouse_id, ownership_type, partner_name, notes)
                    VALUES (?, 'self', NULL, '匈牙利自营仓')
                    ON CONFLICT(warehouse_id) DO UPDATE SET ownership_type='self', updated_at=CURRENT_TIMESTAMP''',
                 (hu_id,))
    if pl_id:
        conn.execute('''INSERT INTO inv_warehouse_ext (warehouse_id, ownership_type, partner_name, notes)
                        VALUES (?, 'partner', '金毅金谷', '波兰合伙人仓(金毅金谷)')
                        ON CONFLICT(warehouse_id) DO UPDATE SET
                            ownership_type='partner', partner_name='金毅金谷', updated_at=CURRENT_TIMESTAMP''',
                     (pl_id,))

    # 3) 市场 → 仓优先级
    def set_route(market, warehouse_id, priority, note):
        if not warehouse_id:
            return
        conn.execute('''INSERT INTO inv_market_warehouses (market_code, warehouse_id, priority, notes)
                        VALUES (?,?,?,?)
                        ON CONFLICT(market_code, warehouse_id) DO UPDATE SET
                            priority=excluded.priority, is_active=1, updated_at=CURRENT_TIMESTAMP''',
                     (market, warehouse_id, priority, note))

    set_route('CZ', hu_id, 10, '捷克优先匈牙利仓(就近)')
    set_route('CZ', pl_id, 20, '捷克次选波兰合伙人仓(缺货拆单)')
    set_route('HU', hu_id, 10, '匈牙利本地仓')
    set_route('PL', pl_id, 10, '波兰本地仓')
    conn.commit()


def down_002(conn):
    """回滚 002:删除本迁移写入的市场路由与仓库扩展标记;
       匈牙利仓是否删除取决于是否由本迁移创建 —— 为稳妥起见保留仓库本身
       (避免误删可能已被引用的主数据),仅清掉扩展属性与市场路由。
    """
    conn.execute("DELETE FROM inv_market_warehouses WHERE market_code IN ('CZ','HU','PL')")
    # 仅删除我们标记过的扩展行(HU 自营 / PL 合伙人)
    conn.execute('''DELETE FROM inv_warehouse_ext WHERE warehouse_id IN (
                        SELECT id FROM warehouses WHERE code IN ('HU','PL') OR country IN ('HU','PL'))''')
    conn.commit()


# ───────────────── 003: 订单库存联动状态表 ─────────────────

def up_003(conn):
    """inv_order_state:记录每张订单当前的库存联动状态,使卖出扣减幂等。

    committed_json 保存当前已提交的明细与批次分配,以便订单状态回退时精确冲销。
    inv_state: reserved(已预留) | shipped(已出库) | released(已释放) |
               returned(已退货入库) | blocked(有未映射,无法扣减) | none。
    """
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_order_state (
            order_id          TEXT PRIMARY KEY,   -- 软关联 orders.id(<site>-<woo>)
            source            TEXT,
            inv_state         TEXT NOT NULL DEFAULT 'none',
            warehouse_id      INTEGER,
            committed_json    TEXT,               -- {mode, warehouse_id, lines:[{sku_id,qty}], batches:[{batch_id,qty}]}
            last_order_status TEXT,
            market_code       TEXT,
            note              TEXT,
            processed_at      TEXT,
            created_at        TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at        TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_ostate_state ON inv_order_state(inv_state)')
    conn.commit()


def down_003(conn):
    conn.execute('DROP TABLE IF EXISTS inv_order_state')
    conn.commit()


# ───────────────── 004: 库存下推 WP 日志 ─────────────────

def up_004(conn):
    """inv_push_logs:记录每次把可用库存下推到各站 WC 商品的结果(审计 + 对账)。"""
    conn.execute('''
        CREATE TABLE IF NOT EXISTS inv_push_logs (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            ts              TEXT DEFAULT CURRENT_TIMESTAMP,
            site_id         INTEGER,
            source          TEXT,
            wc_product_id   INTEGER,
            wc_variation_id INTEGER DEFAULT 0,
            sku_id          INTEGER,
            prev_qty        INTEGER,   -- 推送前 WC 库存(对账时填,纯推送可空)
            pushed_qty      INTEGER,   -- 本次下推的可用数量
            status          TEXT,      -- ok / error / skipped / dry
            error           TEXT,
            operator_id     INTEGER,
            operator_name   TEXT
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_inv_push_site ON inv_push_logs(site_id, ts)')
    conn.commit()


def down_004(conn):
    conn.execute('DROP TABLE IF EXISTS inv_push_logs')
    conn.commit()


MIGRATIONS = [
    ('001', 'core_inv_schema', up_001, down_001),
    ('002', 'seed_hu_pl_markets', up_002, down_002),
    ('003', 'order_inventory_state', up_003, down_003),
    ('004', 'inv_push_logs', up_004, down_004),
]


# ───────────────────────── 运行器 ─────────────────────────

def _backup(tag):
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    dst = f'{DB_FILE}.pre{tag}_{stamp}'
    try:
        shutil.copy2(DB_FILE, dst)
        print(f'  [backup] {DB_FILE} -> {dst}')
    except FileNotFoundError:
        print(f'  [backup] 跳过:{DB_FILE} 不存在')


def cmd_status():
    conn = get_conn()
    applied = _applied(conn)
    print(f'数据库: {DB_FILE}')
    print('版本   状态        名称')
    for v, name, _, _ in MIGRATIONS:
        mark = '已应用' if v in applied else '待应用'
        print(f'{v}    {mark}      {name}')
    conn.close()


def cmd_up(target=None):
    conn = get_conn()
    applied = _applied(conn)
    pending = [m for m in MIGRATIONS if m[0] not in applied
               and (target is None or m[0] <= target)]
    if not pending:
        print('没有待应用的迁移。')
        conn.close()
        return
    _backup('up' + pending[0][0])
    for v, name, up_fn, _ in pending:
        print(f'应用 {v} {name} ...')
        up_fn(conn)
        conn.execute('INSERT OR REPLACE INTO inv_schema_migrations (version, name) VALUES (?,?)',
                     (v, name))
        conn.commit()
        print(f'  ✓ {v} 完成')
    conn.close()
    print('全部应用完毕。')


def cmd_down(target=None):
    """回滚。target=None 只回滚最近一个;target=版本号 回滚 >=该版本 的全部。"""
    conn = get_conn()
    applied = _applied(conn)
    to_rollback = sorted([m for m in MIGRATIONS if m[0] in applied],
                         key=lambda m: m[0], reverse=True)
    if target is None:
        to_rollback = to_rollback[:1]
    else:
        to_rollback = [m for m in to_rollback if m[0] >= target]
    if not to_rollback:
        print('没有可回滚的迁移。')
        conn.close()
        return
    _backup('down' + to_rollback[0][0])
    for v, name, _, down_fn in to_rollback:
        print(f'回滚 {v} {name} ...')
        down_fn(conn)
        conn.execute('DELETE FROM inv_schema_migrations WHERE version=?', (v,))
        conn.commit()
        print(f'  ✓ {v} 已回滚')
    conn.close()
    print('回滚完毕。')


if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'status'
    arg = sys.argv[2] if len(sys.argv) > 2 else None
    if cmd == 'status':
        cmd_status()
    elif cmd == 'up':
        cmd_up(arg)
    elif cmd == 'down':
        cmd_down(arg)
    else:
        print(__doc__)
        sys.exit(1)
