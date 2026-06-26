# 进销存(库存)模块说明

本系统为 woo-analysis 订单系统新增的一套「标准进销存」模块。设计目标:
**以本系统为库存真账本**——进货 / 卖出 / 拆单都在本系统记账,算出每仓「可用 =
现存 − 预留」,再定时下推到各 WooCommerce 站点。面向未来新市场可扩展(新增市场
只加数据、不改代码)。

> 全部功能位于 `inv_*` 新表与 `inv_*.py` 模块,**只读关联**现有表,**不改**
> 现有列含义,**不动**拉单核心(`sync_utils.py`)、撞号代理键(`oid_utils.py`)、
> 备份/灾备/blocklist/补发货逻辑。

## 1. 模块与文件

| 模块 | 文件 | 职责 |
|---|---|---|
| 迁移框架 | `inv_migrations.py` | 可回滚迁移(up/down/status),改动前自动备份 |
| 共享基础 | `inv_common.py` | 连接、审计入口 `record_movement`、权限闸门、角色仓库可见性 |
| 仓库管理 | `inv_warehouses.py` | 仓库增删改、自营/合伙人、市场→仓优先级路由 |
| SKU/映射 | `inv_skus.py` | SKU 主档、站点商品↔SKU 映射、未映射发现 |
| 解析核心 | `inv_resolver.py` | line_items→SKU 多级解析(复用现有品牌解析) |
| 台账/采购 | `inv_inventory.py` | 库存台账、出入库流水、供应商、采购入库、手工调整 |
| 批次/保质期 | `inv_batches.py` | 批次、FEFO 先到期先发、临期/过期 |
| 订单联动 | `inv_orders.py` | 幂等处理器:预留/出库/释放/退货 |
| 分仓引擎 | `inv_allocator.py` | 单仓优先、缺货拆单 |
| 下推 WP | `inv_push.py` / `inv_push_cron.py` | 可用量下推各站(复用 PM 白名单)+ 对账 |
| 通知中心 | `inv_notify.py` / `inv_notify_cron.py` | 补货/临期/滞销提醒,站内 + 邮件 |
| 报表 | `inv_reports.py` | 总览、库存货值、出入库汇总、分仓统计 |

## 2. 数据表(均 `inv_` 前缀)

`inv_skus`(SKU 主档)、`inv_site_sku_map`(站点商品↔SKU)、`inv_stock`(仓×SKU
现存/预留)、`inv_batches`(批次:生产日/到期日/单位成本/剩余量)、`inv_movements`
(出入库流水,只增不改,带操作人+时间+前后数量)、`inv_purchase_orders(+_items)`、
`inv_suppliers`、`inv_fulfillments(+_items)`(分单)、`inv_market_warehouses`
(市场→仓优先级)、`inv_warehouse_ext`(仓库扩展:自营/合伙人/partner_id)、
`inv_order_state`(订单库存联动状态)、`inv_push_logs`、`inv_notifications`。

对现有表唯一的「加法」:`users` 加 `can_view_inventory`/`can_manage_inventory`,
`inv_skus` 的 `reorder_point`(均为纯新增列,不改现有列含义)。

## 3. 核心不变式与原则

- **Σ `inv_batches.qty_remaining` == `inv_stock.on_hand`**(同仓同 SKU)始终成立。
- **可用 = on_hand − reserved**,实时计算,不落库。
- **所有库存写操作**经 `inv_common.record_movement`,自动写审计(操作人+时间+前后值)。
- **库存写操作幂等**:订单联动按 `inv_order_state.committed_json` 精确冲销旧效果。

## 4. 关键业务流

### 4.1 进货
采购单(草稿)→ 收货:每条明细建批次,写 `purchase_in` 流水,台账 on_hand 增加。

### 4.2 卖出 / 订单联动(幂等处理器)
不 hook 进拉单核心,而是按订单当前状态把库存对齐:
- `processing` / `on-hold` → **预留**(reserved 增加)
- `shipped` / `completed` / `delivered` / `partial-shipped` → **出库**(扣 on_hand + FEFO 批次)
- `cancelled` / `failed` / `refunded` / … → **释放预留** 或 **退货入库**(按当前状态)

触发方式:订单库存联动页手动/批量,或独立 cron(不改同步核心)。

### 4.3 自动分仓 / 拆单
订单 line_items 解析到 SKU → 按市场(收货国)仓优先级 + 各仓实时可用:
**优先单仓凑齐整单(少拆包),缺货才拆单**,分仓生成 fulfillment 并预留/出库。
- 该市场**已配置**路由(`inv_market_warehouses`)→ 用分仓引擎;
- **未配置**→ 回退订单既有 `warehouse_id`(单仓),保证老市场照旧。
- 新市场只需在「仓库管理」加路由,自动启用分仓,无需改代码。

捷克(CZ)已配置:HU 自营仓(优先)> PL 合伙人仓(金毅金谷,次选,缺货拆单)。

### 4.4 下推 WordPress
每站每映射商品可发布量 = `floor(Σ服务仓 max(0,on_hand−reserved) / qty_per_item)`。
写入复用 Product Manager 的 PUT 白名单(仅 `manage_stock`/`stock_quantity`)。

## 5. 角色与权限

| 角色 | 范围 | 读写 |
|---|---|---|
| 管理员(admin) | 全部仓 | 读写 + 主数据 |
| 库存管理(can_manage_inventory) | 全部仓 | 读写 |
| 库存查看(can_view_inventory) | 全部仓 | 只读 |
| 发货员(can_ship) | 全部仓 | 只读 |
| 合伙人(partner_users) | **只看自己仓** | 只读 |

合伙人通过「仓库管理」里把合伙人仓**关联合伙人账户**(`inv_warehouse_ext.partner_id`)
+ 现有 `partner_users` 绑定实现只读自己仓。

## 6. 运维 / 定时任务

```bash
# 迁移(改动前自动备份 *.db.pre*)
venv/bin/python inv_migrations.py status      # 查看
venv/bin/python inv_migrations.py up          # 应用
venv/bin/python inv_migrations.py down 001    # 回滚到 001 之前(含)

# 库存下推(与 15 分钟拉单错峰,例如每 30 分钟)
*/30 * * * * cd /www/wwwroot/woo-analysis && venv/bin/python inv_push_cron.py >> inv_push.log 2>&1

# 通知扫描(每天)
0 9 * * * cd /www/wwwroot/woo-analysis && venv/bin/python inv_notify_cron.py >> inv_notify.log 2>&1

# 整体自测(只读副本上跑端到端)
venv/bin/python inv_selftest.py

# 邮件(可选)在 settings 表配置:inv_smtp_host/port/user/pass/from/to/ssl
```

## 7. 导航入口

顶栏「库存管理」下拉:库存台账 / 批次保质期 / 订单库存联动 / 库存下推对账 /
通知中心 / 库存报表 / 仓库管理 / SKU 主档。右上角铃铛显示未读提醒。
