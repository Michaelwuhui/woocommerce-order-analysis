# Woo Analysis 订单与履约系统

当前版本：`2.0.1`（2026-08-21）

Woo Analysis 是面向多 WooCommerce 独立站的内部订单运营系统。它把订单同步、订单分析、人工发货、库存映射、多仓履约、物流跟踪、客户通知、COD 分配与供应商对账集中在同一套数据和权限体系中。

本仓库不只是数据看板。`v2.0.0` 的核心变化是引入可审计的多仓履约领域模型，同时保留波兰人工合作仓“选择承运商 + 录入运单号”的简单操作体验。

> 本项目包含真实订单与外部系统集成能力。任何生产部署都必须使用仓库外的密钥文件、先备份 SQLite 数据库、先验证迁移，再开启 WMS 自动提交或客户通知。

## v2.0.0 重点能力

- 多 WooCommerce 站点增量同步、服务端分页、订单和客户分析。
- `Order → Fulfillment → Shipment → TrackingEvent` 多仓履约模型。
- 一个订单可拆到不同仓库，并分别拥有包裹、承运商和运单号。
- 波兰人工合作仓与金毅金谷临时中转仓可组成“联合发货组”：前台只操作一次、客户只看到一个包裹，后台仍按仓扣减各自库存。
- 临时中转仓只管理实际调拨库存，用完即缺货，不提供无限库存回退。
- 仓库 SKU → 订单系统标准 SKU → WooCommerce 商品/变体的两级映射和批量确认。
- WooCommerce AST/兼容插件的分批发货同步、站点语言客户备注和邮件通知。
- InPost、DPD、Packeta，以及匈牙利 Packeta/Express One 末端派送查询适配。
- 发货接口超时后的远端回读确认、本地 `pending_sync` 暂存、重复运单拦截和人工重试。
- 匈牙利 WMS 与新波兰 WMS 的独立适配器、幂等任务、审计和安全开关。
- 商品批量更新后远端回读验证，以及支持“作为全新商品克隆”的持久化后台任务。
- 订单图片通知、企业微信目标配置和邮件中心只读订单接口。

## 系统边界与当前启用状态

| 能力 | 当前策略 |
|---|---|
| 波兰人工合作仓 | 生产使用；合作方不维护 OMS 数量库存，人工录入承运商和运单号 |
| 金毅金谷临时中转仓 | 生产使用；只管理已调拨 SKU 的有限库存 |
| 联合发货 | 波兰人工合作仓与临时中转仓可合并为一个物理包裹 |
| 捷克/匈牙利人工发货 | 可由波兰合作方使用 Packeta；匈牙利长数字 PLC 单号走 Express One 查询 |
| 匈牙利 WMS | 适配代码保留，路由与自动提交默认关闭，等待库存和受控联调 |
| 新波兰 WMS | 独立适配器已具备；真实路由和自动提交受配置及 HTTP 风险开关保护 |
| 客户运费 | 逐订单读取 WooCommerce `shipping_total`，不写死金额、不在分仓包裹间平均 |
| 订单完成 | 所有未取消包裹妥投后，才允许聚合完成并同步 WooCommerce `completed` |

详细业务规则、状态机、COD 和 WMS 说明见 [MULTI_WAREHOUSE_FULFILLMENT.md](MULTI_WAREHOUSE_FULFILLMENT.md)。

## 架构概览

```text
浏览器
  │
Nginx / HTTPS
  │
Gunicorn + Flask (app.py)
  ├── 订单、报表、客户、商品和权限
  ├── 库存与仓库蓝图 (inv_*.py)
  ├── 多仓履约 API (fulfillment_api.py)
  ├── 发货、AST 与物流查询
  └── 通知、商品克隆和只读集成 API
  │
SQLite (woocommerce_orders.db)
  ├── WooCommerce REST API
  ├── 履约后台任务 / 商品克隆任务
  ├── 匈牙利 WMS / 新波兰 WMS（受开关控制）
  ├── InPost / Track718 / Packeta / Express One
  └── 客户邮件 / 企业微信通知
```

### 主要模块

| 模块 | 作用 |
|---|---|
| `app.py` | Flask 主应用、订单/发货/商品/报表路由及兼容逻辑 |
| `sync_utils.py`, `auto_sync.py` | WooCommerce 增量同步、断点和运行状态 |
| `inv_*.py` | 仓库、SKU、映射、库存流水、批次、补货和对账 |
| `fulfillment_service.py` | 分仓规划、状态转换、缺货、联合发货和聚合状态 |
| `fulfillment_worker.py` | WMS 提交、跟踪、Woo 同步、重试和对账任务 |
| `fulfillment_woocommerce.py` | Shipment 到 AST/WooCommerce 的幂等同步 |
| `order_shipments.py` | 已确认包裹识别、重复运单与跟踪格式判断 |
| `hungary_wms.py` | 匈牙利 WMS 适配器；不得复用于新波兰仓 |
| `poland_wms.py` | 新波兰仓 SZ56T/华磊协议适配器 |
| `carrier_tracking.py` | InPost、DPD、Packeta、Express One 和补充查询 |
| `product_clone_*.py` | 可恢复的商品克隆任务、SKU 冲突处理和后台 Worker |
| `order_notification_*.py` | 订单图片、目标路由、渲染、发送和审计 |
| `mail_center_readonly_api.py` | 使用文件令牌保护的订单只读接口 |
| `inv_migrations.py` | 库存与履约 schema 的可追踪迁移和回滚 |

## 履约领域模型

```text
Order
  └── OrderItem
        └── FulfillmentItem ──> Fulfillment（仓库责任）
                                  └── Shipment（物理包裹）
                                        ├── ShipmentItem
                                        └── TrackingEvent
```

- `Order` 是 WooCommerce 商业订单，不因分仓创建新的客户订单。
- `Fulfillment` 表示某个仓库对当前订单版本承担的履约责任。
- `FulfillmentItem` 固定订单行、SKU、数量和金额分配快照。
- `Shipment` 表示真实包裹，一个订单允许多个包裹和多个运单号。
- `oms_shipment_fulfillments` 允许一个联合发货包裹关联多个仓库履约单。
- `TrackingEvent` 保存官方或第三方物流事件；重复事件去重，乱序事件不会使状态回退。

### 典型状态

```text
订单同步
  → 待规划
  → 已分仓 / 缺货 / 人工处理
  → 待提交或人工待发货
  → 已提交 WMS / 已发货
  → 运输中
  → 已妥投
  → 全部包裹妥投后订单完成
```

取消、缺货、部分失败或“一个仓已发货、另一个仓待取消”的订单进入人工处理，不自动跨仓换货，也不静默吞掉异常。

## 分仓与联合发货规则

1. 分仓以已确认的仓库 SKU 和 WooCommerce 商品/变体映射为依据。
2. 专属仓商品缺货时醒目标记，等待人工换货、补货或取消决定。
3. 波兰人工合作仓不依赖数量库存；临时中转仓必须有可用库存。
4. 联合发货组内可以用一个运单发出一个订单的全部商品。
5. 联合包裹只向 WooCommerce/客户同步一次，但每个仓库的库存流水分别记账。
6. 已分配和已发货的历史订单不会因以后切换 WMS 路由而自动换仓，防止重复发货。

## WooCommerce 与物流同步

### 数据方向

- WooCommerce → 本系统：订单、订单行、地址、币种、运费、支付/COD、状态、备注和站点来源。
- 本系统 → WooCommerce：运单、承运商、包裹商品、发货状态、客户备注和最终完成状态。
- WMS → 本系统：外部单号、面单、运单及轨迹；没有回调的接口由后台任务轮询。
- 官方物流/第三方查询 → `TrackingEvent`：保存事件来源、时间和原始摘要。

### 发货可靠性

- 远端写入使用确定性业务编号或幂等键。
- HTTP 超时不直接判定失败，先通过 WooCommerce/WMS 查询确认是否已写入。
- 无法确认的首次发货可保存为 `pending_sync`，不会冒充已发货包裹。
- 相同承运商和运单号不能重复用于其他订单。
- AST 分批发货按包裹同步商品明细；联合发货只生成一个客户可见包裹。
- 客户备注按站点语言输出波兰语、捷克语、匈牙利语或英语；无法识别时使用英语，不发送中文运营说明。

## 权限模型

权限同时作用于页面和 API，不能只依赖前端隐藏按钮。

- 管理员：系统配置、用户、仓库、映射、任务恢复和审计。
- 站点权限：普通用户只能访问被授权站点的订单、客户和商品。
- 仓库行级权限：仓库员工只能查看和处理被授权仓库的履约单与库存。
- 人工发货员：使用原有发货弹窗，看到本次发货仓提示，只录入承运商和运单号。
- 管理型操作：库存调整、映射批量确认、WMS 配置、手工重试等使用独立权限。
- 邮件中心只读 API：独立文件令牌、站点白名单和只读数据库连接，不复用 Web 登录态。

## 数据库与迁移

默认数据库是项目目录下的 `woocommerce_orders.db`。Worker 可通过 `OMS_DB_FILE`/`INV_DB_FILE` 指向同一数据库。

基础订单表由主初始化脚本和应用启动过程维护；库存与履约表由 `inv_migrations.py` 管理。当前迁移版本到 `020 joint_dispatch_groups`。

```bash
# 查看迁移状态（只读）
python inv_migrations.py status

# 备份后应用全部待执行迁移
python inv_migrations.py up

# 仅回滚最近一个迁移；生产执行前必须同时准备代码回滚
python inv_migrations.py down
```

迁移命令会生成数据库副本，但生产发布仍应先用 SQLite 在线备份脚本制作独立、校验过的备份。

## 安装与本地运行

### 要求

- Linux 生产环境；Windows 可用于开发和测试。
- Python 3.10+。
- SQLite 3、Nginx、Gunicorn。
- 订单邮件截图启用时需要 Playwright Chromium 和支持中日韩文字的字体。

### 安装

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt

# 订单邮件精确渲染需要
python -m playwright install chromium
```

### 初始化

```bash
# 仅用于新环境；不要在未备份的生产数据库上盲目执行
python 1.wooorders_sqlite.py
python inv_migrations.py status
python inv_migrations.py up
```

### 启动开发服务

```bash
python app.py
```

生产环境使用 Nginx + Gunicorn，不使用 Flask debug server。后台任务分别运行：

```bash
python fulfillment_worker.py --idle-sleep 5
python product_clone_worker.py --idle-sleep 2
python auto_sync.py
```

systemd 示例位于 `deploy/woo-fulfillment-worker.service` 和 `deploy/woo-product-clone-worker.service`。实际生产服务用户、可写目录和密钥读取权限应按主机环境收紧，示例中的路径不能直接视为通用安全配置。

## 配置与密钥

真实密钥不得提交到 Git。建议使用 `/etc/woo-analysis/*.env`，所有者为 `root:root`、权限 `0600`，并确保 Web 与相应 Worker 读取同一份配置。

| 变量 | 说明 |
|---|---|
| `OMS_DB_FILE`, `INV_DB_FILE` | 订单/履约数据库绝对路径 |
| `WMS_BASE_URL`, `WMS_SALT` | 匈牙利 WMS 地址与签名密钥 |
| `WMS_WEBHOOK_TOKEN` | WMS 回调独立令牌 |
| `WMS_ALLOW_INSECURE_HTTP` | 明文 HTTP 风险开关，默认必须为 `0` |
| `SZ56T_ORDER_BASE_URL`, `SZ56T_PRINT_BASE_URL` | 新波兰仓下单与打印服务地址 |
| `SZ56T_USERNAME`, `SZ56T_PASSWORD` | 新波兰仓 API 账户 |
| `SZ56T_CANCEL_AUTH` | 新波兰仓取消接口独立认证值 |
| `SZ56T_ALLOW_INSECURE_HTTP` | 新波兰仓明文 HTTP 风险开关，默认必须为 `0` |
| `ORDER_NOTIFICATION_EVENT_SECRET_REF` | 订单通知事件密钥的环境变量名 |
| `ORDER_NOTIFICATION_WEBHOOK_MASTER_KEY` | 数据库内 Webhook 的加密主密钥 |
| `ORDER_NOTIFICATION_IMAGE_DIR` | 私有订单图片目录，不得放在公开静态目录 |
| `ORDER_NOTIFICATION_CHROMIUM_PATH` | Chromium 可执行文件路径 |
| `MAIL_CENTER_ORDER_API_TOKEN_FILE` | 邮件中心 API 的只读令牌文件 |
| `MAIL_CENTER_ORDER_DB_PATH` | 邮件中心只读数据库路径 |
| `MAIL_CENTER_ORDER_ALLOWED_SITES` | 邮件中心允许访问的站点白名单 |

示例文件：

- `deploy/fulfillment.env.example`
- `deploy/order-notification.env.example`
- `deploy/validate_wms_env.sh`

## 测试

完整回归测试不访问真实 WMS，也不会创建真实发货单：

```bash
python -m pytest -q tests
```

发布前最低检查：

```bash
python -m py_compile app.py fulfillment_worker.py product_clone_worker.py
python inv_migrations.py status
git diff --check
```

外部 WMS 的只读校验必须明确使用审计模式，并区分“网络/API 可访问”与“真实创建订单成功”：

```bash
sudo deploy/validate_wms_env.sh
```

只有获得业务授权、确认收件数据和重复单处理方式后，才能进行一单受控创建联调。

## 生产发布与回滚

### 发布前

1. 确认 `git status --short`，任何生产临时修改都必须先归档或合并，不能覆盖。
2. 生成 SQLite 在线备份并执行 `PRAGMA integrity_check`。
3. 记录当前 Git commit、服务状态、Worker 状态和关键配置开关。
4. 在发布提交上执行完整测试和迁移状态检查。

备份脚本只接受受限目录格式：

```bash
sudo sh deploy/backup_live_db.sh \
  /www/backup/woo-analysis-pre-migration-YYYYMMDD-HHMMSS
```

### 发布

1. 停止会写数据库的 Web/同步/履约 Worker。
2. 部署已测试的固定 commit 或 tag，不直接部署未提交工作树。
3. 应用数据库迁移，重新检查 `inv_migrations.py status`。
4. 启动 Web 和 Worker，检查 systemd 日志。
5. 验证登录、订单列表、权限隔离、待发货、发货弹窗、物流查询和 WooCommerce 回读。

### 回滚

- 代码回滚到发布前固定 commit。
- 只有确认迁移的 `down` 不会破坏发布后的业务数据时才执行 schema 回滚。
- 更安全的数据库回滚方式是停写后恢复发布前已校验备份；恢复前保留故障数据库用于审计。
- 已提交到外部 WMS 或已通知客户的操作不能靠数据库回滚撤销，必须进入人工拦截、补偿和对账。

## 运维排查

### 页面提示返回 HTML、JSON 解析失败

`Unexpected token '<'` 通常表示接口收到了登录页、WAF/代理错误页或服务器 5xx HTML，而不是 JSON。检查浏览器 Network 响应、登录会话、Nginx/Gunicorn 日志和接口权限；不能把 HTTP 200 单独当作业务成功。

### 发货后仍显示“继续发货”

先检查本地 `shipping_logs` 是否为 `pending_sync`，再查询 WooCommerce 是否已保存运单。系统会优先远端回读，确认成功后再将本地状态改为 `shipped`；禁止重复使用同一运单盲目提交。

### 有库存但不能发货

检查三个层次：订单商品是否映射到标准 SKU、该 SKU 是否映射到当前仓、可用量是否扣除了预留。联合发货还需确认两个仓属于同一联合发货组。人工合作仓与有限库存中转仓的库存规则不同。

### 日志

```bash
journalctl -u woo-analysis -f
journalctl -u woo-fulfillment-worker -f
journalctl -u woo-product-clone-worker -f
```

日志和外部请求审计不得输出密码、API Key、Cookie、完整授权头或客户敏感信息。

## 进一步文档

- [MULTI_WAREHOUSE_FULFILLMENT.md](MULTI_WAREHOUSE_FULFILLMENT.md)：多仓履约、COD、状态和 WMS 规则。
- [INVENTORY.md](INVENTORY.md)：库存、SKU、批次、映射和对账。
- [ORDER_IMAGE_NOTIFICATIONS.md](ORDER_IMAGE_NOTIFICATIONS.md)：订单图片通知设计与运行方式。
- [OFFLINE_COMPANY_PROFIT_EXPORT.md](OFFLINE_COMPANY_PROFIT_EXPORT.md)：公司利润离线导出。
- [todo-new-poland-warehouse.md](todo-new-poland-warehouse.md)：新波兰仓上线前置条件。
- [todo-wms-go-live.md](todo-wms-go-live.md)：WMS 上线检查项。

## 已知边界

- 当前数据库为 SQLite，适合内部单实例运营；扩展到多租户、高并发 SaaS 前需重做租户隔离、密钥管理、并发写入和灾备设计。
- 主应用仍包含历史兼容代码，新增功能应优先放在独立服务模块，避免继续扩大 `app.py`。
- 外部 WMS 存在明文 HTTP 地址时，系统默认拒绝请求；不能为了联调绕过安全开关后长期遗留。
- WMS 适配器和单元测试通过不等于真实仓库已上线；生产可用必须有受控订单、面单、运单、仓库确认、Woo 回读和客户通知的端到端证据。
- 仓储费、运输费和 COD 手续费按供应商月结对账，不应混入客户应付金额或订单商品收入。

## 版本说明

### 2.0.1 — 2026-08-21

- 修复已完整发货订单因 WooCommerce 状态漂移而错误显示“继续发货”。
- 已退回、问题退货和已确认签收订单不再进入普通待发货队列。
- 保留真正的部分发货订单和 `pending_sync` 失败重试入口。

### 2.0.0 — 2026-08-21

- 合并服务器已上线代码与本地开发分支，建立可追溯的生产快照。
- 上线多仓履约、有限库存中转仓、联合发货和仓库行级权限。
- 完善 Packeta/Express One、DPD、InPost 查询与发货失败恢复。
- 增加仓库优先映射、批量确认、缺货重分仓和补货提醒。
- 改进 WooCommerce 订单分页、空远端结果保护、商品更新回读和新商品克隆。
- 增加订单通知、只读邮件中心接口及完整测试覆盖。

本项目为内部业务系统。许可、分发和对外部署范围由项目所有者另行确定。
