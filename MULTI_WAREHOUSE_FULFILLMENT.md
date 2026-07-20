# 多仓履约实施说明

## 业务规则

- 波兰、匈牙利、捷克站均可由波兰仓和匈牙利仓履约。
- SKU 仅在一个仓库存在时由该仓发；两个仓库均可发时，波兰站优先波兰仓，匈牙利站优先匈牙利仓，捷克站按已配置的有效运费最低值选择。
- 同一订单可拆为多个 `Fulfillment`，每个仓库可有独立 `Shipment` 和运单号。
- 缺货不自动换货或取消；订单进入红色缺货 + 人工处理状态。
- 匈牙利仓映射外部 `HU01`，渠道代码为完整字符串 `欧洲直发-25`。
- 单箱 10kg 仅为运营建议，不在系统中硬拦截。

## 聚合关系

`Order -> OrderItem -> Fulfillment -> FulfillmentItem -> Shipment -> ShipmentItem -> TrackingEvent`

- `Order` 是 WooCommerce 商业订单。
- `Fulfillment` 是一个仓库对该订单当前 revision 的履约责任。
- `Shipment` 是实际包裹；每个包裹独立同步 AST/WooCommerce 并通知客户。
- 所有非取消包裹都 `delivered` 后，订单聚合才进入 `delivered`，随后幂等同步 WooCommerce `completed`。

## 上线开关

迁移 006 默认写入：

- `oms_fulfillment_enabled=0`
- `oms_auto_plan_enabled=0`
- 匈牙利集成 `auto_submit=0`

因此迁移和服务部署本身不会自动分配历史订单，也不会生成真实 WMS 出库单。应在“多仓履约 → 配置”按以下顺序开启：

1. 完成站点商品到 `inv_skus` 的映射。
2. 配置 SKU 可用仓库及匈牙利 WMS 中文/英文品名、条码、图片。
3. 配置捷克站波兰仓/匈牙利仓有效运费。
4. 给仓库员工配置行级仓库权限。
5. 开启后台履约任务，手工规划少量测试订单。
6. 开启新同步订单自动分仓。
7. 确认 WMS 测试和审计日志后，再开启匈牙利 WMS 自动提交。

## 外部系统与数据方向

- WooCommerce → 本系统：订单、商品、地址、付款/COD、站点来源。
- 本系统库存 / WMS 库存 → 分仓器：SKU 可用量。
- 本系统 → 匈牙利 WMS：幂等出库单；WMS → 本系统：接单状态、拣货码、面单、动态运单和官方物流。
- 官方 WMS + InPost/Track718 → `TrackingEvent`：事件去重并防止乱序回退。
- 本系统 → AST/WooCommerce：每个包裹单独同步；部分发货不完成整单，最终妥投后才同步 `completed`。

## 失败恢复

- 所有外部操作通过 `oms_integration_jobs` 执行，包含唯一幂等键、租约、指数退避、最大重试和死信状态。
- WMS 创建请求超时视为“结果未知”，先按确定性 `invoiceCode` 查询，再决定是否重试创建。
- 外部请求写入脱敏的 `oms_external_api_calls`，状态变化写入不可变 `oms_domain_events`。
- 重复物流事件按包裹 + 指纹去重；迟到事件保留审计，但不能使包裹状态回退。
- 日常对账检查 WooCommerce 完成状态与本地聚合状态的差异，写入 `oms_reconciliation_issues`。

## 回滚

1. 先关闭三个上线开关并停止 `woo-fulfillment-worker.service`。
2. 代码回退到部署前分支/提交。
3. 如尚未向 WMS 提交任何真实出库单，可执行 `inv_migrations.py down 006` 删除 V2 表。
4. 如已提交真实出库单，只回退代码并保留审计表；外部 WMS 行为不能通过数据库回滚撤销，须按 WMS 取消接口和人工对账处理。
