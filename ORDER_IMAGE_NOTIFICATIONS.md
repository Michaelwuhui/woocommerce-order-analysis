# 订单图片群通知部署与回滚

## 当前状态

代码完成后默认保持暗启动：迁移 013 写入
`order_notification_enabled=0`、`order_notification_test_send_enabled=0` 和
`order_notification_send_enabled=0`。因此，仅部署代码或执行迁移不会创建任务，
也不会向任何企业微信群发送消息。

发送开关关闭时，Worker 仍可在获批的影子环境将任务终结为 `READY_PREVIEW`；
以后打开开关不会补发这些旧卡。确需发送时必须通过有审计记录的人工重发创建新任务。

普通微信群只支持 `MANUAL_WECHAT`：系统生成私有 PNG，员工登录后下载并人工转发。
不得加入微信 Hook、外挂、群控、按键脚本或模拟点击。

首期生产可见性严格绑定内置超级管理员 `username='admin'`：导航、汇总页、订单通知
详情、目标配置、图片下载和人工重发接口均拒绝其他管理员角色及普通用户。

## 超级管理员控制台使用

打开顶部导航“群通知”后，按以下顺序操作：

1. **真实订单安全预览**：从最近订单选择一笔。默认来源“真实邮件原稿”会只读查询该
   订单的 FluentSMTP 日志，严格选择成功发送给管理员的 WooCommerce 新订单邮件，
   使用原始 HTML 生成 PNG。它包含邮件原有的商品图、订单金额、付款/配送方式和地址，
   因而只允许超级管理员查看。系统不会调用 WooCommerce 的邮件 action，也不会重发
   邮件。若要核对变更/取消/暂停状态，可切换为“系统脱敏卡片（备用）”。两种预览都
   不创建通知任务、不进入队列、不访问企业微信；临时文件在响应后删除。
2. **通知目标与路由**：先使用 `MANUAL_WECHAT`（人工下载转发）或 `FAKE`（不联网）
   完成闭环。路由可按站点、仓库、配送方式逐层收窄，系统拒绝同范围的第二个启用目标。
3. **企业微信目标**：页面只填写服务器环境变量名，例如
   `TEST_WECOM_WEBHOOK`；Webhook 本身必须保存在仓库外 `0600` 环境文件中，页面、
   API 和数据库均不返回密钥值。测试变量名必须含 `TEST`，正式目标不能引用测试变量。
4. **制卡规则**：可编辑状态分组、防抖时间、图片保留期、变更字段和“新订单自动制图
   来源”。邮件原稿只用于 `ORDER_READY`；订单变更、取消和暂停始终使用醒目状态卡，
   避免复用已经过时的新订单邮件。开启“自动生成通知图片”只会启用制图总开关；页面
   不能修改测试发送或正式发送开关。
5. **任务记录**：查看任务、路由、错误和已有卡片。接口接受只代表平台接受，不代表群
   成员已读；影子预览不会自动补发。
6. **隔离测试群单次发送**：生成预览后，页面仅列出与该订单路由匹配、密钥已注入的
   `environment='test'` 企业微信群目标。当前超级管理员必须在 15 分钟内二次确认，
   系统才会创建一条有审计记录的队列任务；同一个预览令牌重复提交不会重复入队。
   此入口不打开自动制卡，因此不会顺带处理其他订单。Worker 发送前会再次确认目标仍为
   测试环境、路由仍匹配且测试发送开关仍开启；任一条件变化即进入死信且不访问企业微信。

正式发送开关由服务端锁定。没有当前会话明确授权，即使超级管理员也不能通过控制台
开启测试群或正式群发送。

## 数据与队列

- `notification_targets`：店铺/仓库/配送路由，仅保存 `secret_ref`，不保存 Webhook。
- `order_notification_event_inbox`：HMAC 事件和防重放记录。
- `order_notification_jobs`：权威订单快照、卡片、状态、幂等和错误。
- `order_notification_attempts`：逐次/逐页发送尝试。
- `notification_audit_logs`：人工重发及配置操作审计。
- 复用 `oms_integration_jobs` 和 `woo-fulfillment-worker.service`；每个通知任务的
  `idempotency_key` 唯一，Worker 使用现有租约、重启恢复、退避和死信。

订单版本不是 `orders.version`（该列实际是 WooCommerce 软件版本），而是
`date_modified + 权威快照 SHA-256`。

## FluentSMTP 原稿安全边界

- 邮件列表与正文只通过站点既有 `woo-tracking` 只读 GET 接口获取；Woo consumer key
  和 secret 仅放在请求头，不进入 URL、浏览器或审计日志。
- 必须同时满足：发送成功、主题含精确订单号、主题属于多语言“管理员新订单”类型、
  收件人不是结账邮箱；找不到时明确失败并重试，不拿客户收件邮件或其他订单替代。
- 渲染前删除脚本、表单、iframe、事件属性、跳转链接和 CSS 外链。仅允许同站 HTTPS
  商品图或明确配置的图片主机；图片由服务端无凭据下载、校验后转为 data URI。
- Chromium 收到的页面已加 CSP，并在网络层拦截所有请求，因此不会触发邮件追踪像素。
- 审计只保存邮件日志 ID、HTML 哈希、图片数量和模板版本，不保存邮件正文、地址或电话。

## 配置顺序（只在获批的测试环境执行）

1. 备份代码和 SQLite，运行 `PRAGMA quick_check`。
2. 安装锁定依赖并准备系统 CJK 字体。香港主机部署使用
   `/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc`；Droid fallback 无法完整覆盖
   中文、数字和波兰字符的混排。若主机字体变化，使用
   `ORDER_NOTIFICATION_FONT` 指向审核过的 CJK 字体。
   邮件原稿渲染还需要 Python `beautifulsoup4`、`playwright` 和经核验的系统 Chromium；
   用 `ORDER_NOTIFICATION_CHROMIUM_PATH` 显式指定可执行文件。若商品图使用独立 CDN，
   仅把审核过的域名写入 `ORDER_NOTIFICATION_IMAGE_HOST_ALLOWLIST`（逗号分隔）。
3. 执行 `python inv_migrations.py up 013`。
4. 将 `deploy/order-notification.env.example` 复制到仓库外的
   `/etc/woo-analysis/order-notification.env`，权限设为 `0600`。
   Web 服务使用 `deploy/woo-analysis.service.d/order-notification.conf` 作为 systemd
   drop-in，Worker 已在自身 unit 中读取同一环境文件；变更后先 `systemctl daemon-reload`。
5. 建立一个 `environment='test'` 的 `FAKE` 目标，先运行 Fake Provider 验收。
6. 由企业微信管理员建立内部测试群及测试机器人；将 Webhook 仅写入
   `WECOM_TEST_ORDERS_WEBHOOK`，数据库目标只写
   `secret_ref='env:WECOM_TEST_ORDERS_WEBHOOK'`。
7. 将站点规则写入 `settings.order_notification_policy_json`，确认 COD、在线付款、
   BACS、测试单、失败单和暂停单规则。
8. 仅在测试数据副本打开 `order_notification_enabled=1`，再打开
   `order_notification_test_send_enabled=1`；生产发送开关继续为 0。
9. 重启 Web 与 Worker，使用不含真实客户资料的测试订单验收。

若使用控制台的“发送当前预览到隔离测试群”，不需要打开
`order_notification_enabled`；只在获批的短时测试窗口打开
`order_notification_test_send_enabled`。该入口仍复用正式队列和 Worker，但只创建当前
确认预览对应的一条测试任务。真实邮件原稿包含客户资料，必须先确认测试群成员边界。

目标路由优先级由已填写字段决定：店铺+仓库+配送、店铺+仓库、店铺、全局。
最高优先级出现两个目标时不会猜测，任务进入路由异常/死信。

## 启动前验证

```bash
./venv/bin/python -m pytest tests -q
./venv/bin/python inv_migrations.py status
sqlite3 -readonly woocommerce_orders.db 'PRAGMA quick_check;'
systemctl show woo-analysis.service woo-fulfillment-worker.service -p ActiveState -p SubState
```

必须额外验证：重复事件只对应一个任务；Worker 中止/重启后任务恢复；限流时保留
任务；变更/取消卡醒目；失效测试 Webhook 进入死信；日志、页面和 Git 无密钥；
测试配置不能指向生产目标；手机和电脑端图片均可读。

## 生产上线前仍需批准

- 当前接收群是否为企业微信内部群；
- 企业微信测试群/机器人及独立的密钥引用；
- 一个店铺、一个仓库的状态规则与路由；
- 卡片允许展示的姓名/电话/地址边界及 30 天图片保留期；
- 影子模式结果和正式 Feature Flag 批准人。

未取得当前会话正式发送授权，不写入真实群目标、不注入 Webhook、不启用测试或正式
发送开关。部署/维护控制台本身不构成发送授权。

> 生产状态更新（2026-08-14）：当前会话后来已取得用户对现有企业微信内部群的明确
> 正式发送授权。生产启用范围、逐站水位、排除站点和回滚证据记录在
> `ORDER_IMAGE_NOTIFICATIONS_IMPLEMENTATION_REPORT.md`。上述授权门槛仍适用于以后新增
> 群、机器人、站点或发送范围的变更，不得由一次部署或历史授权自动外推。

## 回滚

1. 首先将三个通知开关全部设为 `0`，Worker 不再领取新的通知发送任务。
2. 保留任务、快照和尝试审计；待处理任务标记/保留为 HOLD，不批量补发旧单。
3. 回退代码并重启 Web/Worker，恢复人工截图流程。
4. 仅当迁移 013 五张业务表全部为空时，才允许 `python inv_migrations.py down 013`；
   迁移会在存在任何历史记录时拒绝删表。
5. Webhook 若泄漏，在企业微信删除/重建机器人，轮换仓库外的环境变量引用。
