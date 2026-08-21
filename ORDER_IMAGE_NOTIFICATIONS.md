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

## 配置顺序（只在获批的测试环境执行）

1. 备份代码和 SQLite，运行 `PRAGMA quick_check`。
2. 安装锁定依赖并准备系统 CJK 字体。香港主机部署使用
   `/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc`；Droid fallback 无法完整覆盖
   中文、数字和波兰字符的混排。若主机字体变化，使用
   `ORDER_NOTIFICATION_FONT` 指向审核过的 CJK 字体。
3. 执行 `python inv_migrations.py up 013`。
4. 将 `deploy/order-notification.env.example` 复制到仓库外的
   `/etc/woo-analysis/order-notification.env`，权限设为 `0600`。
5. 建立一个 `environment='test'` 的 `FAKE` 目标，先运行 Fake Provider 验收。
6. 由企业微信管理员建立内部测试群及测试机器人；将 Webhook 仅写入
   `WECOM_TEST_ORDERS_WEBHOOK`，数据库目标只写
   `secret_ref='env:WECOM_TEST_ORDERS_WEBHOOK'`。
7. 将站点规则写入 `settings.order_notification_policy_json`，确认 COD、在线付款、
   BACS、测试单、失败单和暂停单规则。
8. 仅在测试数据副本打开 `order_notification_enabled=1`，再打开
   `order_notification_test_send_enabled=1`；生产发送开关继续为 0。
9. 重启 Web 与 Worker，使用不含真实客户资料的测试订单验收。

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

未取得当前会话正式授权，不执行生产迁移、服务重启、目标写入或开关变更。

## 回滚

1. 首先将三个通知开关全部设为 `0`，Worker 不再领取新的通知发送任务。
2. 保留任务、快照和尝试审计；待处理任务标记/保留为 HOLD，不批量补发旧单。
3. 回退代码并重启 Web/Worker，恢复人工截图流程。
4. 仅当迁移 013 五张业务表全部为空时，才允许 `python inv_migrations.py down 013`；
   迁移会在存在任何历史记录时拒绝删表。
5. Webhook 若泄漏，在企业微信删除/重建机器人，轮换仓库外的环境变量引用。
