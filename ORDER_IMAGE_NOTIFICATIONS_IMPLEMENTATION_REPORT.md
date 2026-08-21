# 香港订单图片群通知：只读核验、实施与验证报告

日期：2026-08-13  
本地分支：`codex/order-image-notifications`  
结论：代码、迁移 013 和超级管理员只读入口已于 2026-08-13 部署生产；仅内置超级
管理员 `username='admin'` 可见和使用通知导航、页面、订单详情、图片、配置与重发
接口，其他管理员角色同样返回 403。三个功能/发送开关全部保持关闭，未配置目标、
未创建通知任务，也未向任何群发送测试或正式消息。

## 1. 只读核验结论

### 真实运行栈

- 主机别名：`cangfu_hk`；应用目录：`/www/wwwroot/woo-analysis`。
- `woo-analysis.service`：Gunicorn 4 workers，监听 `127.0.0.1:5000`，核验时 active。
- `woo-fulfillment-worker.service`：现有 Python 持久队列 Worker，核验时 active。
- Python 3.12.3；Flask 3.1.2；Flask-Login 0.6.3；Werkzeug 3.1.4；Gunicorn 23.0；Requests 2.32.5；WooCommerce SDK 3.0。
- SQLite 3.45.1；数据库 `/www/wwwroot/woo-analysis/woocommerce_orders.db`，约 126 MB；`PRAGMA quick_check=ok`。
- 部署前已应用库存/履约迁移 001-012；本次应用迁移 013，并安装 Pillow 12.3.0。
- 初次只读检查仅发现 Droid fallback；Linux 视觉验收证明它不能完整覆盖中文、数字和
  波兰字符混排，因此生产部署改用 Noto CJK，并在渲染器加入逐字符字体回退。

### 订单状态与业务事实

只读统计共 11,095 单：completed 7,780；cancelled 1,835；on-hold 609；shipped 372；failed 286；processing 170；delivered 18；refunded 11；checkout-draft 10；partial-shipped 2；cheat 1；pending 1。

最近 30 天订单以 COD 为主；不同站点的 `cod_on_hold_is_shipped` 配置并不相同。因此通知规则不能把所有 `on-hold` 一概视为可履约，实施中保留了站点级 COD 判断和可配置状态策略。

订单主键是跨站点安全的 `<site_id>-<woo_id>`。`orders.version` 的真实值是 WooCommerce 软件版本（例如 10.4.3/11.0.1），不能作为订单乐观版本；通知版本改用 `date_modified + 权威快照 SHA-256`。

### 队列与测试环境

- 现有 `oms_integration_jobs` 已有唯一幂等键、租约、过期租约恢复、指数退避、最大尝试次数和死信；只读核验时 964 个历史任务均为 succeeded。
- 新通知复用该队列和现有 Worker，不新增第二套易漂移队列。
- 未发现独立 staging/pre-production 服务，也未发现现有企业微信通知实现或已配置测试群机器人。
- 生产仓库当前有既有未提交改动/未跟踪文件，不能以覆盖目录方式部署。

## 2. 已完成的实现

- 迁移 013：目标路由、事件收件箱、通知任务、逐次发送尝试、审计日志；五张表有幂等约束和历史保护。
- 三层暗启动开关全部默认关闭：任务总开关、测试发送开关、生产发送开关。
- HMAC-SHA256 事件入口：5 分钟时间窗、64 KiB 上限、事件 ID 防重放、冲突检测、来源/字段校验；事件只用于触发，发送前重新读取本地权威订单。
- 事件：`ORDER_READY`、`ORDER_UPDATED`、`ORDER_CANCELLED`、`ORDER_HOLD`、`MANUAL_RESEND`。
- 45 秒防抖；结束时重新读取订单并重新计算变更字段。状态已变化或变更已回退时安全跳过。
- 店铺/仓库/配送方式路由；最高优先级并列时不猜测，进入路由异常/死信。
- 同订单事件按现有队列 ID 串行；任务有唯一幂等键、重发序号和审计操作人。
- 1080 px PNG 卡片：新单/变更/取消/暂停/人工重发使用不同醒目色；中波英混排；长单每 7 项分页；每张小于 2 MB；保留商品名、规格、SKU 和数量。
- 收件信息最小化：姓名和电话遮罩，只保留城市、邮编、配送点；不进入完整街道地址。
- 官方企业微信图片适配器：HTTPS 固定官方主机/路径/端口，禁重定向，原图 Base64 + MD5，连接/读取超时区分，429/5xx 重试，业务错误死信。
- 普通微信群仅提供登录后私有图片下载 (`MANUAL_WECHAT`)，无 Hook、外挂、群控或模拟点击。
- 多页部分成功后不重发已接受页；原始图片字节冻结。若部分成功后订单再变化，进入人工复核，禁止拼接两个版本。
- 30 天默认图片保留期；清理只允许删除配置私有目录下的终态卡片，订单任务和审计历史保留。
- 订单详情与汇总页显示任务状态、目标、模板、尝试、错误、预览以及有审计的人工重发；沿用登录、管理员和站点范围权限。

企业微信接口约束已于 2026-08-13 从[官方群机器人配置说明](https://developer.work.weixin.qq.com/document/path/91770)刷新：图片为 JPG/PNG、原图小于 2 MB、正文使用 Base64 与原图 MD5；官方页面标示频率 20 条/分钟。本实现再保守限制为每目标最多 15 条/分钟。

## 3. 验证证据

- 通知专项：19 passed。
- 正式测试目录：114 passed，8 subtests passed。
- Python 语法编译：通过。
- Flask 路由/蓝图及 3 个模板加载烟雾测试：通过（6 routes、3 templates、blueprint registered）。
- `git diff --check`：无空白错误；仅 Windows 工作区既有 LF/CRLF 提示。
- 企业微信 Provider 使用模拟 Session 验证，未访问真实 Webhook；Fake Provider 不发网络请求。
- 可视化样本位于 `C:\Users\Administrator\Documents\香港服务器加固\outputs\order-notification-validation-20260813`，全部使用虚构订单/客户数据。
- 部署后生产核验：Web/Worker 均 active，迁移 013 记录数为 1、五张通知表存在数为 5；
  目标、事件、通知任务、尝试、审计、`ORDER_NOTIFICATION` 队列任务均为 0，三个开关
  均为 0，卡片目录为空，`PRAGMA quick_check=ok`。
- Chrome 真实登录态验证：超级管理员刷新后可见“群通知”，页面返回 200 并显示
  “暂无通知任务”；匿名访问返回 302。普通管理员 403 边界已由本地与 Linux 暂存
  两套自动化测试验证。
- 根目录执行 `pytest` 会收集仓库历史诊断脚本，触发既有 `orders` 表缺失/SQLite 锁冲突；项目正式测试入口是 `python -m pytest tests -q`，本次未把历史脚本问题误记为新回归。

## 4. 正式群发送前门禁

当前仅完成代码/数据库暗启动。以下任一项未完成都不得打开任务或发送开关：

1. 企业微信管理员确认接收渠道是企业微信内部群，并建立与正式群隔离的测试群/机器人。
2. 明确一个试点店铺、一个仓库的状态规则、路由、字段展示边界和批准人。
3. 在脱敏数据上先跑 Fake，再跑隔离测试群；正式发送开关始终为 0。
4. 后续仍须采用逐文件发布和备份/回滚，禁止覆盖生产脏工作树。
5. 旋转并移出源码中的既有密钥：只读扫描发现两个已跟踪旧脚本含 WooCommerce consumer key/secret；`app.py` 还存在硬编码 Flask session secret 和默认管理员口令。报告不记录任何密钥值。完成轮换、外置和登录回归前，不应宣称达到生产安全门槛。
6. 由批准人在测试结果上签字后，另行给出目标写入和 Feature Flag 变更授权。

## 5. 本次生产部署记录

- 可恢复备份：`/www/backup/woo-analysis-order-notification-20260813-135618`，含代码、
  SQLite 在线备份、发布前 pip/dpkg 清单和校验值；数据库自动迁移备份为
  `/www/wwwroot/woo-analysis/woocommerce_orders.db.preup013_20260813_140834`。
- 生产 `app.py` 与 `templates/base.html` 存在既有未提交改动，本次从生产现文件合并
  狭窄通知差异，没有以本地整文件覆盖。
- 安装 Pillow 12.3.0、Ubuntu `fonts-noto-cjk`，Linux 虚构订单卡片经目视验收；未安装
  pytest 到生产虚拟环境。
- Worker 于 14:09:17 HKT 重启，Web 于 14:09:31 HKT 重启；Gunicorn 4 workers、
  本机 5000 端口及登录页恢复正常。
- 发送相关数据与目录均为空，三个开关为 0；部署不构成正式群发送授权。

## 6. 后续启用顺序（另行授权后）

严格按 `ORDER_IMAGE_NOTIFICATIONS.md`：备份与数据库 quick check → 在测试环境安装依赖/CJK 字体检查 → 迁移 013 → 仓库外 0600 环境文件 → Fake 目标 → 隔离企业微信测试目标 → 脱敏测试订单 → 影子预览 → 试点店铺/仓库 → 观察告警与死信 → 单独批准正式发送开关。

回滚时先关闭三个开关，保留任务/尝试/审计，不补发旧影子任务；迁移只在五张通知表完全无数据时允许 down。
