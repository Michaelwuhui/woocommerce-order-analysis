# 香港订单图片群通知：只读核验、实施与验证报告

日期：2026-08-14
本地分支：`codex/order-image-notifications`
结论：代码、迁移 013 和超级管理员控制台已于 2026-08-13 部署生产；仅内置超级
管理员 `username='admin'` 可见和使用通知导航、页面、订单详情、图片、配置与重发
接口，其他管理员角色同样返回 403。控制台现可管理制卡规则、目标/路由并对真实订单
执行不入队、不发送的脱敏预览。三个功能/发送开关全部保持关闭，未配置目标、未创建
通知任务，也未向任何群发送测试或正式消息。

> 2026-08-14 FluentSMTP 增量已写入香港生产主机：管理员新订单邮件原稿读取、安全净化、
> 同站图片内联和 Chromium PNG 渲染均已接入超级管理员预览及影子 Worker。生产安装
> BeautifulSoup 4.14.2、Playwright 1.58.0 和独立 Chromium 145.0.7632.6，Web/Worker
> 重启后均为 active。用既有订单 `3-11510` 只读命中 FluentSMTP 日志 885 并生成一页
> 720×1310 PNG；浏览器渲染网络请求为 0，临时目录自动清理。三个功能/发送开关仍为 0，
> 目标、事件、任务和发送尝试仍为 0；本次部署不构成正式群发送授权。

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
- 超级管理员控制台：四步使用指引、开关/队列状态、真实订单安全预览、状态与保留规则、
  目标/路由管理和历史任务；测试发送与正式发送字段在 API 层拒绝修改。
- 安全预览直接读取权威快照并返回临时脱敏 PNG，不创建 `order_notification_jobs` 或
  `oms_integration_jobs`，不调用 Provider；临时文件即时清理，仅写 `preview_generated`
  审计（事件类型与页数，不保存客户明细）。
- FluentSMTP 原稿模式只读复用已有订单邮件日志接口；严格区分管理员新订单邮件与客户
  收件邮件，精确匹配订单号和成功状态。HTML 删除主动内容/外链，可信商品图服务端无
  凭据内联，Chromium 全网络阻断；预览审计不保存正文或客户资料。
- 新订单影子任务可选择邮件原稿；订单变更、取消、暂停继续生成实时状态卡，避免拿旧
  邮件冒充最新状态。找不到合法原稿时显式重试/失败，不静默回退。
- 企业微信目标仅向页面返回环境变量名及“是否已注入”，绝不返回环境变量值；启用的
  完全相同路由会被拒绝，避免运行期最高优先级歧义。

企业微信接口约束已于 2026-08-13 从[官方群机器人配置说明](https://developer.work.weixin.qq.com/document/path/91770)刷新：图片为 JPG/PNG、原图小于 2 MB、正文使用 Base64 与原图 MD5；官方页面标示频率 20 条/分钟。本实现再保守限制为每目标最多 15 条/分钟。

## 3. 验证证据

- 通知专项：新增邮件链路、Chromium 加固环境及智能状态筛选后 30 passed。
- 正式测试目录：126 passed，8 subtests passed。
- Python 语法编译：通过。
- Flask 路由/蓝图及 3 个模板加载烟雾测试：通过（6 routes、3 templates、blueprint registered）。
- `git diff --check`：无空白错误；仅 Windows 工作区既有 LF/CRLF 提示。
- 企业微信 Provider 使用模拟 Session 验证，未访问真实 Webhook；Fake Provider 不发网络请求。
- 可视化样本位于 `C:\Users\Administrator\Documents\香港服务器加固\outputs\order-notification-validation-20260813`，全部使用虚构订单/客户数据。
- 部署后生产核验：Web/Worker 均 active，迁移 013 记录数为 1、五张通知表存在数为 5；
  目标、事件、通知任务、尝试、审计、`ORDER_NOTIFICATION` 队列任务均为 0，三个开关
  均为 0，卡片目录为空，`PRAGMA quick_check=ok`。
- Chrome 真实登录态验证：超级管理员可见完整控制台；最近订单加载 50 条，对真实订单
  `3-11510` 成功生成 1 页脱敏预览，页面明确显示“未入队 · 未发送”。预览后任务、
  队列、发送尝试和卡片目录仍为 0，仅新增 1 条最小化预览审计；桌面与移动布局无横向
  溢出。匿名访问返回 302，普通管理员 403 边界由自动化测试验证。
- 2026-08-14 生产邮件原稿验收：生产 Chromium 动态库及启动烟雾测试通过；既有真实订单
  `3-11510` 只读命中 FluentSMTP 管理员新订单日志 885，内联 3 张同站商品图，生成一页
  720×1310、113,268 bytes PNG，PNG SHA-256 为
  `e68b6804e688f08b86cc94a671564285b98815b9b6d4e611658b541b3d5a97ab`；浏览器阻断请求数
  为 0，服务器临时渲染目录退出后为 0。超级管理员生产页面已显示“真实邮件原稿
  （FluentSMTP，推荐）”默认选项、“不入队 · 不发送”和“正式发送已锁定”。
- 2026-08-14 订单选择增量验收：新增超级管理员只读订单查询接口，可按站点加载最近
  50 单，并按普通订单号、带 `#` 订单号或内部订单 ID 搜索；响应只含订单 ID、订单号、
  状态、站点、仓库 ID 和修改时间，不含客户资料。Chrome 生产登录态选择 `vapego.pl`
  后 50 个结果全部属于该站点，虚构订单号搜索返回 0 且清空旧内部 ID；未生成预览、
  未写审计、未创建队列任务。
- 2026-08-14 FluentSMTP 漏数修复验收：Strefa 后台确认订单 `21193` 有客户和管理员
  两封已发送邮件，但旧版 `woo-tracking` Logger 路径返回 0 且提前禁用直接表查询。
  MU 插件 1.2.1 改为同时探测 `wp_fsmpt_email_logs` 并按日志 ID 去重。修复后接口返回
  两封唯一日志，香港系统排除客户邮件并选中管理员日志 `9800`；完整 Chromium 渲染
  生成一页 720×1253、102,167 bytes PNG，内联 2 张同站图片，浏览器网络请求为 0，
  临时目录退出后为空。
- 2026-08-14 浏览器 502 修复验收：Cloudflare Ray `a2acf7268d54b190` 的响应正文证明
  502 是源站错误替换页；本机直连 Gunicorn 得到实际错误 `email_html_render_failed`。
  同等 systemd 加固单元复现到 Chromium Crashpad 因 `ProtectHome=true` 无可写数据库目录
  而退出。渲染器现为每次任务建立私有临时 HOME/XDG 目录并在退出时清理。Chrome 生产
  登录态再次预览 `2-21193` 返回 HTTP 200 / JSON，命中日志 `9800`，显示 1 页、
  “未入队 · 未发送”；三个开关、目标、事件、任务、尝试和卡片文件仍为 0。
- 2026-08-14 智能新订单筛选验收：生产近 90 天聚合显示澳洲 0 笔 COD、182 笔
  非 COD，波兰 7,405 笔 COD。筛选依据订单列表既有支付语义：COD `processing=待发货`、
  `pending=待处理`；线上/BACS `processing=已支付或已确认·待发货`；所有 `on-hold`
  均排除，因为 COD/线上语义为已发货，BACS 语义为待转账确认。同时提供全部状态和具体
  状态筛选。Chrome 登录态验证 `vapicoau.com` 智能新单仅返回
  `#11510 processing / 非 COD`；`vapesklep.pl` 智能新单仅返回 COD processing，
  不再混入 COD on-hold；具体 `completed` 筛选的 50 条结果全部为 completed。验收只调用 GET，审计仍为 6，
  三个开关、目标、事件、任务、尝试及卡片文件仍为 0。
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
- 控制台增量部署备份：`/www/backup/woo-analysis-notification-console-20260813-180206`，
  含更新前 API、页面模板和 SQLite 在线备份；生产只替换这两个文件并只重启 Web，
  Worker 未改动。部署后服务 active、`PRAGMA quick_check=ok`、Web 日志无应用异常，
  远端暂存目录已清理。
- FluentSMTP 邮件原稿增量备份：
  `/www/backup/woo-analysis-email-render-20260814-103340`（约 123 MB），含部署前代码、
  systemd 单元与 drop-in、pip 冻结清单、文件校验值及 SQLite 在线备份；备份数据库
  `PRAGMA quick_check=ok`。
- 生产仅逐文件替换邮件渲染器、通知 API/Service、模板、依赖清单和迁移定义；安装
  BeautifulSoup 4.14.2、Playwright 1.58.0、Chromium 145.0.7632.6 及其系统动态库。
  Chromium 固定目录为 `/var/lib/woo-analysis/playwright`，独立 0600 环境文件仅保存
  Chromium 与卡片目录路径，没有写入群机器人密钥。
- Worker 于 2026-08-14 10:38:44 HKT、Web 于 10:38:48 HKT 重启；Gunicorn 4 workers、
  登录页 HTTP 200、匿名通知页 302，重启后日志无新 Traceback/CRITICAL。数据库
  `PRAGMA quick_check=ok`，三个开关为 0，目标/事件/任务/尝试均为 0，卡片目录为空。
- 站点筛选/订单搜索增量备份：
  `/www/backup/woo-analysis-notification-filter-20260814-110727`（约 123 MB），含更新前
  API、页面模板、systemd 单元、SHA-256 清单和 SQLite 在线备份，备份数据库
  `PRAGMA quick_check=ok`。生产仅替换 `order_notification_api.py` 与
  `templates/order_notifications.html`，Web 于 11:09:21 HKT 重启；Worker PID 保持
  `3702183` 未重启。登录页 HTTP 200，匿名控制台及新查询接口均为 302，生产登录态
  联动查询通过。
- Strefa FluentSMTP 拉取修复备份：
  `/root/.codex/backups/strefa-email-log-fallback-20260813-233322`，含更新前 MU 插件及
  SHA-256/文件状态/PHP 版本证据，目录权限 `0700`。生产只将
  `woo-orders-tracking-rest-api.php` 从 1.2.0 更新到 1.2.1，文件保持 `0644 www:www`；
  PHP 8.3.27 语法检查、站点主页 HTTP 200、匿名 REST 403 及香港凭据端到端读取通过，
  未重发邮件、未修改订单或通知发送开关。
- Chromium 加固环境修复备份：
  `/www/backup/woo-analysis-preview-chromium-home-20260814-115451`，含替换前邮件渲染器
  与控制台模板。生产逐文件替换后 Web/Worker 于 11:55 HKT 重启并保持 active；生产
  SHA-256 分别为 `cdb6202d8d14c7485e9f0bb2437bfdfe1fdcac2a0c75361e19274ebee2ef7a08`
  和 `07b467c6fe2cecd152843f79a782a700848aab7521a08b34c1da1c478341eabb`。
- 智能订单状态筛选备份：
  `/www/backup/woo-analysis-notification-smart-status-20260814-120725`，含替换前通知 API
  与控制台模板；支付语义收紧前的中间版本另备份于
  `/www/backup/woo-analysis-notification-smart-status-final-20260814-121144`。生产仅替换
  API 与模板并重启 Web，Worker 未重启且始终 active；最终生产 SHA-256 分别为
  `cd636a6f8bc382bcd4e2b2c0acaa7abc186985a4fb8c1a8969b5a54ecc8927c6` 和
  `633029d6391f53b286395d8c8c9fc8575d02a3178311917921e1e761e47014bf`。
- 隔离测试群一次性发送入口备份：
  `/www/backup/woo-analysis-notification-test-send-20260814-131332`，权限 `0700`
  且含替换前 API、Service、控制台模板、SQLite 在线备份及 SHA-256 清单；备份数据库
  `PRAGMA quick_check=ok`。生产仅替换上述三个文件并重启 Web/Worker；最终 SHA-256
  分别为 `64af4c1386534b703698c6827af5729351b35423bf9a6c135306d911e1a8b050`、
  `ab2ff3a99339a4c970b3f259340fef634beb747e892537dfd265952b162a00a6`、
  `bf56857124a6f7809cfa8c81f461bdf4de6bf13482109a4de88545896bc9abc1`。
  本地全量回归为 `129 passed, 8 subtests passed`。生产登录页 HTTP 200、匿名通知页
  302、两项服务 active、数据库 `quick_check=ok`；三个开关仍为 0，目标/事件/任务/尝试
  仍为空，环境文件没有企业微信 Webhook 变量，因此部署过程零发送。

## 6. 后续启用顺序（另行授权后）

严格按 `ORDER_IMAGE_NOTIFICATIONS.md`：备份与数据库 quick check → 在测试环境安装依赖/CJK 字体检查 → 迁移 013 → 仓库外 0600 环境文件 → Fake 目标 → 隔离企业微信测试目标 → 脱敏测试订单 → 影子预览 → 试点店铺/仓库 → 观察告警与死信 → 单独批准正式发送开关。

回滚时先关闭三个开关，保留任务/尝试/审计，不补发旧影子任务；迁移只在五张通知表完全无数据时允许 down。

## 7. 正式群授权与全站启用（2026-08-14）

- 用户在当前会话明确授权使用现有企业微信内部群自动发送订单图片；机器人采用企业
  微信官方群 Webhook，不保留 Hermes 长连接，不使用个人微信 Hook、外挂或模拟点击。
  Webhook 只存在于服务器 `0600` 环境文件和数据库环境变量引用中，没有写入代码、
  文档、测试产物或 Git 历史。
- 先以 `vapesklep.pl` 灰度验证新订单图片、COD/在线支付新订单状态、FluentSMTP 管理员
  邮件选择、失败重试和群内文字告警；随后按用户授权扩展到其余可用站点。
- 订单系统登记的 30 个站点中，`buchmistrz.pl` 因未安装 FluentSMTP 被明确排除；其余
  29 个站点对应的 27 个有效 WordPress 安装均已使用只读接口 1.2.1，并完成 PHP 语法、
  真实 REST 路由及凭据读取验证。共享多站点安装只部署一份 MU 插件。
- 生产目标 `wecom_vapesklep_production` 已改为唯一的全站生产路由
  `全部站点订单群（生产）`，仍保持单目标限流 15 条/分钟；测试目标和生产目标继续按
  `environment` 隔离，测试发送开关保持关闭。
- 2026-08-14 17:25:29（北京时间）事务性写入 29 个逐站水位。水位取启用瞬间
  WooCommerce 只读 API（`auvapeshub.com` 使用源站 WordPress 只读引导查询）与本地库
  最大订单号的较大值，因此仅发送水位之后同步到的新订单，不补发历史订单。
- 启用前在线备份为
  `/www/backup/woo-analysis/woocommerce_orders.pre-all-sites-notification-20260814-092529Z.db`，
  权限 `0600`，`PRAGMA quick_check=ok`，SHA-256 为
  `1730414525c5a559a48a58cf82022830c9e4473e955a3fa7ef942613a13264b8`。
- 启用审计记录为 `notification_audit_logs.id=66`，未包含 Webhook；提交后活动通知队列、
  历史回灌任务和水位违规均为 0。Web 与 Worker 保持 active，登录页 HTTP 200，启用后
  服务日志没有新错误。
- 最终正式测试命令 `python -m pytest -q tests` 通过：`137 passed, 8 subtests passed`。
