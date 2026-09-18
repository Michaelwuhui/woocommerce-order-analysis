# 2026-09-18 代码统一记录

GitHub `main` 是发布基线。日常开发从最新 `origin/main` 建立 `codex/*` 分支，完成测试后经 PR 合并；生产目录只快进到已验证的合并提交。发布后核对本地、GitHub、生产的提交、代码树、工作区状态及运行服务。

## 本次归并

- [PR #31](https://github.com/Michaelwuhui/woocommerce-order-analysis/pull/31) 修复“作为全新商品克隆”仍被原 SKU 拦截的问题。保留按任务生成的新 SKU 检查，使同一任务重试可以复用已创建商品。
- 补回商品克隆 SKU、创建/跳过计数、邮件中心只读接口、离线利润计算应用上下文、自动确认、库存补货重算、履约终态恢复、发货重试和站点 API 检测测试。
- PostgreSQL 发货重试测试自行准备合成站点及不同权限的用户，不再依赖某次临时测试库的预置数据。
- 站点同步权限测试改用隔离 PostgreSQL 和当前持久化任务状态；模拟任务发布。清理同步在 PostgreSQL Web 端已停用，测试保留该保护。
- 邮件中心测试同时验证 SQLite 与 PostgreSQL；令牌文件权限测试在 Linux 执行。示例配置与 README 改为当前 PostgreSQL/Celery 运维口径。

本次测试不创建真实 WooCommerce 商品、发货、客户邮件或库存变动。集成测试使用生产表结构的空库，订单、站点、用户及外部响应均为合成数据。

## 历史工作保存与处置

整理前已保存 24 个本地工作树、3 个独立 Git 仓库的所有分支和 stash，以及服务器 Git 历史、旧库存工作树差异。`git bundle` 通过校验；未提交文件、二进制补丁和旧源码副本保存在带 SHA256 清单的私有归档中。数据库、凭据、业务导出及原始归档不提交到此公共仓库。

| 历史内容 | 处置 |
| --- | --- |
| 产品克隆未提交差异 | 提取当前主线缺少的修复和测试；不覆盖主线已完成的 PostgreSQL 改造 |
| 离线利润脚本未提交改动 | 实现已在主线，补回应用上下文回归测试 |
| 邮件中心未跟踪代码 | 当前主线已包含 PostgreSQL 适配，保留当前实现、补回测试和示例 |
| 库存、补货重算、终态恢复、自动确认提交 | 实现已由 `e4fc8fb` 等后续提交收录；补回遗漏测试，原始提交另存归档引用与 bundle |
| 本地 stash `b94b216` | 站点负责人权限及状态逻辑已被后续实现覆盖，保存原始 stash |
| 服务器 stash `07550ed` | 邮件主题改动已在主线，保存原始 stash |
| 服务器旧库存工作树 | 鉴权修复已在主线，当前 PHP 模板更新；保存原始差异 |
| 旧 SQLite 诊断脚本、临时文件、部署副本 | 私有归档，不作为新功能并入主线 |
| 无独立 Git 的旧源码目录 | 保留旧版；不恢复已明确删除的利润页面及旧自动同步脚本 |
| 仅 CRLF/LF 差异 | 不作业务代码变更 |

未在远程按原 SHA 保留的 16 个历史提交归档如下；提交拓扑不同不代表功能未上线：

| 原提交 | 内容 |
| --- | --- |
| `c0d86f7` | 顺序站点 API 检查 |
| `e0a580a` | 自动确认生产基线 |
| `3003918` | 发货界面生产基线 |
| `e680ab9` | 持久化自动确认与包裹结果关联 |
| `fa8b0e8` | 规范域名重定向、确认队列限量 |
| `628b223` | 规范域名同步鉴权 |
| `a24a8d5` | 写入前中断任务恢复 |
| `b6066cf` | 收货、调拨、盘点与审批 |
| `f84db66` | 库存 PostgreSQL 迁移说明 |
| `96d80ad` | 缺货展示和临时盘点授权 |
| `bdf51cd` | 补货重算生产基线 |
| `6696c20` | 补货后缺货订单重算 |
| `87e14bb` | 发货终态生产基线 |
| `2ef2e27` | 终态证据恢复与部分发货保护 |
| `3a3424a` | 被拒绝发货的受控重试 |
| `7ea59df` | 订单超时和同步修复合并记录 |

旧工作树保留原分支和未提交内容，作为历史资料；不批量合并、强制重置或清理。以后不从这些目录部署。机器上的开发入口说明记录当前工作目录、私有备份位置及本次验收证据。

## 可重复的验证

本次 Linux 验证结果：离线 Python 测试 676 项通过，82 项按 PostgreSQL、浏览器等专项条件跳过；另行运行上述改动相关的 PostgreSQL 集成测试 51 项，全部通过；站点 API JavaScript 测试 7 项通过。Windows 离线测试亦通过，POSIX 文件权限用例在 Linux 补验。

离线测试在仓库根运行，只收集 `tests/`；根目录的历史 `test_*.py` 诊断脚本可能访问外部系统，不能加入收集范围。

```bash
WOO_DB_BACKEND=sqlite python -m pytest -q -p no:cacheprovider
node --test tests/site_api_checks.test.cjs
```

Windows PowerShell 使用 `$env:WOO_DB_BACKEND='sqlite'`，再执行相同 Python 命令。

PostgreSQL 集成测试需要另一份检出的源码，以及只从 `pg_dump --schema-only --no-owner --no-acl` 生成的空测试库。不能复制生产业务行。以下命令在隔离源码目录执行，使用本机 peer 认证，Redis/Celery 使用内存 broker；不要给测试进程加载生产服务的环境文件：

```bash
sudo -u postgres env PYTHONPATH="$PWD" PYTHONDONTWRITEBYTECODE=1 \
  WOO_DB_BACKEND=postgres WOO_DB_HOST=/var/run/postgresql \
  WOO_DB_USER=postgres WOO_DB_PASSWORD=unused_peer_auth \
  WOO_DB_NAME_OVERRIDE=woo_return_loss_test_unification_20260918 \
  WOO_DB_POOL_MAX=12 FLASK_SECRET_KEY=isolated-unification-test-only \
  CELERY_BROKER_URL=memory:// \
  /path/to/venv/bin/python -m pytest -q -p no:cacheprovider \
  tests/test_shipment_retry_postgres.py tests/test_site_sync_permissions.py \
  tests/test_mail_center_readonly_api.py tests/test_postgres_return_shipping_loss.py
```

这些集成测试拒绝非 `woo_return_loss_test_` 前缀的库名；新适配的测试还核对连接实际数据库名。外部 HTTP 被替换或阻止，任务发布不进入生产队列。此命令只覆盖列出的集成测试；其他 PostgreSQL 专项测试仍按各自隔离库要求运行。

生产验收另行核对已部署提交、服务进程与网页响应。测试通过不等于已经部署，也不等于已对真实外部系统执行过写入。
