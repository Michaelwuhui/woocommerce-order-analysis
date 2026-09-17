# 跨站库存同步

入口：产品管理 → 跨站库存同步（`/product-manager/stock-sync`）。本功能默认关闭，迁移与 worker 均需显式安装。仅新增库存控制、快照、计划、任务与审计表；不改动库存账、订单或商品内容。

## 需求与实现选择

依据《订单系统跨站库存同步开发文档 v1.0》R01–R15 及第一版范围，实现四种操作：统一设为售完、解除指定人工停售并重新计算、按实际可售数量同步、以站点为参照同步状态。实现按已确认映射匹配、SKU 多仓供货、销售单位换算、负责人权限、跨分页全选/部分口味、目标交集、差异预览、持久队列、回读、取消、失败重试与历史记录。

采用的建议默认值：预览 10 分钟、目录 1 小时；人工控制持续有效；超级管理员控制只能由超级管理员解除；共享参照只开放读取；混合供货使用状态管理；缺货预订列冲突。初始一个 worker 顺序处理，每个端点仅允许一个在途写入，可启动多个进程，数据库领取和资源租约会互斥。没有新增自动定时同步。

未采用建议的自动三次写重试：超时先回读，无法确认则隔离资源；确定失败只生成新的人工确认预览。429/5xx 的 Retry-After 持久保存，冷却期间阻止新请求；缺少有效头时分别等待 60/10 秒。恢复时不重放旧数量。该取舍减少未确认请求被重复发送的机会。

## 启用顺序

部署须单独授权。本次开发未执行以下生产步骤。

1. 核对当前生产代码和数据库备份；停止旧库存发布定时任务，排空旧进程的在途库存请求。不要覆盖生产未提交修复。
2. 使用数据库 schema owner、与 Web 相同的数据库配置运行 `python stock_sync_schema.py`。迁移可重复运行，PostgreSQL 使用原生 BOOLEAN/TIMESTAMPTZ，不在 Flask 启动时执行 DDL。
3. 运行账号需要新表 SELECT/INSERT/UPDATE、`stock_sync_resource_leases` 的 DELETE，以及既有用户、站点、映射和库存表的读取权限。迁移表只需读取。新增表没有依赖自动增长序列。先核对实际 Web/worker 角色再授予权限，不授予既有库存账写权限。
4. Web 与 worker 均显式设置 `STOCK_SYNC_ENABLED=1`；复制并调整 `deploy/woo-stock-sync-worker.service` 和 `deploy/stock-sync.env.example`。生产现有服务使用 `/etc/woo-analysis/woo-analysis.env`，新 worker 模板复用该文件；Web 还需加载新开关。模板中工作目录、运行用户和解释器必须按目标主机核对。
5. 启动独立 worker：`python -u stock_sync_worker.py`；检查日志中的 worker ID。Flask/Gunicorn 只负责 API 与页面，不创建业务线程。
6. 先使用授权的可逆测试商品：预览 → 确认 → 逐项回读 → 实际站点查看 → 按当前依据恢复。主子站和插件必须单独验证后才开放对应能力；不要把本机 HTTP 模拟接口结果视为生产插件验收。

只运行一次队列领取可用 `python stock_sync_worker.py --once`。浏览器关闭不影响已经提交的后台任务。

## 库存与范围

- 使用 `inv_site_sku_map` 中启用的明确 product/variation ID；不按名称搜索，不创建缺失商品，不自动路由到主站。
- 数量来自启用的 SKU—仓库关系与市场供货关系。当地数量是 `max(on_hand-reserved,0)`；多仓按物理池去重后合计，安全库存取共享池参与站策略的最大值，扣一次，再按目标 `qty_per_item` 向下取整。
- 外部 WMS 必须明确设置 `stock_sync_available_includes_oms_reservations: true`，表示其可售数已涵盖 OMS 预留。默认新鲜度为 10800 秒，可用 `stock_sync_freshness_seconds` 配置。缺数据、口径不明、订单同步/预留后处理积压、过期都显示冲突，不解释为零。
- `physical_stock_pool` 可标识多条仓库记录对应同一真实库存池。别名数量一致时计一次，数据矛盾时停止。
- 人工停售优先于参照停售；有限数量仍以仓库当前数据为准。恢复不回填旧数量。原先数量管理的商品不能因数量未知而改成无限有货。
- 参照是只读来源。20→3 和 3→20 都只处理明确交集；同一源 SKU 在任务内依据变化时，剩余项停止并要求重新预览。
- 超级管理员按现有 `username == admin` 识别；普通 `role=admin` 不扩权。负责人必须具有产品管理权限，且名称与站点负责人匹配。执行、重试、记录读取均复核当前权限。

## 管理员维护

命令行维护需要本机/服务器数据库访问权，并校验当前有效超级管理员身份。先不传 `--confirm`，读取具体变更和 `confirm_hash`；核对后重跑同一命令，附上该哈希。任何相关数据改变都必须重新核对。维护过程不写 WooCommerce。

登记已验证的子站独立接口能力：

```text
python stock_sync_admin.py --actor-id <admin_id> capability --site-id <site_id> --evidence "独立测试记录路径及结果"
python stock_sync_admin.py --actor-id <admin_id> capability --site-id <site_id> --evidence "同一证据" --confirm <confirm_hash>
```

撤销能力加 `--disable`，同样先预览。能力绑定站点 URL、主站关系和凭据身份；配置变化后需重新验证，不能沿用旧证据。

已接入商品的映射被明确修正后：

```text
python stock_sync_admin.py --actor-id <admin_id> rebind --map-id <mapping_id> --reason "映射修正原因"
python stock_sync_admin.py --actor-id <admin_id> rebind --map-id <mapping_id> --reason "同一原因" --confirm <confirm_hash>
```

预览列出旧绑定、当前映射、将归档的全部旧控制和新基线。确认取得新旧资源锁，保留历史证据，归档旧控制，建立新绑定；不直接恢复销售。随后必须重新生成普通差异预览。原物理资源仍受旧入口保护，即使其映射被删除，也不能绕过新控制。删除映射不是解绑命令。

## 异常恢复与停用

资源租约不会因超时自动被另一个 worker 抢占。先停止并确认旧 worker 进程及其请求已经结束，再执行只读核实：

```text
python stock_sync_worker.py --recover-worker <日志中的完整worker_id> --confirm-worker-stopped
```

恢复只读取网站并补齐本地结果，不再次 PUT。无法核实的资源继续隔离；网站恢复后可再次执行同一核实。确定失败的项目通过页面重新生成预览，成功项目不能加入重试。

取消仅停止未开始的项目；在途结果仍核实，已经保存的控制不撤销。控制已保存但网站写入失败会在页面明确显示，不能据此宣称网站已经停售。

停用时先停止接收新任务并处理在途项、停止 worker，再关闭 Web 开关。保留新表、控制和旧入口保护；仅把开关设为 0 不会移除保护。不要直接删除表或回滚到没有保护的旧写入代码。需要取消接管时应另行核对每个资源及有效控制。

## 本地验证

专项测试不读取生产数据：

```text
python -m pytest -q tests/test_stock_sync.py tests/test_stock_sync_resilience.py
```

SQLite 使用临时数据库。PostgreSQL 专项另设 `STOCK_SYNC_TEST_POSTGRES=1`，以及现有 `WOO_DB_*` 连接参数和以 `woo_stock_sync_test` 开头的专用数据库名。每个测试建立随机隔离 schema 并清理，不能将此开关指向生产。

完整项目验收使用 `scripts/stock_sync_demo.py`，限定本机 PostgreSQL `127.0.0.1:55436/woo_stock_sync_test_browser`，要求先导入空的现有生产结构（仅 schema）。脚本注入人工测试数据，启动完整 Flask、独立 worker 和本机 HTTP Woo 测试服务。它没有生产登录旁路；便利登录路由只注册在本测试入口，且仅监听 127.0.0.1。

```text
python -u scripts/stock_sync_demo.py
```

浏览器测试地址 `http://127.0.0.1:5106/__stock_demo_login/1`（管理员）和 `/__stock_demo_login/2`（负责人）；HTTP 测试商品位于端口 8106。测试数据、日志、PostgreSQL 目录在忽略的 `.test-stock-sync/` 内，不能部署或提交。

Windows 验证设置 `PYTHONUTF8=1` 和 `PYTHONIOENCODING=utf-8`。完整回归以 `python -m pytest -q tests` 为入口；仓库其他目录中存在历史工具脚本，不作为测试发现根。

## 实现位置

`stock_sync_policy.py` 纯决策；`stock_sync_supply.py` 库存口径；`stock_sync_permissions.py` 当前权限；`stock_sync_woo.py` 精确接口与限流；`stock_sync_catalog.py` 完整目录；`stock_sync_planner.py` 不可变预览；`stock_sync_jobs.py` 幂等与租约；`stock_sync_worker.py` 持久执行与核实；`stock_sync_guard.py` 旧入口保护；`stock_sync_admin.py` 经核对的维护；`stock_sync_api.py` 页面 API；`stock_sync_schema.py` 显式迁移。

本系统的锁只能协调本系统写入，不能锁住 Woo 后台人工修改或第三方插件；通过前后读取发现外部变化，不承诺跨站原子更新。
