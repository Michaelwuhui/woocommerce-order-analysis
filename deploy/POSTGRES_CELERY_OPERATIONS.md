# PostgreSQL、Redis 与 Celery 生产切换手册

本文只适用于 `/www/wwwroot/woo-analysis`。所有命令都必须在已测试并固定的
commit 上执行。不得把真实密码、Cookie、Consumer Key 或 Flask secret 写入命令
行、Git、日志或本手册。

## 不可破坏的安全边界

- PostgreSQL 和 Redis 只监听本机地址；Redis 不是业务状态权威，PostgreSQL 才是。
- 真实 `/etc/woo-analysis/woo-analysis.env` 必须为 `root:root`、`0600`，Web 和各
  worker 读取同一份文件。
- 最终迁移源必须来自 SQLite `.backup` 或 Python backup API；WAL 活跃时禁止只复制
  `woocommerce_orders.db` 主文件。
- 快速、自动、深度同步只能通过持久化 `sync_runs` pipeline 运行；切换前必须禁用旧
  cron，Beat 只能启动一个实例。
- 未经业务方再次确认维护窗口，不得停止生产服务、切换数据库或启用 Celery。
- 验收不得触发真实发货、添加运单、客户邮件或 WooCommerce 写操作。
- 当前工作树不得含硬编码 API 凭据；Git 历史中曾出现的任何凭据必须轮换。删除源码
  内容不等于凭据失效，也不能替代轮换。
- 任一关键校验失败立即停止，不在失败的 PostgreSQL 上继续营业。

## 连接预算与队列

- Gunicorn：4 processes × pool max 2 = 8。
- fetch worker：3 processes × pool max 2 = 6。每个进程最多占用一个站点 advisory
  lock 会话和一个短事务连接。
- writer worker：1 process × pool max 1 = 1。
- Beat、履约 worker、产品复制 worker：各 1。
- 在线备份：最多 1 个临时应用角色连接。
- 运维预留：5。

应用角色连接上限为 24；包含在线备份的预计峰值为 19，保留 5 个连接用于恢复和
管理检查。`sync_fetch` 每任务只携带一页 50–100 条的有界数据，worker concurrency
为 3；数据库 advisory lock 保证同站点分页串行。`sync_write` 独立 worker 固定
concurrency 为 1。Celery 只接受 JSON，开启 late acknowledgement、prefetch 1 和
visibility timeout；重复投递由 PostgreSQL receipt/outbox 幂等处理。

## 切换前门禁

以下条件必须同时满足，并记录证据：

1. 功能分支完整测试、迁移测试、恢复 harness、Python 语法和配置检查通过。
2. 已测试 SQLite 快照与 PostgreSQL 测试库 97 张业务表逐表数量、主键、全行摘要、
   金额/汇率/GMV、币种、布尔值、外键、唯一约束、索引、视图和 sequence 全部匹配。
3. 生产代码、配置、systemd、cron 和数据库备份目录存在，SHA256 已验证。
4. 回滚所需的发布前 commit、原 crontab 和原 systemd drop-in 已记录/备份。
5. 当前工作树密钥扫描无命中；Git 历史中仍与生产匹配的 API 凭据已在对应站点轮换，
   数据库已更新且只读连通检查通过。扫描和轮换记录只记数量/站点 ID，不记凭据值。
6. PostgreSQL 本机监听、SCRAM 和应用最小权限已核对；Redis 尚未暴露到公网。
7. 当前无活动同步批次、无遗留同步进程，且业务方确认无人发货或修改订单。
8. 业务方第二次明确确认已进入短暂维护窗口。

任一项不满足则结论为 HOLD。

## 预切换记录与备份

在一次 root shell 中设置一个真实时间戳；示例中的占位时间必须替换，不能原样执行：

```bash
CUTOVER_DIR=/www/backup/woo-analysis-pg-cutover-YYYYMMDD-HHMMSS
install -d -o root -g root -m 0700 "$CUTOVER_DIR"
git -C /www/wwwroot/woo-analysis rev-parse HEAD > "$CUTOVER_DIR/pre-cutover-head.txt"
git -C /www/wwwroot/woo-analysis status --short --branch > "$CUTOVER_DIR/pre-cutover-status.txt"
git -C /www/wwwroot/woo-analysis diff --quiet
git -C /www/wwwroot/woo-analysis diff --cached --quiet
crontab -l > "$CUTOVER_DIR/root.crontab"
cp -a /etc/systemd/system/woo-analysis.service "$CUTOVER_DIR/"
cp -a /etc/systemd/system/woo-analysis.service.d "$CUTOVER_DIR/"
cp -a /etc/systemd/system/woo-fulfillment-worker.service "$CUTOVER_DIR/"
cp -a /etc/systemd/system/woo-product-clone-worker.service "$CUTOVER_DIR/"
if test -e /etc/systemd/system/redis-server.service.d; then
  cp -a /etc/systemd/system/redis-server.service.d "$CUTOVER_DIR/"
fi
cp -a /etc/redis/redis.conf "$CUTOVER_DIR/redis.conf"
cp -a /etc/woo-analysis "$CUTOVER_DIR/etc-woo-analysis"
chmod -R go-rwx "$CUTOVER_DIR"
sha256sum "$CUTOVER_DIR/root.crontab" "$CUTOVER_DIR/redis.conf" \
  > "$CUTOVER_DIR/config-SHA256SUMS"
```

代码使用 `git bundle` 与只读 tar 归档；归档必须排除数据库、WAL、日志、备份和真实
环境文件：

```bash
git -C /www/wwwroot/woo-analysis bundle create "$CUTOVER_DIR/code.bundle" --all
tar --exclude=.git --exclude=woocommerce_orders.db \
  --exclude=woocommerce_orders.db-wal --exclude=woocommerce_orders.db-shm \
  --exclude='*.log' --exclude='*.env' --exclude='*.db' --exclude='*.dump' \
  -C /www/wwwroot -czf "$CUTOVER_DIR/code-files.tar.gz" woo-analysis
sha256sum "$CUTOVER_DIR/code.bundle" "$CUTOVER_DIR/code-files.tar.gz" \
  >> "$CUTOVER_DIR/config-SHA256SUMS"
sha256sum -c "$CUTOVER_DIR/config-SHA256SUMS"
```

## 安装受保护配置

首次切换时将模板的非秘密键复制到仓库外文件，再通过不会回显的交互方式填写密码和
Flask secret。不要用 shell tracing，且不得把真实文件复制回仓库。

```bash
install -d -o root -g root -m 0700 /etc/woo-analysis
test -e /etc/woo-analysis/woo-analysis.env || install -o root -g root -m 0600 \
  /www/wwwroot/woo-analysis/deploy/woo-analysis.env.example \
  /etc/woo-analysis/woo-analysis.env
stat -c '%U:%G %a %n' /etc/woo-analysis/woo-analysis.env
```

实际切换前只允许用“键名检查”确认内容完整，不输出值：

```bash
awk -F= '/^[A-Z0-9_]+=/ {print $1}' /etc/woo-analysis/woo-analysis.env | sort
```

## PostgreSQL 与 Redis 安全核对

PostgreSQL 应由系统稳定版服务管理；角色使用 SCRAM 密码，应用角色只拥有应用库的
连接权和 public schema 中运行所需的表/sequence 权限。核对时不要查询或输出密码：

```bash
systemctl is-active postgresql
sudo -u postgres psql -Atqc "SHOW listen_addresses"
sudo -u postgres psql -Atqc "SHOW password_encryption"
sudo -u postgres psql -Atqc \
  "SELECT rolname,rolsuper,rolcreatedb,rolcreaterole,rolconnlimit FROM pg_roles WHERE rolname IN ('woo_analysis_owner','woo_analysis_app') ORDER BY 1"
ss -ltnp | grep ':5432'
```

Redis 使用主配置末尾 include，确保本项目配置覆盖发行版默认值。修改前必须已经备份
`/etc/redis/redis.conf`：

```bash
install -o root -g redis -m 0640 \
  /www/wwwroot/woo-analysis/deploy/redis/woo-analysis.conf \
  /etc/redis/woo-analysis.conf
grep -qxF 'include /etc/redis/woo-analysis.conf' /etc/redis/redis.conf || \
  printf '\ninclude /etc/redis/woo-analysis.conf\n' >> /etc/redis/redis.conf
install -o root -g root -m 0644 \
  /www/wwwroot/woo-analysis/deploy/sysctl/99-woo-analysis-redis.conf \
  /etc/sysctl.d/99-woo-analysis-redis.conf
install -d -o root -g root -m 0755 \
  /etc/systemd/system/redis-server.service.d
install -o root -g root -m 0644 \
  /www/wwwroot/woo-analysis/deploy/systemd/redis-server.service.d/woo-analysis.conf \
  /etc/systemd/system/redis-server.service.d/woo-analysis.conf
systemctl daemon-reload
sysctl --system
```

Redis 启动后必须同时验证本机监听、AOF、RDB、noeviction 和 systemd 自动恢复：

```bash
systemctl enable --now redis-server
redis-cli ping
redis-cli CONFIG GET bind appendonly appendfsync save maxmemory-policy
ss -ltnp | grep ':6379'
systemctl is-active redis-server
```

`ss` 输出若出现非 loopback 地址立即停止 Redis 并 HOLD。

## 最终停写、SQLite 备份与迁移

第二次确认后按顺序执行。先从原 crontab 同时生成两个已校验版本：维护态会删除旧同步、
SQLite 备份、库存推送和物流结果等全部数据库写入者；PostgreSQL 恢复态永久删除前三类
SQLite 作业，并给库存推送和物流结果作业加载受保护的 PostgreSQL 环境。数量不符合预期
时脚本拒绝输出，且不会影响 Claude 修复脚本、面板任务或其他独立作业。Web 停止后新同步
请求自然被阻断：

```bash
crontab -l > "$CUTOVER_DIR/root.crontab.before-disable"
/www/wwwroot/woo-analysis/venv/bin/python \
  /www/wwwroot/woo-analysis/deploy/prepare_cutover_crontab.py \
  --mode maintenance \
  --input "$CUTOVER_DIR/root.crontab.before-disable" \
  --output "$CUTOVER_DIR/root.crontab.maintenance"
/www/wwwroot/woo-analysis/venv/bin/python \
  /www/wwwroot/woo-analysis/deploy/prepare_cutover_crontab.py \
  --mode postgres \
  --input "$CUTOVER_DIR/root.crontab.before-disable" \
  --output "$CUTOVER_DIR/root.crontab.postgres"
crontab "$CUTOVER_DIR/root.crontab.maintenance"
systemctl stop woo-analysis.service
systemctl stop woo-fulfillment-worker.service woo-product-clone-worker.service
pgrep -af 'auto_sync.py|full_resync_all.py|1.wooorders_sqlite.py|sync_all_sites|inv_push_cron.py|resolve_outcomes.py --live'
```

若最后一条仍显示旧同步进程，或仍有 `inv_push_cron.py`、`resolve_outcomes.py --live`，先
等待其短事务结束；不得直接杀死未知进程。确认无写入者
后，使用 Python SQLite backup API 生成最终一致性备份：

```bash
set -a
. /etc/woo-analysis/woo-analysis.env
set +a
export WOO_DB_BACKEND=sqlite
export WOO_SQLITE_PATH=/www/wwwroot/woo-analysis/woocommerce_orders.db
export WOO_BACKUP_DIR="$CUTOVER_DIR/final-sqlite"
export WOO_BACKUP_KEEP_LOCAL=0
/www/wwwroot/woo-analysis/venv/bin/python /www/wwwroot/woo-analysis/backup_db.py
```

验证 `.sha256` 和 manifest 中 `integrity_check=ok` 及所有逐表记录数。将归档解压成只读
迁移源，迁移过程不得直接读取活动 SQLite 文件：

```bash
SQLITE_GZ=$(find "$CUTOVER_DIR/final-sqlite" -maxdepth 1 -name 'woocommerce_orders_*.db.gz' -type f -print -quit)
sha256sum -c "$SQLITE_GZ.sha256"
gzip -cd "$SQLITE_GZ" > "$CUTOVER_DIR/final-migration-source.db"
chmod 0400 "$CUTOVER_DIR/final-migration-source.db"
sqlite3 "$CUTOVER_DIR/final-migration-source.db" 'PRAGMA integrity_check;'
sha256sum "$CUTOVER_DIR/final-migration-source.db" > "$CUTOVER_DIR/final-migration-source.db.sha256"
```

root-only 备份目录不能放宽权限。另建只允许 `postgres` 访问的暂存目录，将已校验快照
安装为只读副本；以 PostgreSQL 管理身份在同一流程中重建 schema、导入全部数据、
建立约束/索引、创建同步表并重置 sequence：

```bash
MIGRATION_STAGE=/var/lib/postgresql/woo-analysis-cutover-YYYYMMDD-HHMMSS
install -d -o postgres -g postgres -m 0700 "$MIGRATION_STAGE"
install -o postgres -g postgres -m 0400 \
  "$CUTOVER_DIR/final-migration-source.db" "$MIGRATION_STAGE/source.db"
sudo -u postgres /www/wwwroot/woo-analysis/venv/bin/python \
  /www/wwwroot/woo-analysis/migrations/postgresql/migrate.py \
  --sqlite "$MIGRATION_STAGE/source.db" \
  --database woo_analysis --reset --allow-production-reset \
  --report "$MIGRATION_STAGE/postgres-migration-report.json"
sudo -u postgres /www/wwwroot/woo-analysis/venv/bin/python \
  /www/wwwroot/woo-analysis/migrations/postgresql/migrate.py \
  --sqlite "$MIGRATION_STAGE/source.db" \
  --database woo_analysis --verify-only \
  --report "$MIGRATION_STAGE/postgres-verify-report.json"
install -o root -g root -m 0600 "$MIGRATION_STAGE/postgres-migration-report.json" \
  "$CUTOVER_DIR/postgres-migration-report.json"
install -o root -g root -m 0600 "$MIGRATION_STAGE/postgres-verify-report.json" \
  "$CUTOVER_DIR/postgres-verify-report.json"
```

两个命令都必须输出 `{"ok": true, "tables": 97}`。报告还必须证明全部 table count、
关键主键集合、全行摘要、numeric/currency、布尔值、外键孤儿基线、唯一约束、索引、
视图和 sequence 相符。任何 `match=false` 都必须 HOLD 并回滚。

## systemd 安装与启动顺序

安装前确认生产代码已经是已测试 commit。只安装明确列出的 unit/drop-in：

```bash
install -o root -g root -m 0644 deploy/systemd/woo-celery-fetch.service /etc/systemd/system/
install -o root -g root -m 0644 deploy/systemd/woo-celery-write.service /etc/systemd/system/
install -o root -g root -m 0644 deploy/systemd/woo-celery-beat.service /etc/systemd/system/
install -o root -g root -m 0644 deploy/systemd/woo-postgres-backup.service /etc/systemd/system/
install -o root -g root -m 0644 deploy/systemd/woo-postgres-backup.timer /etc/systemd/system/
install -d -o root -g root -m 0755 /etc/systemd/system/woo-analysis.service.d
install -d -o root -g root -m 0755 /etc/systemd/system/woo-fulfillment-worker.service.d
install -d -o root -g root -m 0755 /etc/systemd/system/woo-product-clone-worker.service.d
install -o root -g root -m 0644 deploy/systemd/woo-analysis.service.d/postgresql.conf /etc/systemd/system/woo-analysis.service.d/
install -o root -g root -m 0644 deploy/systemd/woo-fulfillment-worker.service.d/postgresql.conf /etc/systemd/system/woo-fulfillment-worker.service.d/
install -o root -g root -m 0644 deploy/systemd/woo-product-clone-worker.service.d/postgresql.conf /etc/systemd/system/woo-product-clone-worker.service.d/
install -o root -g root -m 0644 deploy/logrotate/woo-analysis-celery /etc/logrotate.d/
systemctl daemon-reload
systemd-analyze verify /etc/systemd/system/woo-celery-fetch.service \
  /etc/systemd/system/woo-celery-write.service \
  /etc/systemd/system/woo-celery-beat.service \
  /etc/systemd/system/woo-postgres-backup.service \
  /etc/systemd/system/woo-postgres-backup.timer
```

启动顺序固定为 PostgreSQL、Redis、串行 writer、fetch、唯一 Beat、其他后台 worker、
最后 Web：

```bash
systemctl enable postgresql redis-server woo-celery-write woo-celery-fetch \
  woo-celery-beat woo-postgres-backup.timer
systemctl start postgresql redis-server
systemctl start woo-celery-write
systemctl start woo-celery-fetch
systemctl start woo-celery-beat
systemctl restart woo-fulfillment-worker woo-product-clone-worker
systemctl start woo-analysis
systemctl start woo-postgres-backup.timer
systemctl is-active postgresql redis-server woo-celery-write woo-celery-fetch \
  woo-celery-beat woo-analysis woo-fulfillment-worker woo-product-clone-worker
crontab "$CUTOVER_DIR/root.crontab.postgres"
```

只有上述服务全部为 `active` 后才恢复两个数据库 cron；其余旧 SQLite 同步/备份作业不会
出现在 PostgreSQL crontab 中。安装后再次 `crontab -l` 检查两个保留作业都加载了
`/etc/woo-analysis/woo-analysis.env`，但不得输出该文件内容。

## 生产验收（不能只看 HTTP 200）

1. `systemctl is-active` 检查 PostgreSQL、Redis、两个 Celery worker、Beat、Gunicorn、
   履约 worker 和产品复制 worker；`is-enabled` 检查持久启动。
2. `ss -ltnp` 证明 5432、6379、5000 只在 loopback 监听。
3. `celery inspect ping/active/registered` 证明 fetch 有 3 个并发槽、writer 只有 1 个；
   `systemctl status` 证明 Beat 只有一个主进程。
4. 未登录请求必须被拒绝/跳转；使用测试客户端执行登录后主要页面/API 只读冒烟，检查
   JSON 类型、模板标志和实际数据库内容，不以状态码为唯一证据。
5. 查询 PostgreSQL `pg_stat_activity`，应用连接总量不得超过预算且无长期 idle in
   transaction。
6. 对照迁移报告抽样订单、站点、用户权限、备注、金额/币种/汇率和本地发货状态。
7. API 连续两次提交同步时只能存在一个活动 `run_id`；受控快速同步最多并行 3 个不同
   站点，同站点页码严格递增，writer 进程 concurrency=1。
8. 观察 `sync_runs`、`sync_site_progress`、`sync_page_receipts`、outbox 与 heartbeat；页面
   刷新或 Gunicorn reload 后从 PostgreSQL 恢复同一 run。
9. 在无真实外部副作用模式下验证 fetch/writer/Redis 重启后的重投和幂等；任何 receipt
   重复、订单/备注重复或进度永久转圈都必须回滚。
10. 检查日志中没有秘密、认证头和客户敏感数据，也没有持续增长的异常堆栈。

常用只读检查：

```bash
systemctl is-active postgresql redis-server woo-celery-write woo-celery-fetch \
  woo-celery-beat woo-analysis woo-fulfillment-worker woo-product-clone-worker
systemctl is-enabled redis-server woo-celery-write woo-celery-fetch \
  woo-celery-beat woo-analysis woo-postgres-backup.timer
ss -ltnp | grep -E ':(5432|6379|5000)[[:space:]]'
/www/wwwroot/woo-analysis/venv/bin/celery -A celery_app:celery_app inspect ping
journalctl -u woo-analysis -u woo-celery-fetch -u woo-celery-write \
  -u woo-celery-beat --since '-15 minutes' --no-pager
```

## 回滚到 SQLite

回滚不会撤销已经成功的外部物流/WooCommerce 操作。先读取 `external_operations` 与
outbox，把 `external_success` 但未 `local_committed/notified` 的记录交给人工对账；禁止
再次发货或再次通知。

在已停写状态下执行：

```bash
systemctl stop woo-analysis woo-celery-beat woo-celery-fetch woo-celery-write
systemctl stop woo-fulfillment-worker woo-product-clone-worker
systemctl disable woo-celery-beat woo-celery-fetch woo-celery-write \
  woo-postgres-backup.timer
```

保留失败 PostgreSQL 数据库用于审计，不 DROP。确认没有新产生的 tracked 修改；若有，
先归档并 HOLD，`git switch` 必须失败而不能覆盖。随后从发布前 commit 创建带时间戳的
紧急回滚分支（不移动远端 main），并把新增 PostgreSQL 环境 drop-in 移入本次备份目录
（可恢复，不删除）：

```bash
git -C /www/wwwroot/woo-analysis diff --quiet
git -C /www/wwwroot/woo-analysis diff --cached --quiet
ROLLBACK_BRANCH="codex/rollback-$(basename "$CUTOVER_DIR")"
git -C /www/wwwroot/woo-analysis switch -c "$ROLLBACK_BRANCH" \
  "$(cat "$CUTOVER_DIR/pre-cutover-head.txt")"
install -d -o root -g root -m 0700 "$CUTOVER_DIR/disabled-postgresql-dropins"
mv /etc/systemd/system/woo-analysis.service.d/postgresql.conf \
  "$CUTOVER_DIR/disabled-postgresql-dropins/"
mv /etc/systemd/system/woo-fulfillment-worker.service.d/postgresql.conf \
  "$CUTOVER_DIR/disabled-postgresql-dropins/fulfillment-postgresql.conf"
mv /etc/systemd/system/woo-product-clone-worker.service.d/postgresql.conf \
  "$CUTOVER_DIR/disabled-postgresql-dropins/product-clone-postgresql.conf"
if test -e /etc/systemd/system/redis-server.service.d/woo-analysis.conf; then
  mv /etc/systemd/system/redis-server.service.d/woo-analysis.conf \
    "$CUTOVER_DIR/disabled-postgresql-dropins/redis-woo-analysis.conf"
fi
crontab "$CUTOVER_DIR/root.crontab"
systemctl daemon-reload
```

原 SQLite/WAL 从未删除；正常回滚优先继续使用停写前原文件，并对最终备份做校验。只有
确认原文件损坏时，才把已验证归档恢复到一个新路径，校验后通过
`WOO_SQLITE_PATH` 指向它，不能覆盖原文件：

```bash
sha256sum -c "$SQLITE_GZ.sha256"
gzip -cd "$SQLITE_GZ" > "$CUTOVER_DIR/restored-for-rollback.db"
sqlite3 "$CUTOVER_DIR/restored-for-rollback.db" 'PRAGMA integrity_check;'
chmod 0400 "$CUTOVER_DIR/restored-for-rollback.db"
```

最后启动原 SQLite 服务链并做内容级验证：

```bash
systemctl restart woo-fulfillment-worker woo-product-clone-worker
systemctl start woo-analysis
systemctl is-active woo-analysis woo-fulfillment-worker woo-product-clone-worker
curl -fsS http://127.0.0.1:5000/ -o /dev/null
journalctl -u woo-analysis -u woo-fulfillment-worker \
  -u woo-product-clone-worker --since '-10 minutes' --no-pager
```

登录后确认订单数量、抽样订单/备注、权限、待发货与本地运单状态均来自 SQLite，再解除
维护窗口。Redis/PostgreSQL 可以保持本机运行以供审计，但 Celery/Beat 必须保持停止，
旧 cron 恢复后不得出现双重调度。

## 上线后备份与观察

PostgreSQL 备份 timer 每小时运行一次，`backup_db.py` 使用 `pg_dump` custom format，
生成 SHA256 和不含秘密的逐表数量 manifest。首次 timer 运行后必须执行：

```bash
systemctl start woo-postgres-backup.service
systemctl status woo-postgres-backup.service --no-pager
find /www/backups/woo-orders -maxdepth 1 -name 'woo_analysis_*.dump' -type f -print
pg_restore --list /www/backups/woo-orders/woo_analysis_YYYYMMDD_HHMMSS.dump >/dev/null
sha256sum -c /www/backups/woo-orders/woo_analysis_YYYYMMDD_HHMMSS.dump.sha256
```

上线后至少观察一个完整快速同步和一个计划调度周期。保留原 SQLite、WAL、所有备份、
旧配置、迁移报告和失败 PostgreSQL 数据，直到业务方书面确认回滚窗口结束。
