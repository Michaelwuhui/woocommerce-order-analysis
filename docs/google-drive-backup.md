# Google Drive 日备份与失败提醒

Google Drive 日备份独立于现有 PostgreSQL → R2 小时备份。每天北京时间 04:20（最多随机延迟 3 分钟）选择不超过 2 小时、SHA256 和 `pg_restore --list` 均通过的 PostgreSQL 备份，打包当前 Git 提交、指定的运行配置及恢复清单。

## 数据与权限

- 目标为现有备份账号的私有目录。备份包含订单、配置与恢复凭据，不能公开共享。
- 凭据仅存 `/etc/woo-analysis/drive-backup-oauth.json`，格式为 `client_id`、`client_secret`、`refresh_token`；不得提交仓库或输出到日志。
- `/etc/woo-analysis/drive-backup.json` 指定目标目录、运行配置文件及保留天数；示例见 `deploy/drive-backup.json.example`。配置及状态目录均为 root 私有。
- 生产存在未提交的受跟踪代码修改时，备份会明确失败，避免归档提交与运行代码不一致。
- 每份归档包含 `database/`、`app/source.tar.gz`、`config/runtime-config.tar.gz` 和 `recovery.json`。其中保存内容校验值与代码提交，运行配置按主机清单选取。
- 本地上传暂存文件在云端核验后清除。Drive 中只轮换本任务专用目录、带 `managed_by=woo-analysis-daily-v1` 标记且超过保留期的文件，移入回收站；不处理历史宝塔备份。

## 上传与失败判定

预先保存 Google 文件 ID 和续传会话；发生网络中断时读取服务器确认的上传偏移后续传。Google 限额、429 和暂时性 5xx 使用有上限的指数退避。成功必须回读文件大小和 Google 内容校验值；HTTP 上传完成本身不会写成功状态。实现依据 [Google 续传上传](https://developers.google.com/workspace/drive/api/guides/manage-uploads)和[错误处理](https://developers.google.com/workspace/drive/api/guides/handle-errors)。

状态和接收地址在“系统设置 → 数据备份与灾备”中展示，仅管理员可见。备份不依赖 Flask、Celery 或数据库写入即可发送告警。

## 邮件与健康检查

`backup-alerts.json` 使用明确指定的 SMTP 账号和单一接收地址。发送失败不会被视为已通知，每 15 分钟检查会继续尝试；同一持续故障已发成功后 6 小时内不重复发送。恢复并核验备份后发送恢复通知。

健康检查覆盖备份程序退出失败、定时器未运行、超过 26 小时无成功备份及首次配置超过 2 小时仍未完成。systemd 在进程超时或退出失败时也触发检查。主机完全断网或关机时无法从本机发出邮件，需独立于本机的监控才能覆盖该情况。

## 部署与验收

1. 准备私有 OAuth、备份配置和 SMTP 配置，权限 `0600`；创建 `/var/lib/woo-analysis-backup`，权限 `0700`。
2. 初始化该目录的 `status.json`，记录 UTC `configured_at`、`schedule`、`alert_recipient`、`drive_folder_url` 和 `retention_days`。
3. 安装 `deploy/systemd/woo-drive-backup*`，执行 `systemctl daemon-reload` 并启用两个 timer。
4. 先执行 `backup_drive.py test-email`，核对邮件接收；再启动一次 `woo-drive-backup.service`。
5. 核对云端归档元数据、SHA256、归档恢复清单，检查定时器和管理员 API。恢复演练必须使用隔离数据库，不能覆盖生产库。

运行：`systemctl start woo-drive-backup.service`。日志：`journalctl -u woo-drive-backup -u woo-drive-backup-health`。不要直接修改上传收据或把上传地址复制到日志中。

恢复时先验证 `recovery.json` 的文件校验值，再将 PostgreSQL 归档恢复到空的隔离库进行检查。审查配置及服务版本后制定生产切换，深度同步不能代替数据库恢复。
