# 公司盈利 Web 下线与离线月报

公司盈利页面、导航入口及 `/api/company-profit/*` 接口已永久下线。生产系统不提供财务报表查看、修改或导出入口，历史核算表只作为离线月报的数据源保留。

## 每月生成 Excel

在已配置 `cangfu_hk` SSH 别名的办公电脑执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\export_company_profit_month.ps1 -Month 2026-07
```

默认输出到代码仓库同级的 `outputs\company-profit\company-profit-monthly-2026-07.xlsx`。使用 ASCII 文件名是为了兼容 Windows PowerShell 5.1 的脚本编码行为；工作簿内部仍使用完整中文标题。

导出流程：

1. 生产服务器用 SQLite 只读连接创建一致性临时副本。
2. 核算算法只在临时副本上运行，不改写线上数据库。
3. 私有 JSON 快照通过 SSH 下载到 Windows 临时目录。
4. 使用 `@oai/artifact-tool` 生成 Excel，随后删除服务器和本机临时 JSON。

Excel 包含管理摘要、市场明细、收支明细、销售情景、口径说明和趋势数据。管理摘要及销售情景的核心结果使用公式连接明细表；工资仅展示团队汇总，不输出个人工资明细。

## 安全边界

- 快照程序拒绝把 JSON 写入网站应用目录。
- 快照文件权限为 `0600`，只允许当前系统用户读取。
- Web 路由必须保持不存在；不能新增浏览器导出按钮或下载 API。
- Excel 属于机密离线文件，应通过公司批准的文件渠道交付老板，不上传公共网盘。
- 数据不完整时仍可导出，但“管理摘要”和“口径说明”会明确标记待补数据，不能把预测当成已确认利润。
