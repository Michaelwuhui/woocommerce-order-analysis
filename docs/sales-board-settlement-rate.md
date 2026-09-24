# 销售看板实际结算汇率

在“销售看板 → 汇率来源”选择月份，按币种填写 **实际结算汇率（人民币/外币）**，即实际人民币到账金额除以相应外币回款金额。保存后重新计算该月销售看板和之后生成的 Excel 导出；已有导出文件不重写。清空该输入可恢复自动取值。

取值顺序：实际结算汇率 → 当月回款加权汇率 → 自定义备用汇率 → 系统汇率。原有备用汇率仍只在没有有效回款、也未设置实际结算汇率时生效。设置按月份、币种独立保存，只有管理员可修改；不修改合伙人回款记录或全局系统汇率。导出文件的“规则说明”会记录每种实际使用的汇率和来源。

PostgreSQL 部署前，先用数据库管理员执行 [004_sales_board_settlement_rates.sql](../migrations/postgresql/004_sales_board_settlement_rates.sql)。例如在生产主机上：

```bash
sudo -u postgres psql -X -v ON_ERROR_STOP=1 -1 -d woo_analysis \
  -f /www/wwwroot/woo-analysis/migrations/postgresql/004_sales_board_settlement_rates.sql
```

再核对表和应用账号权限后发布代码。SQLite 开发库由应用启动时创建该表。此迁移不写入任何汇率；8 月实际结算值必须经人工核对到账凭证后在页面填写，不能由系统推测。
