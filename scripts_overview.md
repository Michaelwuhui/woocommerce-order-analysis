# WooCommerce 脚本清单与建议

## 推荐的最新流程
1. **数据拉取**：优先使用 `wooorders_sqlite.py` 做多站点增量/修改同步到 `woocommerce_orders.db`（与各分析脚本数据源一致）。
2. **导出与分析**：根据需求选择  
   - 快速富格式导出：`enhanced_excel_export.py`  
   - 月度目标跟踪（2000 支/盈亏对比+图表）：`monthly_sales_report.py`  
   - 状态分层与站点对比：`improved_sales_analysis.py` 或版式更复杂的 `comprehensive_sales_analysis.py`  
   - 轻量月度概览：`monthly_stats_display.py`
3. **健康检查**：同步后可用 `check_db.py`、`analyze_line_items.py` 进行数据和字段校验。

## 脚本作用与保留建议

| 文件 | 功能概述 | 数据源/输出 | 使用建议 |
| --- | --- | --- | --- |
| `wooorders_sqlite.py` | 多站点 WooCommerce 订单增量/修改同步，写入 SQLite，含重试与字段标准化。 | 读 WooCommerce API → 写 `woocommerce_orders.db` | **主推**：当前默认拉取流程。 |
| `wooorders.py` | MySQL 版同步脚本，含代理/状态同步逻辑。 | 读 WooCommerce API → 写 MySQL `woocommerce_orders` | **旧管线**：仅在继续用 MySQL 时保留，否则可停用。 |
| `enhanced_excel_export.py` | 从 SQLite 导出增强版 Excel，附产品数量/SKU/客户信息，含统计与格式化。 | 读 SQLite → 写增强 Excel | **主推导出**：日常汇报用。 |
| `monthly_sales_report.py` | 月度销量目标跟踪（2000 支），站点拆分、每日/每周明细、图表输出。 | 读 SQLite → 多工作表 Excel + PNG 图 | **主推月报**：需要目标完成度/图表时使用。 |
| `improved_sales_analysis.py` | 按月按站点的状态分层（完成/在途/取消/失败）与汇总，格式化 Excel。 | 读 SQLite → Excel | **推荐分析**：替代早期 `sales_analysis.py`。 |
| `comprehensive_sales_analysis.py` | 用户定制版式的综合分析，含月度摘要、产品明细等多工作表。 | 读 SQLite → Excel | **按需**：版式复杂时使用；否则 `improved`/`monthly` 更轻便。 |
| `monthly_stats_display.py` | CLI 快速生成月度统计（成功/失败/取消）并输出 Excel+HTML。 | 读 SQLite → `monthly_stats_from_db.xlsx`、`monthly_stats.html` | **轻量**：快速预览或嵌入网页。 |
| `sales_analysis.py` | 早期汇总：站点总额/订单/产品数，简单模板 Excel。 | 读 SQLite → Excel | **可归档**：已被 `improved` 覆盖。 |
| `export_to_excel.py` | 基础的 SQLite 全量导出，无额外统计。 | 读 SQLite → Excel | **备用**：需原始数据时可用。 |
| `export_to_csv.py` | 从 MySQL 导出全部订单为 CSV。 | 读 MySQL → CSV | **旧管线**：仅在 MySQL 环境需要。 |
| `analyze_line_items.py` | 解析 `line_items`，输出产品种类/数量和月度目标完成度，便于调试。 | 读 SQLite → 控制台 | **保留作诊断**。 |
| `check_db.py` | 快速查看订单总数、各站点分布、最新订单。 | 读 SQLite → 控制台 | **保留作健康检查**。 |
| `check_excel_content.py` | 读取生成的 Excel，检查列和示例行。 | 读现有 Excel → 控制台 | **临时工具**：验证导出结果时使用。 |

## 老脚本是否需要
- 日常同步与分析建议统一走 `wooorders_sqlite.py` + 上述“主推”分析脚本；这套流程与当前 SQLite 数据源最匹配。  
- MySQL 相关 (`wooorders.py`, `export_to_csv.py`) 仅在仍需 MySQL 落地或历史数据迁移时使用。  
- 早期分析/导出脚本 (`sales_analysis.py`, `export_to_excel.py`) 功能已被增强版覆盖，可作为参考或回退，不必在常规流程中调用。  
- 诊断/验证脚本 (`analyze_line_items.py`, `check_db.py`, `check_excel_content.py`) 建议保留，用于排查数据质量或导出文件。
