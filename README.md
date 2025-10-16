# WooCommerce 订单分析系统

一个功能完整的WooCommerce订单数据同步和分析系统，支持多站点数据拉取、订单状态分析、销售报告生成等功能。

## 🚀 功能特性

### 数据同步
- **多站点支持**: 同时支持多个WooCommerce网站的数据同步
- **增量同步**: 智能增量更新，避免重复拉取数据
- **错误重试**: 内置重试机制，确保数据同步的稳定性
- **代理支持**: 支持HTTP代理配置

### 数据分析
- **订单状态分析**: 支持5种订单状态（成功签收、发货未签收、缺货、取消订单、失败订单）
- **销售趋势分析**: 按月份、网站维度进行销售数据分析
- **产品数量统计**: 自动解析订单中的产品数量信息
- **多维度报告**: 生成详细的Excel分析报告

### 数据导出
- **Excel导出**: 支持多种格式的Excel报告导出
- **CSV导出**: 支持CSV格式的原始数据导出
- **自动格式化**: 自动调整列宽、添加样式和图表

## 📁 项目结构

```
woocommerceimport/
├── README.md                          # 项目说明文档
├── requirements.txt                   # Python依赖包
├── .gitignore                        # Git忽略文件配置
├── woo2woo.md                        # 详细使用文档
│
├── 数据同步脚本/
│   ├── wooorders.py                  # MySQL版本数据同步脚本
│   └── wooorders_sqlite.py           # SQLite版本数据同步脚本
│
├── 分析报告脚本/
│   ├── improved_sales_analysis.py    # 改进版销售分析（推荐）
│   ├── comprehensive_sales_analysis.py # 综合销售分析
│   ├── monthly_sales_report.py       # 月度销售报告
│   └── sales_analysis.py             # 基础销售分析
│
├── 数据导出脚本/
│   ├── enhanced_excel_export.py      # 增强版Excel导出（推荐）
│   ├── export_to_excel.py            # 基础Excel导出
│   └── export_to_csv.py              # CSV导出
│
└── 工具脚本/
    ├── check_db.py                   # 数据库状态检查
    ├── check_excel_content.py        # Excel内容检查
    └── analyze_line_items.py         # 订单项目分析
```

## 🛠️ 安装配置

### 1. 环境要求
- Python 3.8+
- SQLite3 或 MySQL 5.7+

### 2. 安装依赖
```bash
pip install -r requirements.txt
```

### 3. 配置WooCommerce API
在相应的脚本中配置您的WooCommerce API信息：

```python
sites = [
    {
        'name': 'your-site-name',
        'url': 'https://your-site.com',
        'consumer_key': 'your_consumer_key',
        'consumer_secret': 'your_consumer_secret'
    }
]
```

## 📖 使用指南

### 数据同步
```bash
# SQLite版本（推荐）
python wooorders_sqlite.py

# MySQL版本
python wooorders.py
```

### 生成销售分析报告
```bash
# 生成改进版销售分析报告（包含所有订单状态）
python improved_sales_analysis.py

# 生成月度销售报告
python monthly_sales_report.py
```

### 导出订单数据
```bash
# 导出增强版Excel报告
python enhanced_excel_export.py

# 导出基础Excel文件
python export_to_excel.py
```

### 检查数据状态
```bash
# 检查数据库状态
python check_db.py

# 检查Excel文件内容
python check_excel_content.py
```

## 📊 报告说明

### 销售分析报告特性
- **总体统计**: 总销售金额、订单数量、产品数量
- **状态分类**: 按订单状态分类统计
- **网站维度**: 按网站分别统计
- **时间维度**: 按月份统计销售趋势
- **自动格式化**: 专业的Excel格式和样式

### 支持的订单状态
- ✅ **成功签收** (completed)
- 🚚 **发货未签收** (on-hold)
- ⏳ **缺货** (processing)
- ❌ **取消订单** (cancelled)
- 💥 **失败订单** (failed)

## 🔧 高级配置

### 代理设置
```python
proxies = {
    'http': 'http://proxy-server:port',
    'https': 'https://proxy-server:port'
}
```

### 数据库配置
```python
# SQLite配置（推荐）
DATABASE_PATH = 'woocommerce_orders.db'

# MySQL配置
DB_CONFIG = {
    'host': 'localhost',
    'user': 'username',
    'password': 'password',
    'database': 'database_name'
}
```

## 🤝 贡献指南

1. Fork 本项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建 Pull Request

## 📝 更新日志

### v1.0.0 (2024-10-15)
- ✨ 初始版本发布
- 🚀 支持多站点WooCommerce数据同步
- 📊 完整的销售分析报告系统
- 📁 多格式数据导出功能
- 🔧 SQLite和MySQL双数据库支持

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 🆘 支持

如果您遇到问题或有建议，请：
1. 查看 [woo2woo.md](woo2woo.md) 详细文档
2. 创建 [Issue](../../issues)
3. 联系项目维护者

---

**注意**: 请确保在使用前正确配置WooCommerce API密钥，并根据您的实际需求调整脚本参数。