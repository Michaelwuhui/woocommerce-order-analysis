# WooCommerce 订单分析系统

一个功能完整的 WooCommerce 订单数据同步和分析系统，提供多站点数据拉取、订单状态分析、客户分析、产品分析、销售报告生成等功能。

---

## 📋 目录

1. [功能特性](#-功能特性)
2. [技术架构](#-技术架构)
3. [项目结构](#-项目结构)
4. [安装部署](#-安装部署)
5. [配置说明](#️-配置说明)
6. [使用指南](#-使用指南)
7. [用户权限](#-用户权限)
8. [API 接口](#-api-接口)
9. [数据库结构](#-数据库结构)
10. [常见问题](#-常见问题)

---

## 🚀 功能特性

### 仪表板 (Dashboard)
- **实时统计**: 总订单数、总销售额、平均订单金额、有效订单等核心指标
- **多货币支持**: 支持 PLN、AUD、AED 等多种货币，自动换算为 CNY
- **销售趋势图表**: 按日/月展示订单量和销售额趋势
- **订单状态分布**: 饼图展示各状态订单占比
- **网站统计**: 按网站维度统计销售数据
- **最近订单**: 显示最新订单列表及客户信息

### 订单管理 (Orders)
- **高级筛选**: 支持按网站、状态、日期范围、关键词筛选
- **网站负责人筛选**: 按负责人快速筛选订单
- **客户属性显示**: 在订单列表中显示客户等级和购买次数
- **订单详情弹窗**: 点击订单号查看完整订单信息
- **客户分析弹窗**: 点击客户名称查看客户画像

### 月度统计 (Monthly)
- **月度汇总**: 按月份统计订单和销售数据
- **网站对比**: 对比不同网站的月度表现
- **趋势分析**: 查看销售额和订单量的月度变化

### 客户管理 (Customers)
- **客户等级**: 自动计算客户等级 (VIP/优质/普通/新客/劣质)
- **手动调整**: 支持手动设置客户等级
- **购买历史**: 查看客户的完整购买记录
- **消费分析**: 统计客户总消费、购买频率、常购商品

### 产品分析 (Products)
- **销量排行**: 按销量排序的产品列表
- **品牌分析**: 按品牌统计销售数据
- **口味/规格分析**: 支持按口味、puffs 数等维度分析
- **产品映射**: 将原始产品名称映射为标准化名称

### 系统设置 (Settings)
- **网站管理**: 添加/编辑/删除 WooCommerce 网站
- **数据同步**: 手动触发或设置自动同步
- **汇率管理**: 配置各货币对 CNY 的汇率
- **产品映射**: 管理产品名称标准化规则
- **品牌管理**: 管理产品品牌分类

### 用户管理 (Users)
- **多用户支持**: 创建和管理多个用户账户
- **权限控制**: 按网站分配访问权限
- **角色管理**: 管理员和普通用户角色

---

## 🏗 技术架构

### 后端技术栈
| 技术 | 用途 |
|------|------|
| **Python 3.8+** | 后端开发语言 |
| **Flask** | Web 框架 |
| **Flask-Login** | 用户认证 |
| **SQLite3** | 数据库 |
| **Gunicorn** | WSGI 服务器 |
| **WooCommerce REST API** | 数据同步 |

### 前端技术栈
| 技术 | 用途 |
|------|------|
| **Bootstrap 5** | UI 框架 |
| **Chart.js** | 图表可视化 |
| **Bootstrap Icons** | 图标库 |
| **Jinja2** | 模板引擎 |
| **JavaScript (ES6)** | 前端交互 |

### 系统架构图

```
┌─────────────────────────────────────────────────────────────────┐
│                         用户浏览器                               │
└──────────────────────────────┬──────────────────────────────────┘
                               │ HTTP/HTTPS
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Nginx 反向代理                              │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Gunicorn WSGI 服务器                          │
│                    (systemd 管理)                                │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Flask 应用 (app.py)                         │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐            │
│  │ 仪表板   │ │ 订单管理 │ │ 客户管理 │ │ 产品分析 │            │
│  └──────────┘ └──────────┘ └──────────┘ └──────────┘            │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐                         │
│  │ 月度统计 │ │ 系统设置 │ │ 用户管理 │                         │
│  └──────────┘ └──────────┘ └──────────┘                         │
└──────────────────────────────┬──────────────────────────────────┘
                               │
           ┌───────────────────┼───────────────────┐
           ▼                   ▼                   ▼
┌─────────────────┐ ┌─────────────────┐ ┌─────────────────┐
│  SQLite 数据库   │ │ WooCommerce API │ │   自动同步任务   │
│ (orders, sites, │ │  (多站点数据源)   │ │  (auto_sync.py) │
│  users, etc.)   │ │                  │ │                 │
└─────────────────┘ └─────────────────┘ └─────────────────┘
```

---

## 📁 项目结构

```
woo-analysis/
├── app.py                     # 主应用程序 (Flask 路由和业务逻辑)
├── sync_utils.py              # 数据同步工具函数
├── auto_sync.py               # 自动同步脚本
├── requirements.txt           # Python 依赖
│
├── templates/                 # Jinja2 模板
│   ├── base.html              # 基础布局模板 (导航栏、弹窗)
│   ├── login.html             # 登录页面
│   ├── dashboard.html         # 仪表板页面
│   ├── orders.html            # 订单列表页面
│   ├── monthly.html           # 月度统计页面
│   ├── customers.html         # 客户管理页面
│   ├── products.html          # 产品分析页面
│   ├── settings.html          # 系统设置页面
│   └── users.html             # 用户管理页面
│
├── static/                    # 静态资源
│   └── css/                   # 样式文件
│
├── woocommerce_orders.db      # 主数据库文件
├── auto_sync.log              # 同步日志
├── gunicorn.log               # 服务日志
│
└── venv/                      # Python 虚拟环境
```

---

## 📦 安装部署

### 环境要求

- **操作系统**: Linux (推荐 Ubuntu 20.04+)
- **Python**: 3.8 或更高版本
- **Web 服务器**: Nginx (反向代理)
- **进程管理**: systemd

### 第一步: 克隆项目

```bash
cd /www/wwwroot
git clone <repository-url> woo-analysis
cd woo-analysis
```

### 第二步: 创建虚拟环境

```bash
python3 -m venv venv
source venv/bin/activate
```

### 第三步: 安装依赖

```bash
pip install -r requirements.txt
```

`requirements.txt` 内容：
```
flask
flask-login
gunicorn
requests
openpyxl
```

### 第四步: 初始化数据库

> ⚠️ **重要**: 首次部署后，必须先初始化数据库才能正常访问前端页面！

**4.1 创建数据库并初始化表结构**

运行初始化脚本创建必要的数据库表：

```bash
# 激活虚拟环境
source venv/bin/activate

# 运行数据库初始化脚本 (创建表结构)
python 1.wooorders_sqlite.py
```

该脚本会创建 `woocommerce_orders.db` 数据库文件及所有必需的表。

**4.2 初始化客户设置表 (可选)**

```bash
python init_customer_settings.py
```

> **注意**: 如果跳过此步骤直接启动服务，访问首页会报错，因为数据库表不存在。

### 第五步: 配置 Gunicorn 服务

创建 systemd 服务文件 `/etc/systemd/system/woo-analysis.service`:

```ini
[Unit]
Description=WooCommerce Analysis Web App
After=network.target

[Service]
User=www
Group=www
WorkingDirectory=/www/wwwroot/woo-analysis
Environment="PATH=/www/wwwroot/woo-analysis/venv/bin"
ExecStart=/www/wwwroot/woo-analysis/venv/bin/gunicorn --workers 4 --bind 127.0.0.1:5000 app:app
ExecReload=/bin/kill -HUP $MAINPID
Restart=always

[Install]
WantedBy=multi-user.target
```

启动服务：
```bash
sudo systemctl daemon-reload
sudo systemctl enable woo-analysis
sudo systemctl start woo-analysis
```

### 第六步: 配置 Nginx

在 Nginx 配置中添加：

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

重启 Nginx：
```bash
sudo systemctl reload nginx
```

### 第七步: 配置自动同步 (可选)

添加 crontab 定时任务：

```bash
crontab -e
```

添加以下内容 (每小时同步一次):
```
0 * * * * cd /www/wwwroot/woo-analysis && /www/wwwroot/woo-analysis/venv/bin/python auto_sync.py >> auto_sync.log 2>&1
```

---

## ⚙️ 配置说明

### 默认管理员账户

首次安装后，使用以下账户登录：

| 用户名 | 密码 | 角色 |
|--------|------|------|
| admin | admin123 | 管理员 |

> ⚠️ **重要**: 请在首次登录后立即修改密码！

### 添加 WooCommerce 网站

1. 登录后进入 **设置** 页面
2. 在 **网站管理** 区域点击 **添加网站**
3. 填写以下信息：
   - **网站 URL**: 例如 `https://www.your-store.com`
   - **Consumer Key**: WooCommerce REST API Key
   - **Consumer Secret**: WooCommerce REST API Secret
   - **负责人**: 网站负责人姓名 (可选)
4. 点击保存后，点击 **同步** 按钮拉取订单数据

### 获取 WooCommerce API 密钥

1. 登录 WordPress 后台
2. 进入 **WooCommerce → 设置 → 高级 → REST API**
3. 点击 **添加密钥**
4. 权限选择 **只读** (Read)
5. 点击生成并复制 Consumer Key 和 Consumer Secret

### 汇率配置

1. 进入 **设置 → 汇率管理**
2. 为每种货币设置对 CNY 的汇率
3. 支持按月份设置不同汇率

---

## 📖 使用指南

### 仪表板

仪表板是系统的首页，展示核心业务指标：

- **顶部卡片**: 显示总订单数、总销售额、平均订单金额、有效订单数
- **筛选器**: 支持按时间范围、网站、负责人筛选数据
- **趋势图表**: 展示订单量和销售额的历史趋势
- **网站统计表**: 按网站和货币维度的详细统计
- **最近订单**: 最新的 10 条订单，包含客户等级和购买次数

### 订单列表

订单列表页面支持详细的订单查询和分析：

- **筛选功能**: 按网站、状态、日期、负责人、关键词筛选
- **快捷日期**: 今天、昨天、本周、本月、上月、今年
- **月份分页**: 按月份快速切换查看
- **客户信息**: 显示客户等级徽章 (VIP/优质/普通/新客/劣质) 和购买次数

### 客户等级说明

系统自动根据以下因素计算客户等级：

| 等级 | 条件 | 图标 |
|------|------|------|
| **VIP** | 高消费 + 高频购买 + 多次成功订单 | ⭐ 金色 |
| **优质** | 较高消费或较高购买频率 | 💎 绿色 |
| **普通** | 一般消费水平 | ✓ 蓝色 |
| **新客** | 首次购买客户 | ✨ 青色 |
| **劣质** | 手动标记的问题客户 | ✖ 红色 |

> 管理员可以在客户详情弹窗中手动调整客户等级。

---

## 🔐 用户权限

### 角色说明

| 角色 | 权限 |
|------|------|
| **管理员** (admin) | 可访问所有功能，包括用户管理和系统设置 |
| **普通用户** (user) | 只能查看被分配权限的网站数据 |

### 权限分配

1. 进入 **用户管理** 页面
2. 选择用户并点击 **编辑**
3. 在 **网站权限** 中勾选允许访问的网站
4. 保存更改

> 普通用户只能看到被分配的网站的订单、客户、产品数据。

---

## 🔌 API 接口

### 订单详情

```
GET /api/order/<order_id>
```

返回指定订单的完整信息。

### 客户详情

```
GET /api/customer/<email>
```

返回客户的分析数据，包括购买历史、消费统计、常购商品等。

### 更新客户等级

```
POST /api/customer/quality
Content-Type: application/json

{
    "email": "customer@example.com",
    "quality": "vip"  // vip, good, normal, new, bad, auto
}
```

### 同步网站数据

```
POST /api/sync/<site_id>
```

触发指定网站的数据同步。

---

## 🗄 数据库结构

### 主要表结构

| 表名 | 说明 |
|------|------|
| `orders` | 订单数据 |
| `sites` | 网站配置 |
| `users` | 用户账户 |
| `user_site_permissions` | 用户网站权限 |
| `customer_settings` | 客户手动设置 |
| `exchange_rates` | 汇率配置 |
| `product_mappings` | 产品映射规则 |
| `brands` | 品牌数据 |
| `settings` | 系统设置 |

### orders 表字段

| 字段 | 类型 | 说明 |
|------|------|------|
| id | INTEGER | 订单 ID (主键) |
| number | TEXT | 订单号 |
| status | TEXT | 订单状态 |
| currency | TEXT | 货币 |
| total | REAL | 订单总额 |
| shipping_total | REAL | 运费 |
| billing | TEXT | 账单信息 (JSON) |
| shipping | TEXT | 收货信息 (JSON) |
| line_items | TEXT | 订单商品 (JSON) |
| date_created | TEXT | 创建时间 |
| source | TEXT | 来源网站 |

---

## ❓ 常见问题

### Q: 同步数据时显示成功但没有订单?

**A**: 可能原因：
1. WooCommerce 网站上没有订单
2. API 密钥权限不足
3. 网站 URL 格式不正确 (应包含 `https://`)

检查同步日志 `auto_sync.log` 获取详细信息。

### Q: 用户只能看到部分网站?

**A**: 检查用户权限设置。进入 **用户管理** 确认该用户被分配了正确的网站权限。

### Q: 汇率显示为 `-`?

**A**: 需要在 **设置 → 汇率管理** 中添加对应货币和月份的汇率。

### Q: 如何重启服务?

**A**: 使用以下命令：
```bash
# 方式1: systemctl (推荐)
sudo systemctl restart woo-analysis

# 方式2: 发送 HUP 信号
kill -HUP <gunicorn_master_pid>
```

### Q: 如何查看服务日志?

**A**: 
```bash
# 应用日志
tail -f /www/wwwroot/woo-analysis/gunicorn.log

# 同步日志
tail -f /www/wwwroot/woo-analysis/auto_sync.log

# systemd 日志
journalctl -u woo-analysis -f
```

---

## 📝 更新日志

### v2.0.0 (2024-12)
- ✨ 全新 Web 界面
- 🔐 用户认证和权限管理
- 📊 仪表板和图表可视化
- 👥 客户等级自动计算
- 💱 多货币和汇率支持
- 🏷️ 产品映射和品牌管理

### v1.0.0 (2024-10)
- 🚀 初始版本发布
- 📦 多站点数据同步
- 📑 Excel 报告导出

---

## 📄 许可证

本项目采用 MIT 许可证。

---

## 🆘 技术支持

如遇问题，请：
1. 查看本文档的 **常见问题** 部分
2. 检查日志文件获取错误信息
3. 联系系统管理员