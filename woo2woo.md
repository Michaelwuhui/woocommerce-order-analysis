# WooCommerce 外链商品导出与多站点导入脚本使用说明

## 1. 功能概览

该脚本实现了 WooCommerce 变量商品的 **导出、拆分、转换为外链商品并导入到多个站点** 的功能，包括：

* 拆分 `pa_smaki` 多值属性生成单独商品
* 美化 SKU 和商品名称
* 自动匹配主图和 Gallery
* 自动生成 CSV（含 `Source URL` 和 `Button Text`）
* 多站点导入（商品类型为外链商品）
* 跳过已存在 SKU
* 导入失败自动重试 3 次
* 缺失图片自动记录
* 日志记录导出和导入
* 批量导入进度条显示

---

## 2. 环境要求

* Python 3.8+
* 安装依赖库：

```bash
pip install woocommerce tqdm
```

* WooCommerce 站点开启 REST API 并生成 `Consumer Key` 和 `Consumer Secret`

---

## 3. 配置说明

在脚本顶部修改以下配置：

```python
# 导出站点
EXPORT_SITE = {
    "url": "https://source-store.com",
    "consumer_key": "ck_source_xxx",
    "consumer_secret": "cs_source_xxx"
}

# 导入站点，可配置多个
IMPORT_SITES = [
    {
        "name": "Site A",
        "url": "https://site-a.com",
        "consumer_key": "ck_a_xxx",
        "consumer_secret": "cs_a_xxx"
    },
    {
        "name": "Site B",
        "url": "https://site-b.com",
        "consumer_key": "ck_b_xxx",
        "consumer_secret": "cs_b_xxx"
    }
]

# CSV 文件输出路径
EXPORT_CSV = "simple_products_final.csv"

# 日志文件
LOG_FILE = "import_export_log.txt"

# 缺失图片记录文件
MISSING_IMAGE_LOG = "missing_images_log.txt"
```

---

## 4. 脚本使用步骤

### 4.1 导出变量商品到 CSV

运行脚本时会自动：

1. 拉取导出站点的所有变量商品
2. 拆分 `pa_smaki` 多值属性
3. 匹配主图和 Gallery
4. 自动生成 CSV，包含以下字段：

   * Name、Type、Regular price、Sale price、SKU
   * Description、Short description
   * Images、Gallery
   * Stock、Stock status、Published
   * Categories、Tags
   * EAN/UPC、Source URL、Button Text
   * 各属性列（Attribute 1 name/value…）

缺失主图的商品会记录到 `missing_images_log.txt`。

---

### 4.2 导入到多个站点

* 自动将 CSV 中每条商品导入到配置的导入站点
* 商品类型为 **外链商品**（External/Affiliate Product）
* `external_url` 指向源站点的商品页面
* 支持批量导入进度条显示：

```text
Importing products:  25%|█████████▌          | 50/200 [00:15<00:45, 3.30product/s]
```

* 已存在 SKU 会自动跳过
* 导入失败会自动重试 3 次，失败后记录日志

---

### 4.3 日志文件说明

* **import_export_log.txt**：导出和导入操作记录
* **missing_images_log.txt**：缺失主图商品记录

---

## 5. CSV 文件字段说明

| 字段                     | 说明                       |
| ---------------------- | ------------------------ |
| Name                   | 商品名称（包含口味）               |
| Type                   | 商品类型（固定 `external`）      |
| Regular price          | 原价                       |
| Sale price             | 促销价（可为空）                 |
| SKU                    | 商品 SKU（自动美化口味）           |
| Description            | 商品描述                     |
| Short description      | 商品短描述                    |
| Images                 | 主图 URL                   |
| Gallery                | 画廊图片 URL（逗号分隔）           |
| Stock                  | 库存数量                     |
| Stock status           | 库存状态（instock/outofstock） |
| Published              | 是否发布（1=已发布）              |
| Categories             | 分类（逗号分隔）                 |
| Tags                   | 标签（逗号分隔）                 |
| EAN/UPC                | SKU 或条码，可为空              |
| Source URL             | 源站点商品 URL（外链商品跳转目标）      |
| Button Text            | 外链商品按钮文本（默认 `"Buy Now"`） |
| Attribute N name/value | 商品属性名称和值（根据站点属性动态生成）     |

---

## 6. 批量导入大商品量建议

* 对大量商品（>1000）建议分批导入
* tqdm 进度条可以实时监控导入进度
* 导入失败自动重试 3 次
* 查看 `import_export_log.txt` 可确认导入状态

---

## 7. 运行示例

```bash
python woocommerce_export_import.py
```

运行后，终端会显示：

1. 导出进度和导出日志
2. 导入进度条和导入日志
3. 最终导入完成提示

CSV 和日志文件保存在脚本所在目录，可直接查看或修改。

---

## 8. 注意事项

1. 导入站点需要开启 REST API 并具有写入权限
2. 若导入站点已有相同 SKU，自动跳过，不会覆盖
3. 缺失主图的商品需要人工补充或确认
4. `pa_smaki` 以外的属性默认单值，仍会导出并保留
