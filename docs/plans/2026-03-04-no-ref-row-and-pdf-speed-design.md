# 设计文档：无参考行写入 + PDF解析加速

日期：2026-03-04

## 背景

当前系统在新单录入时，若找不到与目标货号完全相同的参考行，会跳过不写入并显示红色警告。
这导致新货号（首次出现在排期中的SKU）必须手动录入，降低了效率。
同时，PDF解析使用pdfplumber全量处理，速度较慢。

## 目标

1. 找不到精确参考行时仍然写入，尽量从PDF提取可用数据
2. 加速PDF解析

---

## 一、三级匹配写入策略

### 匹配级别

| 级别 | 条件 | 说明 |
|------|------|------|
| **exact** | 货号完全一致（现有逻辑） | 所有产品属性从参考行复制 |
| **prefix** | 同sheet内前缀匹配（如9296匹配9296-S001） | 公式/格式从参考行复制，产品属性从PDF取 |
| **none** | 同sheet内完全找不到匹配 | 写固定公式，产品属性从PDF取 |

### 各级别字段来源

| 字段 | exact | prefix | none |
|------|-------|--------|------|
| 货号(ITEM#) | PDF sku_spec | PDF sku_spec | PDF sku_spec |
| 中文名 | 参考行 | **留空** | **留空** |
| 内箱数 | 参考行 | **PDF inner_qty** | **PDF inner_qty** |
| 外箱数 | 参考行 | **PDF outer_qty** | **PDF outer_qty** |
| 总箱数 | 参考行公式 | 前缀参考行公式 | **固定: =数量列/外箱列** |
| 卡板 | 参考行公式 | 前缀参考行公式 | **搜附近行FormulaR1C1**（仅125160排期有此列） |
| 金额 | 参考行公式 | 前缀参考行公式 | **固定公式（见下方规则）** |
| PO数量 | PDF qty | PDF qty | PDF qty |
| 出货期 | PDF delivery | PDF delivery | PDF delivery |
| 客户PO | PDF customer_po | PDF customer_po | PDF customer_po |
| 走货国家 | PDF destination_cn | PDF destination_cn | PDF destination_cn |
| 接单期 | PDF po_date | PDF po_date | PDF po_date |
| 跟单 | PDF from_person | PDF from_person | PDF from_person |
| 单价 | PDF price | PDF price | PDF price |

### 金额固定公式规则（match_quality=none）

- **普通货号**：金额 = 数量(QTY) × 单价(Price)
- **含SLT/SLD/SLB/SK的货号**：金额 = 总箱数 × 单价

判断逻辑：检查sku_spec是否包含"SLT"、"SLD"、"SLB"、"SK"（大小写不敏感）

---

## 二、PDF解析速度优化 — 混合引擎

### 当前架构

```
pdfplumber.open(pdf) → 遍历每页:
  page.extract_text()    ← 慢
  page.extract_tables()  ← 慢
```

### 优化后架构

```
fitz.open(pdf) → 遍历每页:
  page.get_text()        ← 快（PyMuPDF）

pdfplumber.open(pdf) → 遍历每页:
  page.extract_tables()  ← 保留（准确）
```

- 文本提取用PyMuPDF（快5-10倍）
- 表格提取保留pdfplumber（准确性好）
- 新增依赖：PyMuPDF (`pip install PyMuPDF`)

---

## 三、改动文件清单

### 1. pdf_parser.py

- `parse()`: 文本提取改用fitz，表格提取保留pdfplumber
- `_build_col_map()`: 增加inner_pcs列检测（匹配"inner"/"inner qty"等表头）
- `_extract_line()`: 增加inner_qty字段提取（从inner_pcs列或产品名中解析）
- 新增依赖: import fitz

### 2. excel_po_parser.py

- 列检测增加INNER QTY / INNER列匹配
- 返回dict增加inner_qty字段

### 3. excel_handler.py — _search_sku_in_file()

- 返回值dict增加`match_quality`字段: "exact" / "prefix" / "none"
- `no_item=True`时不再直接返回，改为`match_quality="none"`并继续
- 使用最后数据行作为格式/公式兜底参考

### 4. excel_handler.py — _do_new_com()

- 读取match_quality分支处理
- **exact**: 现有逻辑不变
- **prefix**:
  - 第7步：从前缀参考行复制公式和格式
  - 第8步：写入时不覆写中文名（留空），内箱/外箱从PDF数据写入
- **none**:
  - 第7步：从最后数据行复制行格式
  - 第7.5步扩展：总箱数写固定公式(=数量列/外箱列)，卡板搜附近行，金额根据SLT/SLD/SLB/SK判断
  - 第8步：同prefix模式

### 5. app.py

- `no_item`/`sku_not_found`类别改为黄色提示
- 提示文案改为："无精确匹配货号，已使用[前缀参考/固定公式]写入，请手动检查中文名等字段"
- 不再跳过不写入

### 6. requirements.txt

- 新增: PyMuPDF

---

## 四、风险与注意事项

1. **inner_qty提取率**: 不是所有PDF都有内箱数据，提取不到时为0（留空）
2. **无参考行的行格式**: 用最后数据行的格式做基础，可能不完美但可接受
3. **前端区分提示**: 用户需知道哪些行是无精确匹配写入的，方便后续补充中文名
4. **SLT/SLD/SLB/SK判断**: 检查sku_spec字段，大小写不敏感
5. **PyMuPDF兼容性**: fitz的get_text()输出格式可能与pdfplumber的extract_text()略有差异，需确保_header()和其他文本解析正则仍能正常工作
