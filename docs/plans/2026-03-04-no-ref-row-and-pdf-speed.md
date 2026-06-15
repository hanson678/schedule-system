# 无参考行写入 + PDF解析加速 — 实施计划

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 让找不到精确参考行的新货号也能自动写入排期，同时用PyMuPDF加速PDF文本提取。

**Architecture:** 三级匹配(exact/prefix/none)决定写入策略。`_search_sku_in_file`返回match_quality字段，`_do_new_com`据此分支处理。PDF解析改用fitz做文本提取、pdfplumber仅做表格提取。

**Tech Stack:** Python 3.12, Flask, openpyxl(只读), WPS COM(写入), pdfplumber(表格), PyMuPDF/fitz(文本)

---

## Task 1: 安装PyMuPDF依赖

**Files:**
- Modify: `requirements.txt`

**Step 1: 安装PyMuPDF**

Run:
```bash
"C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe" -m pip install PyMuPDF
```

**Step 2: 更新requirements.txt**

在 `requirements.txt` 末尾添加:
```
PyMuPDF>=1.24
```

**Step 3: 验证安装**

Run:
```bash
"C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe" -c "import fitz; print(fitz.__version__)"
```
Expected: 版本号输出，无报错

**Step 4: 同步到测试版**

Run:
```bash
cp "C:/Users/Administrator/Desktop/排期系统/requirements.txt" "C:/Users/Administrator/Desktop/排期系统-测试版/requirements.txt"
```

---

## Task 2: PDF解析器 — 混合引擎 + inner_qty提取

**Files:**
- Modify: `pdf_parser.py:1-44` (parse方法) 和 `pdf_parser.py:474-510` (_build_col_map) 和 `pdf_parser.py:512-580` (_extract_line)
- Modify: `excel_po_parser.py:200-216` (列检测) 和 `excel_po_parser.py:265` (outer_qty附近)

**Step 1: 修改pdf_parser.py — parse()方法改用fitz做文本提取**

在文件顶部(第3行`import pdfplumber`之后)添加:
```python
try:
    import fitz as _fitz  # PyMuPDF — 快速文本提取
except ImportError:
    _fitz = None
```

修改 `parse()` 方法(第28-44行)，将文本提取改为fitz:
```python
def parse(self, pdf_path):
    # 文本提取：优先用PyMuPDF（快5-10倍）
    full_text = ''
    if _fitz:
        try:
            doc = _fitz.open(pdf_path)
            for page in doc:
                full_text += (page.get_text() or '') + '\n'
            doc.close()
        except Exception:
            full_text = ''  # fitz失败则fallback到pdfplumber

    # 表格提取：始终用pdfplumber（准确）
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        if not full_text:
            # fitz不可用时fallback
            for page in pdf.pages:
                full_text += (page.extract_text() or '') + '\n'
        for page in pdf.pages:
            tbls = page.extract_tables()
            if tbls:
                all_tables.extend(tbls)

    header = self._header(full_text)
    lines = self._lines(all_tables, full_text)
    lines = self._resolve_mixed_cartons(lines, full_text)
    reqs = self._requirements(full_text)
    is_cancel = self._detect_cancel(full_text)
    return {**header, 'lines': lines, **reqs,
            'is_cancel': is_cancel, 'raw_text': full_text[:8000]}
```

**Step 2: 修改pdf_parser.py — _build_col_map()增加inner_pcs检测**

在 `_build_col_map()` 方法中(第474-510行)，在外箱检测之后(约第501行后)添加内箱检测:
```python
elif 'inner' in cl and ('qty' in cl or 'pcs' in cl or cl == 'inner'):
    if 'inner_pcs' not in cm: cm['inner_pcs'] = j
```

同时在子表头检测(第282-285行)增加inner_pcs识别：
在 `if len(pcs_cols) >= 2:` 块之后添加:
```python
if len(pcs_cols) >= 3:
    cm['inner_pcs'] = pcs_cols[0]  # 第一个pcs列通常是内箱
```

**Step 3: 修改pdf_parser.py — _extract_line()增加inner_qty字段**

在 `_extract_line()` 方法中，在outer_qty提取之后(第577行之后)、`line['item_code']`之前，添加:
```python
inner = 0
ip = g('inner_pcs')
im = re.search(r'(\d+)', ip)
if im:
    inner = int(im.group(1))
line['inner_qty'] = inner
```

同时在兜底文本提取 `_extract_lines_from_text()` 返回的dict中(约第610行)，添加 `'inner_qty': 0`。

**Step 4: 修改excel_po_parser.py — 增加INNER QTY列检测**

在 `excel_po_parser.py` 第214-215行(`elif v in ('OUTER', 'OUTER QTY')`)之后添加:
```python
elif v in ('INNER', 'INNER QTY') or 'INNER QTY' in v:
    col_map['inner_qty'] = i
```

在第265行(`outer_qty = _to_int(get('outer_qty', 0))`)之后添加:
```python
inner_qty = _to_int(get('inner_qty', 0))
```

在返回的dict(约第275行`lines.append({`)中添加 `'inner_qty': inner_qty,`。

**Step 5: 验证**

用一个现有PDF测试解析速度和结果:
```bash
"C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe" -c "
import time, json
from pdf_parser import PDFParser
p = PDFParser()
# 用uploads目录中任意一个PDF测试
import glob
pdfs = glob.glob('uploads/*.pdf')
if pdfs:
    t0 = time.time()
    r = p.parse(pdfs[0])
    t1 = time.time()
    print(f'解析耗时: {t1-t0:.2f}s')
    for ln in r.get('lines', []):
        print(f'  SKU={ln.get(\"sku\")}, inner_qty={ln.get(\"inner_qty\",\"N/A\")}, outer_qty={ln.get(\"outer_qty\")}')
else:
    print('uploads目录无PDF文件')
"
```

**Step 6: 同步到测试版**

```bash
cp "C:/Users/Administrator/Desktop/排期系统/pdf_parser.py" "C:/Users/Administrator/Desktop/排期系统-测试版/pdf_parser.py"
cp "C:/Users/Administrator/Desktop/排期系统/excel_po_parser.py" "C:/Users/Administrator/Desktop/排期系统-测试版/excel_po_parser.py"
```

---

## Task 3: _search_sku_in_file — 返回match_quality字段

**Files:**
- Modify: `excel_handler.py:550-740` (_search_sku_in_file)
- Modify: `excel_handler.py:470-473` (auto_find中no_item处理)

**核心改动思路：**
`_search_sku_in_file` 当前返回 `{'file', 'fname', 'sheet', 'ref', 'cnt', 'mcol'}` 或 None。
需要在返回dict中增加 `match_quality` 字段。

**Step 1: 修改 _search_sku_in_file 返回值 — 增加match_quality**

在 `_search_sku_in_file` 方法中，每个成功返回结果的地方增加 `match_quality` 字段。

(a) 第662-672行，三级匹配选择 ref 后：

当前代码:
```python
ref = ref_spec_named or ref_spec_any or ref_exact_named or ref_exact_any or ref_prefix_named or ref_prefix_any
cnt = cnt_spec or cnt_exact or cnt_prefix
```

在这两行之后、`if ref:`之前，添加match_quality判断:
```python
# 判断匹配质量
if ref_spec_named or ref_spec_any:
    mq = 'exact'
elif ref_exact_named or ref_exact_any:
    mq = 'exact'
elif ref_prefix_named or ref_prefix_any:
    mq = 'prefix'
else:
    mq = 'none'
```

(b) 第670行的 `result = {...}` 中增加 `'match_quality': mq`:
```python
result = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': ref,
          'cnt': cnt, 'mcol': ws.max_column or 30, 'match_quality': mq}
```

(c) 第674-678行，target_sheet无精确匹配但有last_data_row时：
```python
elif target_sheet and sn in matched_target_sheets:
    if last_data_row:
        logging.info(f"[auto_find] {fn}/{sn} 无精确匹配行，使用最后数据行{last_data_row}作参考")
        best = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': last_data_row,
                'cnt': 1, 'mcol': ws.max_column or 30, 'match_quality': 'none'}
```

(d) 非只读重试成功(第722-728行)：根据匹配方式设置match_quality
```python
if ref_retry:
    # 判断匹配质量
    retry_mq = 'exact' if ref_spec2 else ('exact' if ref_retry == ref_spec2 else 'prefix')
    best = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': ref_retry,
            'cnt': cnt_retry, 'mcol': ws2.max_column or 30, 'match_quality': retry_mq}
elif last_row2:
    best = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': last_row2,
            'cnt': 1, 'mcol': ws2.max_column or 30, 'match_quality': 'none'}
```

(e) COM后备搜索结果(第736行)：如果_search_sku_com返回结果，默认match_quality='exact'（COM搜索已有精确匹配逻辑）。在com_result中补充:
```python
if com_result:
    com_result.setdefault('match_quality', 'exact')
    best = com_result
```

(f) MA辅助定位兜底(约第790行)：任何通过MA找到的结果也需要match_quality='none'。

**Step 2: 修改 auto_find 中的 no_item 处理**

当前第470-473行:
```python
if all_matched:
    logging.info(f"[auto_find] SKU '{sku}' 映射文件找到但货号不在排期中 (no_item)")
    return {'no_item': True}
```

改为：不直接返回 `{'no_item': True}`，而是重新搜索获取last_data_row作为兜底参考行:
```python
if all_matched:
    logging.info(f"[auto_find] SKU '{sku}' 映射文件找到但货号不在排期中，尝试获取兜底参考行")
    # 尝试获取第一个匹配文件的最后数据行作为兜底
    for fp, fn in all_matched:
        fallback = self._search_sku_in_file(fp, fn, num, sku_upper, item,
                                             target_sheet=target_sheet, sku_spec=spec)
        if fallback:
            fallback['match_quality'] = 'none'
            return fallback
    # 所有文件都无法获取参考行
    return {'no_item': True}
```

注意：实际上 `_search_sku_in_file` 在 target_sheet 有但 item 不在时已经返回 last_data_row 了（第674-678行）。所以这里的改动关键是：**上面已经调用过了 _search_sku_in_file 且返回了 None**，说明连 last_data_row 都没有。这种情况确实只能跳过。

更好的方案：在第466-469行的循环中，除了记录 `result`，也记录每个文件搜索过程中的last_data_row作为兜底。

修改第465-473行:
```python
fallback_result = None  # 兜底：有sheet但无匹配的文件/sheet信息
for fp, fn in all_matched:
    result = self._search_sku_in_file(fp, fn, num, sku_upper, item,
                                      target_sheet=target_sheet, sku_spec=spec)
    if result:
        if not result.get('match_quality'):
            result['match_quality'] = 'exact'
        return result
    # 记录兜底信息（文件存在但item未找到）
    if not fallback_result:
        fallback_result = self._get_fallback_ref(fp, fn, target_sheet)

if all_matched:
    if fallback_result:
        logging.info(f"[auto_find] SKU '{sku}' 映射文件找到但货号不在排期中，使用兜底参考行")
        return fallback_result
    logging.info(f"[auto_find] SKU '{sku}' 映射文件找到但货号不在排期中且无兜底 (no_item)")
    return {'no_item': True}
```

**Step 3: 新增 _get_fallback_ref 辅助方法**

在 `_search_sku_in_file` 方法之后添加新方法:
```python
def _get_fallback_ref(self, fp, fn, target_sheet=None):
    """当文件存在但item不在排期中时，获取兜底参考行（最后数据行）
    返回: {'file', 'fname', 'sheet', 'ref', 'cnt', 'mcol', 'match_quality': 'none'} 或 None
    """
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception:
        return None
    try:
        sheets = wb.sheetnames
        if target_sheet:
            matched = [sn for sn in sheets if target_sheet in sn or sn in target_sheet]
            if not matched:
                ts_digits = re.match(r'\d+', target_sheet)
                if ts_digits:
                    matched = [sn for sn in sheets if ts_digits.group() in sn
                               and '取消' not in sn and not _is_ma_sheet(sn)]
            if matched:
                sheets = matched
        for sn in sheets:
            if any(k in sn for k in ('取消', '对应', '总', '旧', '样板')) or _is_ma_sheet(sn):
                continue
            ws = wb[sn]
            last_row = None
            for row in ws.iter_rows(min_row=2, max_col=10):
                row_num = getattr(row[0], 'row', None)
                if row_num is None:
                    continue
                if any(ci < len(row) and row[ci].value for ci in range(min(8, len(row)))):
                    last_row = row_num
            if last_row:
                result = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': last_row,
                          'cnt': 0, 'mcol': ws.max_column or 30, 'match_quality': 'none'}
                wb.close()
                return result
    except Exception:
        pass
    try:
        wb.close()
    except:
        pass
    return None
```

**Step 4: 同步到测试版**

```bash
cp "C:/Users/Administrator/Desktop/排期系统/excel_handler.py" "C:/Users/Administrator/Desktop/排期系统-测试版/excel_handler.py"
```

---

## Task 4: app.py — 前端提示改为黄色警告（不再跳过）

**Files:**
- Modify: `app.py:296-317`

**Step 1: 修改app.py中sku_not_found的处理**

当前第296-317行逻辑：`sku_cache.get(sku)` 返回None或 `{'no_item': True}` 时显示红色danger警告。

修改为：当 `sched` 存在且有 `match_quality` 时，根据质量决定提示颜色。

将第296-317行替换为:
```python
for ln in data.get('lines', []):
    sku = ln.get('item_code') or ln.get('sku', '')
    sched = sku_cache.get(sku)
    no_item = bool(sched and sched.get('no_item'))
    mq = sched.get('match_quality', 'exact') if sched else None

    if not sched:
        # 完全找不到排期文件
        issue = {
            'category': 'sku_not_found',
            'title': f'找不到排期文件 · {sku} · PO {po}',
            'icon': 'bi-exclamation-triangle',
            'color': 'danger',
            'filename': filename,
            'sku': sku,
            'time': now_str,
            'tip': (f'PO {po} 的货号 "{sku}" 在Z盘所有排期文件中未找到。\n'
                    '请检查货号是否正确，或手动到对应排期文件中录入。')
        }
        all_issues.append(issue)
    elif no_item:
        # 找到排期文件但货号不在其中，且无兜底参考行
        issue = {
            'category': 'sku_not_found',
            'title': f'无相同货号(item)，无法写入 · {sku} · PO {po}',
            'icon': 'bi-exclamation-triangle',
            'color': 'danger',
            'filename': filename,
            'sku': sku,
            'time': now_str,
            'tip': (f'PO {po} 的货号 "{sku}" 在排期文件中未找到参考行且无兜底行。\n'
                    '此行跳过不写入，请手动录入。')
        }
        all_issues.append(issue)
    elif mq == 'prefix':
        # 前缀匹配（非精确）
        issue = {
            'category': 'sku_prefix_match',
            'title': f'前缀匹配写入 · {sku} · PO {po}',
            'icon': 'bi-info-circle',
            'color': 'warning',
            'filename': filename,
            'sku': sku,
            'time': now_str,
            'tip': (f'PO {po} 的货号 "{sku}" 无精确匹配，使用前缀参考行写入。\n'
                    '中文名留空，内箱/外箱从PDF提取，请手动检查。')
        }
        all_issues.append(issue)
    elif mq == 'none':
        # 完全无匹配但有兜底参考行
        issue = {
            'category': 'sku_no_match',
            'title': f'无匹配参考行，固定公式写入 · {sku} · PO {po}',
            'icon': 'bi-info-circle',
            'color': 'warning',
            'filename': filename,
            'sku': sku,
            'time': now_str,
            'tip': (f'PO {po} 的货号 "{sku}" 在排期中无匹配货号。\n'
                    '已使用固定公式写入，中文名留空，请手动补充。')
        }
        all_issues.append(issue)
    # mq == 'exact' → 不生成issue（正常情况）

    actions.append({
        'type': 'new', 'line': ln, 'schedule': sched if not no_item else None,
        'sku': ln.get('sku', ''),
        'detail': f"新增 {ln.get('sku','')} {ln.get('qty',0)}pcs"
    })
```

注意：`schedule` 字段现在当 `match_quality` 为 'prefix' 或 'none' 时也会有值（有file/sheet/ref），这样 `batch_process` 就不会跳过这些action了。

**Step 2: 同步到测试版**

```bash
cp "C:/Users/Administrator/Desktop/排期系统/app.py" "C:/Users/Administrator/Desktop/排期系统-测试版/app.py"
```

---

## Task 5: _do_new_com — 根据match_quality分支写入

**Files:**
- Modify: `excel_handler.py:1896-2170` (_do_new_com方法)

**核心改动：**
`_do_new_com` 方法签名不变，但需要从 `act['schedule']` 中获取 `match_quality`。问题是 `_do_new_com` 当前不接收schedule参数，只接收 `ref_row`。

**方案A：** 在 `_do_new_com` 方法签名中新增 `match_quality='exact'` 参数。
调用处(batch_process第1684行)传入:
```python
mq = act['schedule'].get('match_quality', 'exact')
pos, w = self._do_new_com(ws, adj_ref, mc, ops['header'], act['line'],
                           start_after=last_insert_pos.get(sn, 0),
                           match_quality=mq)
```

**Step 1: 修改 _do_new_com 方法签名**

第1896行:
```python
def _do_new_com(self, ws, ref_row, max_col, header, ln, start_after=0, match_quality='exact'):
```

**Step 2: 修改第7步（逐列复制值和公式）**

当前第1940-1952行逐列复制所有值和公式。对于 prefix/none 模式，只复制公式，不复制产品属性值。

将第1940-1952行替换为:
```python
# 7. 逐列复制值和公式
# match_quality='exact': 复制所有内容（现有逻辑）
# match_quality='prefix': 只复制公式，不复制产品属性值（中文名、内箱、外箱）
# match_quality='none': 不复制值（只用格式），公式由后续步骤处理
skip_value_cols = set()  # prefix/none模式下不复制值的列
if match_quality != 'exact':
    for skip_key in ('product_name', 'inner_box', 'outer_box'):
        sc = dcols.get(skip_key)
        if sc:
            skip_value_cols.add(sc)

for c in range(1, mc + 1):
    try:
        ref_cell = ws.Cells(actual_ref, c)
        if ref_cell.HasFormula:
            if match_quality != 'none':
                # exact/prefix: 复制参考行公式
                ws.Cells(pos, c).FormulaR1C1 = ref_cell.FormulaR1C1
            # none模式：公式由7.5步处理
        else:
            if c not in skip_value_cols:
                v = ref_cell.Value
                if v is not None:
                    ws.Cells(pos, c).Value = v
            # skip_value_cols中的列不复制值（留空或由后续PDF数据覆盖）
    except:
        pass
```

**Step 3: 修改第7.5步（公式修复 + none模式固定公式）**

将第1954-1971行替换为扩展版本:
```python
# 7.5 公式处理
# exact/prefix: 修复缺失公式（现有逻辑：搜附近行）
# none: 写固定公式（总箱=数量/外箱，金额根据SLT/SLD/SLB/SK判断）
sku_spec_val = (ln.get('sku_spec', '') or ln.get('sku', '')).upper()
is_slt_type = any(tag in sku_spec_val for tag in ('SLT', 'SLD', 'SLB', 'SK'))

if match_quality == 'none':
    # 总箱数 = 数量 / 外箱数
    tb_col = dcols.get('total_box')
    qty_col = dcols.get('qty')
    ob_col = dcols.get('outer_box')
    if tb_col and qty_col and ob_col:
        try:
            # R1C1公式：=RC[qty_col相对偏移]/RC[ob_col相对偏移]
            # 用绝对列号更可靠
            ws.Cells(pos, tb_col).FormulaR1C1 = f"=RC{qty_col}/RC{ob_col}"
        except Exception as e:
            logging.warning(f"[固定公式] 总箱数公式写入失败: {e}")

    # 卡板：搜附近行公式（仅125160排期有此列）
    plt_col = dcols.get('pallets')
    if plt_col:
        try:
            for sr in range(pos - 1, max(3, pos - 50), -1):
                if ws.Cells(sr, plt_col).HasFormula:
                    ws.Cells(pos, plt_col).FormulaR1C1 = ws.Cells(sr, plt_col).FormulaR1C1
                    break
        except:
            pass

    # 金额
    usd_col = dcols.get('total_usd')
    price_col = dcols.get('price')
    if usd_col and price_col:
        try:
            if is_slt_type and tb_col:
                # SLT/SLD/SLB/SK: 金额 = 总箱数 × 单价
                ws.Cells(pos, usd_col).FormulaR1C1 = f"=RC{tb_col}*RC{price_col}"
            elif qty_col:
                # 普通: 金额 = 数量 × 单价
                ws.Cells(pos, usd_col).FormulaR1C1 = f"=RC{qty_col}*RC{price_col}"
        except Exception as e:
            logging.warning(f"[固定公式] 金额公式写入失败: {e}")
else:
    # exact/prefix模式：原有公式修复逻辑
    calc_col_keys = ['total_box', 'pallets', 'total_usd']
    for ck in calc_col_keys:
        fc = dcols.get(ck)
        if not fc:
            continue
        try:
            if not ws.Cells(pos, fc).HasFormula:
                for sr in range(pos - 1, max(3, pos - 50), -1):
                    try:
                        if ws.Cells(sr, fc).HasFormula:
                            ws.Cells(pos, fc).FormulaR1C1 = ws.Cells(sr, fc).FormulaR1C1
                            break
                    except:
                        pass
        except:
            pass
```

**Step 4: 修改第8步 — prefix/none模式下写入内箱/外箱**

在第2091-2093行的注释之后（不覆写产品属性的注释），添加 prefix/none 模式的内箱/外箱写入:
```python
# prefix/none模式：中文名留空，内箱/外箱从PDF写入
if match_quality != 'exact':
    # 清空中文名（不保留参考行的值）
    pn_col = dcols.get('product_name')
    if pn_col:
        try:
            if not ws.Cells(pos, pn_col).HasFormula:
                ws.Cells(pos, pn_col).ClearContents()
        except:
            pass
    # 内箱从PDF
    ib_col = dcols.get('inner_box')
    inner_val = ln.get('inner_qty', 0)
    if ib_col and inner_val:
        _sv_com(ws, pos, ib_col, inner_val)
    # 外箱从PDF
    ob_col = dcols.get('outer_box')
    outer_val = ln.get('outer_qty', 0)
    if ob_col and outer_val:
        _sv_com(ws, pos, ob_col, outer_val)
```

**Step 5: 修改 batch_process 中的调用**

第1684行，传入 match_quality:
```python
mq = act['schedule'].get('match_quality', 'exact')
pos, w = self._do_new_com(ws, adj_ref, mc, ops['header'], act['line'],
                           start_after=last_insert_pos.get(sn, 0),
                           match_quality=mq)
```

**Step 6: 同步到测试版**

```bash
cp "C:/Users/Administrator/Desktop/排期系统/excel_handler.py" "C:/Users/Administrator/Desktop/排期系统-测试版/excel_handler.py"
```

---

## Task 6: 集成测试

**Step 1: 启动测试版系统**

```bash
cd "C:/Users/Administrator/Desktop/排期系统-测试版" && "C:\Users\Administrator\AppData\Local\Programs\Python\Python312\python.exe" app.py
```

访问 http://localhost:5001

**Step 2: 测试场景A — exact匹配（回归）**

上传一个已知货号的PDF（排期中存在相同货号），验证:
- 写入行为与之前完全一致
- 中文名/内箱/外箱从参考行复制
- 公式正确

**Step 3: 测试场景B — prefix匹配**

上传一个PDF，其货号在排期中有同前缀的（如9296-S003，排期中有9296-S001），验证:
- 前端显示黄色警告"前缀匹配写入"
- 中文名留空
- 内箱/外箱从PDF提取
- 公式从前缀参考行复制

**Step 4: 测试场景C — none匹配**

上传一个完全新货号的PDF，验证:
- 前端显示黄色警告"无匹配参考行，固定公式写入"
- 中文名留空
- 总箱数公式 = 数量/外箱
- 金额公式：普通货号=数量×单价，含SLT/SLD/SLB/SK=总箱×单价

**Step 5: 测试PDF解析速度**

对比优化前后的解析时间（fitz vs pdfplumber纯文本提取）

**Step 6: 确认无误后同步正式版**

确认所有测试通过后:
```bash
cp "C:/Users/Administrator/Desktop/排期系统-测试版/pdf_parser.py" "C:/Users/Administrator/Desktop/排期系统/pdf_parser.py"
cp "C:/Users/Administrator/Desktop/排期系统-测试版/excel_po_parser.py" "C:/Users/Administrator/Desktop/排期系统/excel_po_parser.py"
cp "C:/Users/Administrator/Desktop/排期系统-测试版/excel_handler.py" "C:/Users/Administrator/Desktop/排期系统/excel_handler.py"
cp "C:/Users/Administrator/Desktop/排期系统-测试版/app.py" "C:/Users/Administrator/Desktop/排期系统/app.py"
```
