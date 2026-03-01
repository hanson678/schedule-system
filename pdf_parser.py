# -*- coding: utf-8 -*-
"""PO PDF完整解析 v4 - 不漏行 + 取消单检测 + 异常分类"""
import os, re
import pdfplumber


def _normalize_date(s):
    """将各种日期格式统一为YYYY-MM-DD"""
    if not s:
        return ''
    s = str(s).strip().replace('/', '-')
    m = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r'(\d{1,2})-(\d{1,2})-(\d{4})', s)
    if m:
        a, b, year = int(m.group(1)), int(m.group(2)), m.group(3)
        if a > 12:
            return f"{year}-{b:02d}-{a:02d}"
        elif b > 12:
            return f"{year}-{a:02d}-{b:02d}"
        else:
            return f"{year}-{a:02d}-{b:02d}"
    return s


class PDFParser:
    def parse(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ''
            all_tables = []
            for page in pdf.pages:
                full_text += (page.extract_text() or '') + '\n'
                tbls = page.extract_tables()
                if tbls:
                    all_tables.extend(tbls)

        header = self._header(full_text)
        lines = self._lines(all_tables, full_text)
        reqs = self._requirements(full_text)
        is_cancel = self._detect_cancel(full_text)
        return {**header, 'lines': lines, **reqs,
                'is_cancel': is_cancel, 'raw_text': full_text[:8000]}

    def _detect_cancel(self, text):
        """检测取消单：PDF中有取消水印/印章（排除备注中的'取消'）"""
        clean = re.sub(r'(?:Remark|Packaging\s+Info|备注)[：:\s].*', '', text,
                       flags=re.DOTALL | re.I)
        if '取消' in clean or '取 消' in clean:
            return True
        if re.search(r'(?:CANCEL|VOID|CANCELLED)', clean, re.I):
            return True
        return False

    def _header(self, t):
        def f(p, default=''):
            m = re.search(p, t, re.IGNORECASE)
            return m.group(1).strip() if m else default

        # PO号：支持多种格式
        po = (f(r'ZURU\s+Inc\s+PO#[:\s]*(\d{10})') or
              f(r'PO#[:\s]*(4500\d{6})') or
              f(r'Purchase\s+Order[:\s#]*(\d{10})') or
              f(r'PO\s+Number[:\s]*(4\d{9})') or
              f(r'PO[:\s]*#?\s*(4500\d{6})') or
              f(r'Order\s+No\.?[:\s]*(4500\d{6})'))
        dest_raw = (f(r'Destination\s+Country[:\s]*(.+?)(?:\s{2,}|\n)') or
                    f(r'Ship\s+To\s+Country[:\s]*(.+?)(?:\s{2,}|\n)') or
                    f(r'Destination[:\s]*(.+?)(?:\s{2,}|\n)'))
        # 支持YYYY-MM-DD和DD-MM-YYYY/MM-DD-YYYY两种日期格式
        ship_date = (f(r'Shipment\s+Date[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                     f(r'Shipment\s+Date[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})') or
                     f(r'Ship\s+Date[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                     f(r'Ship\s+Date[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})') or
                     f(r'Delivery\s+Date[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                     f(r'Delivery\s+Date[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})') or
                     f(r'ETD[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                     f(r'ETD[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})') or
                     f(r'Required\s+Date[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                     f(r'Required\s+Date[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})'))
        customer = (f(r'Customer\s+Name[:\s]*(.+?)(?:\s{2,}|Payment|Loading|Shipment|\n)') or
                    f(r'Sold\s+To[:\s]*(.+?)(?:\s{2,}|Payment|Loading|Shipment|\n)') or
                    f(r'Bill\s+To[:\s]*(.+?)(?:\s{2,}|Payment|Loading|Shipment|\n)') or
                    f(r'Buyer[:\s]*(.+?)(?:\s{2,}|Payment|Loading|Shipment|\n)'))
        po_date_raw = (f(r'(?<![a-zA-Z])Date[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                       f(r'(?<![a-zA-Z])Date[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})') or
                       f(r'PO\s+Date[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                       f(r'PO\s+Date[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})') or
                       f(r'Order\s+Date[:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})') or
                       f(r'Order\s+Date[:\s]*(\d{1,2}[-/]\d{1,2}[-/]\d{4})'))
        return {
            'po_number': po,
            'po_date': _normalize_date(po_date_raw),
            'customer': customer,
            'customer_po_header': self._clean_cpo_header(
                                   f(r'Customer\s+PO#?[:\s]*(.+?)(?:\s{2,}|Loading|Shipment|Payment|\n)') or
                                   f(r'External\s+Ref[:\s]*(.+?)(?:\s{2,}|\n)')),
            'from_person': self._clean_from(
                            f(r'From[:\s]*(.+?)(?:\s{2,}|\n)') or
                            f(r'Contact[:\s]*(.+?)(?:\s{2,}|\n)') or
                            f(r'Buyer[:\s]*(.+?)(?:\s{2,}|\n)')),
            'ship_date': _normalize_date(ship_date),
            'ship_type': (f(r'Shipment\s+Type[:\s]*(.+?)(?:\s{2,}|\n)') or
                          f(r'Ship\s+Mode[:\s]*(.+?)(?:\s{2,}|\n)') or
                          f(r'Incoterm[:\s]*(.+?)(?:\s{2,}|\n)')),
            'sales_order': f(r'Sales\s+Order#?[:\s]*(\d+)'),
            'destination': dest_raw,
            'destination_cn': self._country(dest_raw),
            'loading_port': (f(r'Loading\s+Port[:\s]*(.+?)(?:\s{2,}|\n)') or
                             f(r'Port\s+of\s+Loading[:\s]*(.+?)(?:\s{2,}|\n)')),
        }

    @staticmethod
    def _clean_cpo_header(s):
        """清理customer_po_header：过滤误取的其他字段值"""
        if not s:
            return ''
        # 如果值看起来像其他字段名（Loading Port, Shipment, Supplier等），说明误取
        if re.match(r'(?:Loading|Shipment|Payment|Ship|Destination|Supplier|From|Att)',
                    s.strip(), re.I):
            return ''
        return s.strip()

    @staticmethod
    def _clean_from(s):
        """清理from_person：只保留人名，去掉后面混入的国家名/地址"""
        if not s:
            return ''
        # 在国家名、字段标签、或常见地名前截断
        s = re.split(r'\s+(?:Australia|China|United|America|France|Germany|Japan|Korea|'
                     r'New\s+Zealand|Singapore|Thailand|Vietnam|Hong\s+Kong|Taiwan|'
                     r'Malaysia|Indonesia|India|Brazil|Canada|Mexico|Italy|Spain|'
                     r'Netherlands|UK|US|USA|EU|Russia|Russian|Turkey|South\s+Africa|'
                     r'Destination|Sales\s+Order|Loading\s+Port)',
                     s, maxsplit=1, flags=re.I)[0].strip()
        # 去掉尾部逗号和多余空格
        return s.rstrip(' ,')

    def _lines(self, tables, full_text):
        lines = []
        last_cm = None      # 上一个成功的列映射
        last_data_start_offset = 0

        for table in tables:
            if not table or len(table) < 1:
                continue

            # 跳过非商品表格（Special Requirements, Additional Clause等）
            first_text = ' '.join([str(c) for c in table[0] if c]).lower() if table[0] else ''
            for r in table[:3]:
                first_text += ' ' + (' '.join([str(c) for c in r if c]).lower() if r else '')
            if any(kw in first_text for kw in ('special requirement', 'additional clause',
                                                 'confirmed and accepted',
                                                 'product requeriments',
                                                 'product requirements')):
                continue

            # 尝试找到表头
            hi, hdr = -1, None
            for i, row in enumerate(table):
                txt = ' '.join([str(c) for c in row if c]).lower()
                if 'line' in txt and ('sku' in txt or 'qty' in txt):
                    hdr = row; hi = i; break

            if hdr is not None:
                # 有表头，构建列映射
                cm = self._build_col_map(hdr)
                data_start = hi + 1
                # 检查子表头行（pcs/size等）
                if data_start < len(table):
                    sub = table[data_start]
                    stxt = ' '.join([str(c) for c in sub if c]).lower() if sub else ''
                    if 'pcs' in stxt or 'size' in stxt or 'version' in stxt:
                        data_start += 1
                        pcs_cols = [j for j, c in enumerate(sub or [])
                                    if c and 'pcs' == str(c).strip().lower()]
                        if len(pcs_cols) >= 2:
                            cm['outer_pcs'] = pcs_cols[1]
                last_cm = cm
            elif last_cm:
                # 没有表头但有上一个表的列映射 → 这是续表！
                # 先尝试自动检测列映射（pdfplumber跨页列数可能不同）
                auto_cm = self._auto_col_map_from_data(table)
                if auto_cm:
                    cm = auto_cm
                    # 如果原表有cpo列但auto_cm没检测到，标记cpo存在以阻止fallback误取qty
                    if 'cpo' in last_cm and 'cpo' not in cm:
                        cm['cpo'] = 9999  # sentinel：cpo列存在但为空
                else:
                    cm = last_cm
                data_start = 0
            else:
                # 既没有表头也没有历史映射，跳过
                continue

            for i in range(data_start, len(table)):
                row = table[i]
                if not row:
                    continue

                # 只跳过真正的合计行：第一个非空单元格是Total/Totals
                first_val = ''
                for c in row:
                    if c and str(c).strip():
                        first_val = str(c).strip().lower()
                        break
                if first_val in ('total', 'totals', 'grand total'):
                    continue

                # 必须包含至少一个3位以上数字
                rtxt = ' '.join([str(c) for c in row if c])
                if not re.search(r'\d{3,}', rtxt):
                    continue

                line = self._extract_line(row, cm)
                # 安全检查：SKU不应超过30字符（超过的肯定不是真实SKU）
                if line and line.get('sku') and len(line['sku']) > 30:
                    continue
                if line and (line['qty'] > 0 or line['sku']):
                    # 去重：同SKU同数量不重复添加
                    dup = False
                    for existing in lines:
                        if existing['sku'] == line['sku'] and existing['qty'] == line['qty']:
                            dup = True; break
                    if not dup:
                        lines.append(line)

        # ===== 兜底：用文本提取验证，防止遗漏 =====
        text_lines = self._extract_lines_from_text(full_text)
        for tl in text_lines:
            found = False
            for el in lines:
                if el['sku'] == tl['sku']:
                    found = True; break
            if not found and (tl['qty'] > 0 or tl['sku']):
                lines.append(tl)

        return lines

    def _auto_col_map_from_data(self, table):
        """从数据行自动检测列映射（用于续表/无表头的表格）
        利用barcode(12-14位纯数字)或日期模式作为锚点，推导其他列位置"""
        # 找第一个有效数据行（跳过Totals、空行等）
        row = None
        for r in table[:5]:
            if not r:
                continue
            first_val = ''
            for c in r:
                if c and str(c).strip():
                    first_val = str(c).strip().lower()
                    break
            if first_val in ('total', 'totals', 'grand total', ''):
                continue
            content = sum(1 for c in r if c and str(c).strip())
            if content >= 5:
                row = r
                break
        if not row:
            return None

        cm = {}
        n = len(row)

        # === 1. 左侧找 line, sku, spec ===
        for j in range(min(5, n)):
            c = row[j]
            if not c:
                continue
            v = re.sub(r'\s+', '', str(c).strip())
            if not v:
                continue
            if 'line' not in cm and re.match(r'^\d{2,3}$', v):
                cm['line'] = j
            elif 'sku' not in cm and re.match(r'^\d{4,}[A-Za-z]*\d*$', v):
                cm['sku'] = j
            elif 'spec' not in cm and re.search(r'\d{4,}[A-Za-z]*\d*-S\d+', v, re.I):
                cm['spec'] = j

        # === 2. 找barcode锚点（12-14位纯数字）===
        barcode_idx = None
        for j in range(n):
            c = row[j]
            if c and re.match(r'^\d{12,14}$', str(c).strip()):
                barcode_idx = j
                cm['barcode'] = j
                break

        # === 3. 以barcode为锚点推导后续列 ===
        if barcode_idx is not None:
            idx = barcode_idx + 1
            # delivery: 紧随barcode的日期
            if idx < n and row[idx]:
                v = str(row[idx]).strip()
                if re.search(r'\d+[-/]\d+[-/]\d+', v):
                    cm['delivery'] = idx
                    idx += 1
            # price: 小数（单价）
            if idx < n and row[idx]:
                try:
                    pv = float(str(row[idx]).strip())
                    if 0 < pv < 100:
                        cm['price'] = idx
                        idx += 1
                except:
                    pass
            # qty: 整数（数量）
            if idx < n and row[idx]:
                v = str(row[idx]).strip().replace(',', '')
                if v.isdigit() and int(v) > 0:
                    cm['qty'] = idx

        # === 3b. 没有barcode时用日期模式作为锚点 ===
        if barcode_idx is None:
            for j in range(n):
                c = row[j]
                if c and re.search(r'\d{1,4}[-/]\d{1,2}[-/]\d{1,4}', str(c).strip()):
                    cm['delivery'] = j
                    # delivery之前一列可能是barcode（可能为空），delivery之后是price、qty
                    idx = j + 1
                    if idx < n and row[idx]:
                        try:
                            pv = float(str(row[idx]).strip())
                            if 0 < pv < 100:
                                cm['price'] = idx
                                idx += 1
                        except:
                            pass
                    if idx < n and row[idx]:
                        v = str(row[idx]).strip().replace(',', '')
                        if v.isdigit() and int(v) > 0:
                            cm['qty'] = idx
                    break

        # === 4. 找name（长文本产品描述）===
        for j in range(n):
            c = row[j]
            if not c or j in cm.values():
                continue
            v = str(c).strip()
            if len(v) > 20 and re.search(r'[A-Za-z]', v):
                cm['name'] = j
                break

        # === 5. 找cpo（最后有内容的非状态列）===
        for j in range(n - 1, max(n - 5, -1), -1):
            c = row[j]
            if not c or j in cm.values():
                continue
            v = str(c).strip()
            if not v:
                continue
            # 跳过状态/运输关键词
            if any(kw in v for kw in ('正式', '暂估', 'Shipp', 'LCL', 'FCL', '40F', '20F', 'ing')):
                continue
            # 跳过小数值（如CBM 1.068、单价0.7920等）
            if re.match(r'^\d*\.\d+$', v):
                continue
            cm['cpo'] = j
            break

        # 验证：至少要有sku和qty
        if 'sku' in cm and 'qty' in cm:
            return cm
        return None

    def _build_col_map(self, hdr):
        """从表头行构建列索引映射（支持多种PO格式）"""
        cm = {}
        for j, c in enumerate(hdr):
            if not c:
                continue
            # 标准化：去换行、多空格合并为单空格（解决"Customer\nPO"匹配问题）
            cl = re.sub(r'\s+', ' ', str(c).strip()).lower()
            if cl in ('line', 'line#', 'line no', 'line no.', 'item', 'no.', 'no', '#'):
                if 'line' not in cm: cm['line'] = j
            elif cl in ('sku', 'sku#', 'item code', 'item no', 'item#', 'article',
                         'product code', 'material', 'part no'):
                if 'sku' not in cm: cm['sku'] = j
            elif any(k in cl for k in ('spec', 'description', 'variant')):
                if 'spec' not in cm: cm['spec'] = j
            elif any(k in cl for k in ('name', 'product name', 'item name', 'item description')):
                if 'name' not in cm: cm['name'] = j
            elif 'barcode' in cl or 'ean' in cl or 'upc' in cl:
                if 'barcode' not in cm: cm['barcode'] = j
            elif 'delivery' in cl or 'ship date' in cl or 'del date' in cl:
                if 'delivery' not in cm: cm['delivery'] = j
            elif ('price' in cl or 'unit cost' in cl or 'unit price' in cl) and 'total' not in cl:
                if 'price' not in cm: cm['price'] = j
            elif cl in ('qty', 'quantity', 'order qty', 'pcs', 'qty ordered'):
                if 'qty' not in cm: cm['qty'] = j
            elif 'total' in cl and ('usd' in cl or 'amount' in cl or 'value' in cl):
                if 'total_usd' not in cm: cm['total_usd'] = j
            elif 'total' in cl and ('ctn' in cl or 'carton' in cl):
                if 'total_ctns' not in cm: cm['total_ctns'] = j
            elif any(k in cl for k in ('customer po', 'cust po', 'customer ref',
                                        'external ref', 'buyer ref', 'client po')):
                if 'cpo' not in cm: cm['cpo'] = j
            elif 'ship' in cl and 'type' in cl:
                if 'ship_type' not in cm: cm['ship_type'] = j
            elif 'cbm' in cl and 'total' not in cl:
                if 'cbm' not in cm: cm['cbm'] = j
        return cm

    def _extract_line(self, row, cm):
        """从一行数据提取订单行信息"""
        def g(k, join_char=' '):
            if k in cm and cm[k] < len(row):
                v = row[cm[k]]
                if not v:
                    return ''
                # 清理换行符（PDF单元格常因换行拆分）
                return re.sub(r'\s*\n\s*', join_char, str(v)).strip()
            return ''

        line = {
            'line_no': g('line', ''), 'sku': g('sku', ''), 'sku_spec': g('spec', ''),
            'name': g('name', ' '), 'barcode': g('barcode', ''),
            'delivery': _normalize_date(g('delivery', '')),
        }

        ps = g('price')
        pm = re.search(r'(\d+\.?\d*)', ps)
        line['price'] = float(pm.group(1)) if pm else 0

        qs = g('qty')
        qm = re.search(r'([\d,]+)', qs)
        line['qty'] = int(qm.group(1).replace(',', '')) if qm else 0

        ts = g('total_usd')
        tm = re.search(r'([\d,.]+)', ts)
        line['total_usd'] = float(tm.group(1).replace(',', '')) if tm else 0

        cs = g('total_ctns')
        csm = re.search(r'([\d,]+)', cs)
        line['total_ctns'] = int(csm.group(1).replace(',', '')) if csm else 0

        cpo = g('cpo')
        # 验证cpo：小数值不是客PO（如1.068是CBM误取）
        if cpo and re.match(r'^\d*\.\d+$', cpo):
            cpo = ''
        # 只在列映射中没有cpo列时才做fallback搜索
        # 如果有cpo列但值为空，说明该PO确实没有客户PO，不做fallback（防止取到qty值）
        if not cpo and 'cpo' not in cm:
            # 先从末尾找数字型客户PO（4位以上纯数字）
            for cell in reversed(row):
                if cell:
                    val = str(cell).strip().replace(',', '')
                    if re.match(r'^\d{4,}$', val):
                        cpo = val; break
            # 再找非纯数字的文本PO
            if not cpo:
                for cell in reversed(row):
                    if cell:
                        val = str(cell).strip()
                        if val and not re.match(r'^[\d,.\-/\s]+$', val) and val not in ('正式', '暂估', ''):
                            if len(val) >= 3:
                                cpo = val; break
        line['customer_po'] = cpo

        outer = 0
        op = g('outer_pcs')
        om = re.search(r'(\d+)', op)
        if om:
            outer = int(om.group(1))
        if outer == 0 and line['name']:
            nm = re.search(r'(\d+)\s*PCS/(?:PDQ/)?CTN', line['name'], re.I)
            if nm:
                outer = int(nm.group(1))
        line['outer_qty'] = outer
        line['item_code'] = line['sku']

        return line

    def _extract_lines_from_text(self, text):
        """从纯文本中提取行数据作为兜底（防止表格解析遗漏）"""
        lines = []
        # 匹配模式：行号(10/20/...) + SKU编号 + 数量
        # 典型格式：90 125160D 125160D-S01 S001-BONKERS-... 1,176
        pattern = re.compile(
            r'(?:^|\n)\s*(\d{2,3})\s+'       # line_no
            r'(\d{4,}[A-Z]?)\s*'              # sku (如125160D)
            r'(\d{4,}[A-Z]?-S\d+)\s+'         # sku_spec (如125160D-S01)
            r'(S\d+-[A-Z\-]+.*?)\s+'          # name
            r'.*?'
            r'([\d,]+)\s+'                     # qty
            r'([\d,.]+)\s*$',                  # price or total
            re.MULTILINE
        )
        for m in pattern.finditer(text):
            try:
                qty_str = m.group(5).replace(',', '')
                qty = int(qty_str) if qty_str.isdigit() else 0
                if qty > 0:
                    lines.append({
                        'line_no': m.group(1),
                        'sku': m.group(2),
                        'sku_spec': m.group(3),
                        'name': m.group(4).strip(),
                        'barcode': '', 'delivery': '',
                        'price': 0, 'qty': qty,
                        'total_usd': 0, 'total_ctns': 0,
                        'customer_po': '', 'outer_qty': 0,
                        'item_code': m.group(2),
                    })
            except:
                continue
        return lines

    def _requirements(self, t):
        tracking = ''
        m = re.search(r'日期码格式[：:]\s*(.*?)\s*日期[：:]\s*(.*?)(?:\n)', t)
        if m:
            tracking = f'日期码格式：{m.group(1).strip()} 日期：{m.group(2).strip()}'

        def ext(p, lim=2000):
            m = re.search(p, t, re.DOTALL | re.I)
            return m.group(1).strip()[:lim] if m else ''

        packaging = ext(r'Packaging\s+Info[：:\s]*(.*?)(?=Remark[：:\s]|$)')
        remark = ext(r'Remark[：:\s]*(.*?)(?=Order Modifiable|$)', 3000)

        # 修订记录（Order Modifiable Records）
        revision = ''
        rev_m = re.search(r'Order\s+Modifiable\s+Records\s*(.*?)(?=Special|Additional|Confirmed|$)',
                          t, re.DOTALL | re.I)
        if rev_m:
            entries = []
            for line in rev_m.group(1).strip().split('\n'):
                line = line.strip()
                if not line or ('Revision' in line and '#' in line) or ('Date' in line and 'Comment' in line):
                    continue
                rm = re.match(r'(\d+)\s+(\d{2}-\d{2}-\d{4})\s+(.*)', line)
                if rm:
                    entries.append(f"Rev.{rm.group(1)} ({_normalize_date(rm.group(2))}): {rm.group(3).strip()}")
            if entries:
                revision = '; '.join(entries)

        return {'tracking_code': tracking, 'packaging_info': packaging,
                'remark': remark, 'revision': revision}

    def _country(self, c):
        if not c: return ''
        m = {'usa': '美国', 'us': '美国', 'united states': '美国', 'u.s.a': '美国',
             'france': '法国', 'fr': '法国',
             'germany': '德国', 'de': '德国', 'deutschland': '德国',
             'uk': '英国', 'gb': '英国', 'united kingdom': '英国', 'great britain': '英国',
             'australia': '澳大利亚', 'au': '澳大利亚',
             'canada': '加拿大', 'ca': '加拿大',
             'japan': '日本', 'jp': '日本',
             'netherlands': '荷兰', 'nl': '荷兰', 'holland': '荷兰',
             'spain': '西班牙', 'es': '西班牙',
             'italy': '意大利', 'it': '意大利',
             'slovakia': '斯洛伐克', 'slovakia,slovakia': '斯洛伐克', 'sk': '斯洛伐克',
             'czech republic': '捷克', 'czechia': '捷克', 'cz': '捷克',
             'poland': '波兰', 'pl': '波兰',
             'new zealand': '新西兰', 'nz': '新西兰',
             'south korea': '韩国', 'korea': '韩国', 'kr': '韩国',
             'mexico': '墨西哥', 'mx': '墨西哥',
             'brazil': '巴西', 'br': '巴西',
             'india': '印度', 'in': '印度',
             'south africa': '南非', 'za': '南非',
             'china': '中国', 'cn': '中国', 'hong kong': '香港', 'hk': '香港',
             'taiwan': '台湾', 'tw': '台湾',
             'singapore': '新加坡', 'sg': '新加坡',
             'malaysia': '马来西亚', 'my': '马来西亚',
             'thailand': '泰国', 'th': '泰国',
             'indonesia': '印度尼西亚', 'id': '印度尼西亚',
             'philippines': '菲律宾', 'ph': '菲律宾',
             'vietnam': '越南', 'vn': '越南',
             'sweden': '瑞典', 'se': '瑞典',
             'norway': '挪威', 'no': '挪威',
             'denmark': '丹麦', 'dk': '丹麦',
             'finland': '芬兰', 'fi': '芬兰',
             'belgium': '比利时', 'be': '比利时',
             'austria': '奥地利', 'at': '奥地利',
             'switzerland': '瑞士', 'ch': '瑞士',
             'portugal': '葡萄牙', 'pt': '葡萄牙',
             'greece': '希腊', 'gr': '希腊',
             'ireland': '爱尔兰', 'ie': '爱尔兰',
             'turkey': '土耳其', 'tr': '土耳其',
             'russia': '俄罗斯', 'ru': '俄罗斯', 'russian fed': '俄罗斯', 'russian federation': '俄罗斯',
             'uae': '阿联酋', 'united arab emirates': '阿联酋',
             'saudi arabia': '沙特', 'sa': '沙特',
             'chile': '智利', 'cl': '智利',
             'argentina': '阿根廷', 'ar': '阿根廷',
             'colombia': '哥伦比亚', 'co': '哥伦比亚',
             'peru': '秘鲁', 'pe': '秘鲁',
             'romania': '罗马尼亚', 'ro': '罗马尼亚',
             'hungary': '匈牙利', 'hu': '匈牙利',
             'croatia': '克罗地亚', 'hr': '克罗地亚',
             'slovenia': '斯洛文尼亚', 'si': '斯洛文尼亚',
             'israel': '以色列', 'il': '以色列',
             'egypt': '埃及', 'eg': '埃及',
             }
        cl = c.strip().lower().split(',')[0].strip().rstrip('.')
        result = m.get(cl, c.split(',')[0].strip())
        _cn_simplify = {'俄罗斯联邦': '俄罗斯', '大韩民国': '韩国', '阿拉伯联合酋长国': '阿联酋'}
        return _cn_simplify.get(result, result)

    # =================== 异常分类与验证 ===================

    @staticmethod
    def classify_error(filename, error):
        """将解析错误分类为用户能看懂的提示"""
        ext = os.path.splitext(filename)[1].lower() if filename else ''
        err = str(error).lower()

        if ext and ext not in ('.pdf', '.xlsx', '.xls'):
            return {
                'category': 'unsupported_format',
                'title': '文件格式不支持',
                'icon': 'bi-file-earmark-x',
                'color': 'danger',
                'tip': (f'上传的是 {ext} 格式文件，系统只能处理PDF文件。\n'
                        '如果是图片(jpg/png)，可以让对方发电子版PDF。\n'
                        '如果是Word/Excel，用WPS另存为PDF再上传。')
            }
        if 'password' in err or 'encrypted' in err:
            return {
                'category': 'encrypted',
                'title': 'PDF有密码保护',
                'icon': 'bi-file-earmark-lock',
                'color': 'warning',
                'tip': ('这个PDF设了密码，系统打不开。\n'
                        '解决：联系发件人要无密码版本，\n'
                        '或者用WPS打开后"另存为"一个新PDF。')
            }
        if any(w in err for w in ('corrupt', 'damage', 'invalid', 'eof marker',
                                   'startxref', 'not a pdf', 'no objects')):
            return {
                'category': 'corrupted',
                'title': 'PDF文件损坏',
                'icon': 'bi-file-earmark-excel',
                'color': 'danger',
                'tip': ('文件可能在传输中损坏了。\n'
                        '解决：从邮件重新下载附件，\n'
                        '或者让对方重新发送邮件。')
            }
        return {
            'category': 'parse_failed',
            'title': '解析失败',
            'icon': 'bi-bug',
            'color': 'danger',
            'tip': (f'系统无法从这个PDF中读取数据。\n'
                    '可能不是标准ZURU PO格式。\n'
                    '建议手动打开PDF查看，手动录入。\n'
                    f'技术详情：{error}')
        }

    @staticmethod
    def validate(data, filename=''):
        """验证解析结果，返回问题/警告列表"""
        issues = []
        raw = data.get('raw_text', '')
        lines = data.get('lines', [])

        # 扫描件/图片检测
        if len(raw.strip()) < 50 and not lines:
            issues.append({
                'category': 'scanned_image',
                'title': '疑似扫描件/图片PDF',
                'icon': 'bi-file-earmark-image',
                'color': 'danger',
                'tip': ('这个PDF几乎没有文字，很可能是扫描件或截图。\n'
                        '系统只能识别"电子版"PDF（文字可以复制的那种），\n'
                        '扫描件需要手动录入。建议让对方发电子版PO。')
            })
            return issues

        po = data.get('po_number', '')
        sku_list = ', '.join([ln.get('sku', '?') for ln in lines[:5]]) if lines else ''

        # PO号缺失
        if not po:
            sku_hint = f'\n涉及货号: {sku_list}' if sku_list else ''
            issues.append({
                'category': 'no_po',
                'title': f'未识别到PO号 · {filename}',
                'icon': 'bi-hash',
                'color': 'warning',
                'sku': sku_list,
                'tip': (f'文件 {filename} 找不到PO号（正常是4500开头的10位数字）。{sku_hint}\n'
                        '可能不是ZURU标准PO格式。\n'
                        '请手动打开PDF确认PO号。')
            })

        # 无商品行
        if not lines and len(raw.strip()) >= 50:
            issues.append({
                'category': 'no_lines',
                'title': f'未识别到商品行 · PO {po}' if po else '未识别到商品行',
                'icon': 'bi-list-ul',
                'color': 'danger',
                'tip': (f'PO {po} PDF有文字但找不到商品行（Line/SKU/Qty表格）。\n'
                        '可能这个PDF不是PO订单。\n'
                        '请手动打开确认内容。')
            })

        # 缺出货日期
        if not data.get('ship_date') and lines:
            issues.append({
                'category': 'no_ship_date',
                'title': f"缺少出货日期 · PO {po}" if po else '缺少出货日期',
                'icon': 'bi-calendar-x',
                'color': 'warning',
                'sku': sku_list,
                'tip': (f"PO {po} 没有检测到出货日期(Shipment Date)。\n"
                        f"涉及货号: {sku_list}\n"
                        '出货日期列会留空，请手动补上。')
            })

        # 缺客户名
        if not data.get('customer') and lines:
            issues.append({
                'category': 'no_customer',
                'title': f'缺少客户名 · PO {po}' if po else '缺少客户名',
                'icon': 'bi-person-x',
                'color': 'info',
                'sku': sku_list,
                'tip': (f'PO {po} 没有识别到客户名(Customer Name)。\n'
                        f'涉及货号: {sku_list}\n'
                        'B列会留空，建议手动补充。')
            })

        return issues
