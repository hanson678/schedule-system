# -*- coding: utf-8 -*-
"""PO PDF完整解析 v4 - 不漏行 + 取消单检测 + 异常分类"""
import os, re, logging
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
            text_parts = []
            all_tables = []
            for page in pdf.pages:
                text_parts.append(page.extract_text() or '')
                tbls = page.extract_tables()
                if tbls:
                    all_tables.extend(tbls)
            full_text = '\n'.join(text_parts)

        header = self._header(full_text)
        lines = self._lines(all_tables, full_text)
        lines = self._merge_cross_page_lines(lines)  # 跨页断裂行合并
        lines = self._resolve_pallet_groups(lines)  # 卡板货号合并（SLB/SLD/SLT/SK），必须在混装处理前
        lines = self._resolve_mixed_cartons(lines, full_text)  # 混装箱处理
        mixed_groups = getattr(self, '_mixed_groups_info', [])
        # 填充PO号到mixed_groups
        po = header.get('po_number', '')
        for mg in mixed_groups:
            mg['po_number'] = po
        reqs = self._requirements(full_text)
        is_cancel = self._detect_cancel(full_text)
        return {**header, 'lines': lines, **reqs,
                'is_cancel': is_cancel, 'raw_text': full_text[:8000],
                'mixed_groups': mixed_groups}

    def _detect_cancel(self, text):
        """检测取消单：PDF中有取消水印/印章（排除备注/修订记录中的'取消'）"""
        # 排除备注、包装信息、修订记录段落（修改单的修订历史常含CANCELLED字样）
        clean = re.sub(r'(?:Remark|Packaging\s+Info|备注|Order\s+Modif|Modifiable\s+Records?|Revision|Change\s+Log)[：:\s].*', '', text,
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
                     r'Guatemala|Uruguay|Costa\s+Rica|Panama|Dominican|Puerto\s+Rico|'
                     r'Ecuador|Venezuela|Paraguay|Bolivia|Honduras|El\s+Salvador|'
                     r'Nicaragua|Jamaica|Colombia|Peru|Argentina|Chile|'
                     r'Sweden|Norway|Denmark|Finland|Belgium|Austria|Switzerland|'
                     r'Portugal|Greece|Ireland|Israel|Egypt|Morocco|Kenya|Nigeria|'
                     r'Czech|Poland|Romania|Hungary|Croatia|Slovenia|Slovakia|'
                     r'Bulgaria|Serbia|Estonia|Latvia|Lithuania|Iceland|Luxembourg|'
                     r'Destination|Sales\s+Order|Loading\s+Port)',
                     s, maxsplit=1, flags=re.I)[0].strip()
        # 去掉尾部逗号和多余空格
        return s.rstrip(' ,')

    # =================== 混装箱处理 ===================

    @staticmethod
    def _is_std_product(line):
        """判断是否是STD PRODUCT / DUMMY子行（混装箱散件）
        识别模式：含STD或DUMMY的SPEC、或price=0且SPEC含PRODUCT/PRODCC等变体"""
        spec = (line.get('sku_spec') or '').upper().replace('\n', ' ')
        name = (line.get('name') or '').upper()
        if 'STD' in spec:
            return True
        # price=0 + SPEC含PRODUCT/PRODCC等变体（如77772-PRODCC1-PRODUCT）
        if 'DUMMY' in spec:
            return True
        if line.get('price', 1) == 0 and line.get('total_usd', 1) == 0:
            if 'PRODUCT' in spec or 'PRODCC' in spec:
                return True
            if 'STD PRODUCT' in name or 'BULK' in name:
                return True
        return False

    @staticmethod
    def _extract_std_product_qtys(full_text):
        """从PDF文本提取STD PRODUCT子行的数量。
        返回: {line_no_str: [{'sku': sku, 'qty': qty}, ...]}
        典型格式：10 7149 7149-STD[\\n]PRODUCT ... 0.0000 360.000 0.00 0 0.000
        """
        result = {}
        pattern = re.compile(
            r'(?:^|\n)[ \t]*(\d{2,3})[ \t]+'      # 行号 (group 1)
            r'(\d{4,}[A-Za-z]?)[ \t]+'              # SKU  (group 2)
            r'\d{4,}[A-Za-z]?-STD'                  # spec 开头含 -STD
            r'[\s\S]{0,500}?'                        # 中间内容（非贪婪）
            r'0\.\d{3,4}[ \t]+'                     # CBM = 0.xxxx
            r'([\d,]{1,10})(?:\.\d+)?[ \t]+'          # qty  (group 3) 支持千分位逗号
            r'0\.00[ \t]+'                           # Total USD = 0.00
            r'0[ \t]+'                               # Total CTNS = 0
            r'0\.\d',                                # Total CBM = 0.x
            re.MULTILINE
        )
        for m in pattern.finditer(full_text):
            line_no = m.group(1)
            sku = m.group(2)
            try:
                qty = int(float(m.group(3).replace(',', '')))
                if qty > 0:
                    result.setdefault(line_no, []).append({'sku': sku, 'qty': qty})
            except Exception:
                pass
        return result

    def _resolve_mixed_cartons(self, lines, full_text):
        """处理混装箱和同line多行:
        1. MEC类（字母前缀+有价格+STD组件）→ 拆分为各组件独立行
        2. 通用同line多行（同line_no有price=0子行）→ 子行继承父行字段，各自独立
        所有混装组标记needs_user_confirmation，由前端弹窗确认
        """
        # Step 1: 从已解析行提取STD子行数量（table解析已含正确qty）
        std_by_lineno = {}
        zero_qty_stds = []  # qty=0的STD行，需文本兜底
        for line in lines:
            if self._is_std_product(line):
                ln = str(line.get('line_no', '')).strip()
                sk = line.get('sku', '') or line.get('item_code', '')
                qt = line.get('qty', 0)
                if ln and sk and qt > 0:
                    std_by_lineno.setdefault(ln, []).append({'sku': sk, 'qty': qt})
                elif ln and sk and qt == 0:
                    zero_qty_stds.append((ln, sk))

        # Step 1.5: 对table中qty=0的STD行，从文本中补查qty
        for zln, zsk in zero_qty_stds:
            existing = {c['sku'] for c in std_by_lineno.get(zln, [])}
            if zsk in existing:
                continue
            pat = re.compile(
                r'(?:^|\n)\s*' + re.escape(zln) + r'\s+' + re.escape(zsk) +
                r'\s.*?0\.0{3,4}\s+([\d,]+)(?:\.\d+)?\s+0\.00\s+0\s+0\.\d',
                re.DOTALL
            )
            m = pat.search(full_text)
            if m:
                try:
                    qt = int(float(m.group(1).replace(',', '')))
                    if qt > 0:
                        std_by_lineno.setdefault(zln, []).append({'sku': zsk, 'qty': qt})
                except:
                    pass

        # Step 2: 正则提取作为补充（防止table解析漏掉）
        regex_std = self._extract_std_product_qtys(full_text)
        for ln, comps in regex_std.items():
            existing_skus = {c['sku'] for c in std_by_lineno.get(ln, [])}
            for c in comps:
                if c['sku'] not in existing_skus:
                    std_by_lineno.setdefault(ln, []).append(c)

        # Step 2.5: 检测同line_no的非STD price=0子行（通用同line多行）
        from collections import defaultdict
        non_std_by_lineno = defaultdict(list)
        for i, line in enumerate(lines):
            if self._is_std_product(line):
                continue
            ln = str(line.get('line_no', '')).strip()
            if ln:
                non_std_by_lineno[ln].append((i, line))

        general_mixed = {}   # line_no → {'parent_idx': int, 'children': [(idx, line)]}
        general_skip = set()  # 子行原始索引，主循环中跳过
        for ln, entries in non_std_by_lineno.items():
            if len(entries) <= 1:
                continue
            parents = [(i, l) for i, l in entries if (l.get('price', 0) or 0) > 0]
            children = [(i, l) for i, l in entries if (l.get('price', 0) or 0) == 0]
            if not parents or not children:
                continue
            # 排除MEC父行（字母前缀在MEC分支中处理）
            parent_sku = parents[0][1].get('sku', '')
            parent_base = parent_sku.split('-')[0]
            if re.match(r'[A-Za-z]', parent_base):
                continue
            general_mixed[ln] = {'parent_idx': parents[0][0], 'children': children}
            for ci, cl in children:
                general_skip.add(ci)

        # Step 3: 构建结果
        result = []
        mixed_groups_info = []

        for i, line in enumerate(lines):
            if i in general_skip:
                continue
            if self._is_std_product(line):
                continue

            line_no = str(line.get('line_no', '')).strip()
            sku = line.get('sku', '') or ''
            sku_spec = line.get('sku_spec', '') or ''
            price = line.get('price', 0) or 0
            base = sku.split('-')[0]
            has_letter_prefix = bool(re.match(r'[A-Za-z]', base))

            if has_letter_prefix and price > 0 and line_no in std_by_lineno:
                # === MEC父行（有价格+有STD组件）→ 拆分为各组件独立行 ===
                components = std_by_lineno[line_no]
                carton_count = line.get('qty', 0)
                # 从sku_spec提取base之后的全部后缀（含年份+S编号）
                # 例: MB212-2025-S001 → after_base="-2025-S001"
                #     MEC123-S001     → after_base="-S001"
                if sku_spec.upper().startswith(base.upper() + '-'):
                    after_base = sku_spec[len(base):]
                else:
                    after_base = ''
                    spec_parts = sku_spec.split('-')
                    for p in spec_parts[1:]:
                        if re.match(r'S\d+', p, re.I):
                            after_base = f'-{p}'
                            break

                group_comps = []
                for comp in components:
                    new_line = dict(line)
                    comp_sku = f"{base}-{comp['sku']}{after_base}"
                    new_line['sku'] = comp_sku
                    new_line['sku_spec'] = comp_sku
                    new_line['item_code'] = comp['sku']
                    new_line['qty'] = comp['qty']
                    new_line['carton_count'] = carton_count
                    new_line['is_mixed_carton'] = True
                    new_line['mixed_parent_sku'] = sku
                    new_line['mixed_components'] = components
                    new_line['mixed_qty_original'] = carton_count
                    new_line['needs_user_confirmation'] = True
                    result.append(new_line)
                    group_comps.append({'sku': comp_sku, 'qty': comp['qty']})

                mixed_groups_info.append({
                    'line_no': line_no,
                    'parent_sku': sku_spec,
                    'type': 'mec_split',
                    'components': group_comps,
                    'customer_po': line.get('customer_po', ''),
                })

            elif has_letter_prefix and price > 0 and line_no not in std_by_lineno:
                line['mec_split_failed'] = True
                line['mec_fail_reason'] = f'{sku}未找到组件子行，需手动入单'
                result.append(line)

            elif has_letter_prefix and price == 0:
                continue

            else:
                # === 通用处理 ===
                has_std = line_no in std_by_lineno
                has_general = line_no in general_mixed

                if has_std or has_general:
                    # 同line多行：父行保留 + 子行各自独立（继承父行字段）
                    line['is_mixed_carton'] = True
                    line['needs_user_confirmation'] = True
                    result.append(line)

                    group_comps = []

                    # STD组件 → 各自独立行（继承父行的price/delivery/inner_pcs等）
                    if has_std:
                        for comp in std_by_lineno[line_no]:
                            new_line = dict(line)
                            comp_sku = comp['sku']
                            comp_sku = re.sub(r'-?STD\s*PRODUCT$', '', comp_sku, flags=re.I).strip()
                            comp_sku = re.sub(r'-?STDPRODUCT$', '', comp_sku, flags=re.I).strip()
                            new_line['sku'] = comp_sku
                            new_line['sku_spec'] = comp_sku
                            new_line['item_code'] = comp_sku.split('-')[0] if '-' in comp_sku else comp_sku
                            new_line['qty'] = comp['qty']
                            new_line['is_mixed_carton'] = True
                            new_line['needs_user_confirmation'] = True
                            new_line['mixed_parent_sku'] = sku_spec
                            result.append(new_line)
                            group_comps.append({'sku': comp_sku, 'qty': comp['qty']})

                    # 非STD的price=0子行 → 各自独立行（继承父行的price/delivery等）
                    if has_general:
                        for ci, child in general_mixed[line_no]['children']:
                            new_line = dict(line)
                            child_sku = child.get('sku_spec') or child.get('sku', '')
                            child_sku = re.sub(r'-?STD\s*PRODUCT$', '', child_sku, flags=re.I).strip()
                            child_sku = re.sub(r'-?STDPRODUCT$', '', child_sku, flags=re.I).strip()
                            new_line['sku'] = child_sku
                            new_line['sku_spec'] = child_sku
                            new_line['item_code'] = child_sku.split('-')[0] if '-' in child_sku else child_sku
                            new_line['qty'] = child.get('qty', 0)
                            if child.get('name'):
                                new_line['name'] = child['name']
                            if child.get('barcode'):
                                new_line['barcode'] = child['barcode']
                            new_line['is_mixed_carton'] = True
                            new_line['needs_user_confirmation'] = True
                            new_line['mixed_parent_sku'] = sku_spec
                            result.append(new_line)
                            group_comps.append({'sku': child_sku, 'qty': child.get('qty', 0)})

                    if group_comps:
                        mixed_groups_info.append({
                            'line_no': line_no,
                            'parent_sku': sku_spec,
                            'type': 'same_line_split',
                            'components': group_comps,
                            'customer_po': line.get('customer_po', ''),
                        })
                else:
                    result.append(line)

        self._mixed_groups_info = mixed_groups_info
        return result

    # 卡板货号（SLB/SLD/SLT/SK/MTQ）识别正则
    _PALLET_MAIN_RE = re.compile(r'^(?:\d+(?:SLB|SLD|SLT|SK)\d*|MTQ\d+)(?:-(?!P\d)|$)', re.I)
    _PALLET_PART_RE = re.compile(r'^(?:\d+(?:SLB|SLD|SLT|SK)\d*|MTQ\d+)-P\d', re.I)

    def _resolve_pallet_groups(self, lines):
        """卡板货号合并：同Line下 MAIN + Px零件 + PRODUCT 合并为一条
        规则：
        - MAIN行（如15752SLB-S002）：提供货号(sku_spec)和单价(price)
        - Px行（如15752SLB-P2）：零件，丢弃不入排期
        - PRODUCT行（如15752-STD PRODUCT）：提供产品件数
        - PO数量 = 同Line所有子行中QTY最大的（即产品件数）
        - 单价 = MAIN行Price（每卡板价格）
        """
        import logging
        if not lines:
            return lines

        # 按line_no分组
        from collections import OrderedDict
        groups = OrderedDict()
        for ln in lines:
            lno = str(ln.get('line_no', '')).strip()
            groups.setdefault(lno, []).append(ln)

        result = []
        for lno, group in groups.items():
            if len(group) == 1:
                result.append(group[0])
                continue

            # 找MAIN行（SKU含SLB/SLD/SLT/SK且非零件）
            main = None
            for ln in group:
                sku = ln.get('sku_spec', '') or ln.get('sku', '') or ''
                if self._PALLET_MAIN_RE.match(sku) and not self._PALLET_PART_RE.match(sku):
                    main = ln
                    break

            if not main:
                # 非卡板组，保留所有行
                result.extend(group)
                continue

            # 取所有子行中最大的QTY作为产品件数
            max_qty = 0
            for ln in group:
                q = ln.get('qty', 0) or 0
                if q > max_qty:
                    max_qty = q

            # 合并：用MAIN行数据，替换QTY为产品件数
            merged = dict(main)
            merged['qty'] = max_qty
            # 标记为卡板货号（供后续金额公式处理）
            merged['is_pallet'] = True
            merged['pallet_count'] = main.get('qty', 0)  # 原始卡板数
            logging.info(f"[卡板合并] Line {lno}: {main.get('sku_spec','')} "
                         f"卡板数={main.get('qty',0)} 产品件数={max_qty} 单价={main.get('price',0)}")
            result.append(merged)

        return result

    def _lines(self, tables, full_text):
        lines = []
        last_cm = None      # 上一个成功的列映射
        last_data_start_offset = 0

        for table in tables:
            if not table or len(table) < 1:
                continue

            # 跳过非商品表格（Special Requirements, Additional Clause, Order Modifiable Records等）
            first_text = ' '.join([str(c) for c in table[0] if c]).lower() if table[0] else ''
            for r in table[:3]:
                first_text += ' ' + (' '.join([str(c) for c in r if c]).lower() if r else '')
            if (any(kw in first_text for kw in ('special requirement', 'additional clause',
                                                  'confirmed and accepted',
                                                  'product requeriments',
                                                  'product requirements',
                                                  'order modifiable', 'modifiable records'))
                    or ('revision' in first_text and 'comment' in first_text)):
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
                        if len(pcs_cols) >= 1:
                            cm['inner_pcs'] = pcs_cols[0]
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
                # SKU必须含数字才是有效货号（过滤中文备注/日期被误取为SKU的情况）
                if line and line.get('sku') and not re.search(r'\d', line['sku']):
                    continue
                if line and (line['qty'] > 0 or line['sku']):
                    # 去重：同SKU同数量不重复添加（但line_no不同的是不同行）
                    dup = False
                    for existing in lines:
                        if existing['sku'] == line['sku'] and existing['qty'] == line['qty']:
                            # 如果两行都有line_no且不同，说明是不同订单行，不去重
                            if (line.get('line_no') and existing.get('line_no')
                                    and line['line_no'] != existing['line_no']):
                                continue
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

    # ---------- 跨页断裂行合并 ----------
    @staticmethod
    def _has_valid_item(sku):
        """检查SKU是否包含有效货号模式（4位以上连续数字，如92129H、15726A）"""
        return bool(sku and re.search(r'\d{4,}', sku))

    def _merge_cross_page_lines(self, lines):
        """合并跨页断裂的行：数据行(有qty无有效货号) + 头部行(有货号无qty)
        PDF跨页时同一Line可能被拆成两段：
        - 页末：有数量/价格/日期/barcode，但货号缺失或不完整
        - 页首：有货号/产品名，但数量=0
        """
        if len(lines) < 2:
            return lines

        result = []
        skip = set()

        for i in range(len(lines)):
            if i in skip:
                continue
            cur = lines[i]
            nxt = lines[i + 1] if i + 1 < len(lines) and i + 1 not in skip else None

            if nxt:
                cur_has_item = self._has_valid_item(cur.get('sku', ''))
                nxt_has_item = self._has_valid_item(nxt.get('sku', ''))

                # 检测跨页断裂：一行有qty无货号 + 另一行有货号无qty
                if (cur['qty'] > 0) != (nxt['qty'] > 0) and cur_has_item != nxt_has_item:
                    data_line = cur if cur['qty'] > 0 else nxt
                    item_line = nxt if cur['qty'] > 0 else cur
                    merged = dict(data_line)
                    # 从货号行取标识字段
                    merged['sku'] = item_line['sku']
                    merged['sku_spec'] = item_line.get('sku_spec') or data_line.get('sku_spec', '')
                    merged['name'] = item_line.get('name') or data_line.get('name', '')
                    merged['item_code'] = item_line.get('item_code') or data_line.get('item_code', '')
                    merged['customer_po'] = item_line.get('customer_po') or data_line.get('customer_po', '')
                    if item_line.get('line_no'):
                        merged['line_no'] = item_line['line_no']
                    logging.info('[跨页合并] %s qty=%s', merged['sku'], merged['qty'])
                    result.append(merged)
                    skip.add(i + 1)
                    continue

            result.append(cur)

        return result

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

        # 内箱数
        inner = 0
        ip = g('inner_pcs')
        im = re.search(r'(\d+)', ip)
        if im:
            inner = int(im.group(1))
        if inner == 0 and line['name']:
            # 从名称提取：如"8PCS/PDQ" → inner=8
            inm = re.search(r'(\d+)\s*PCS/PDQ', line['name'], re.I)
            if inm:
                inner = int(inm.group(1))
        line['inner_pcs'] = inner

        # 外箱数
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

        def ext(p):
            m = re.search(p, t, re.DOTALL | re.I)
            return m.group(1).strip() if m else ''

        packaging = ext(r'Packaging\s+Info[：:\s]*(.*?)(?=Remark[：:\s]|$)')
        remark = ext(r'Remark[：:\s]*(.*?)(?=Order Modifiable|$)')

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

    # 优先使用的短名/别名覆盖表（pycountry返回的正式名太长，或无法识别的缩写）
    _COUNTRY_OVERRIDES = {
        'usa': '美国', 'u.s.a': '美国', 'u.s.a.': '美国',
        'uk': '英国', 'great britain': '英国',
        'uae': '阿联酋', 'united arab emirates': '阿联酋',
        'utd.arab emir': '阿联酋', 'utdarab emir': '阿联酋',
        'russia': '俄罗斯', 'russian fed': '俄罗斯', 'russian fed.': '俄罗斯',
        'russian federation': '俄罗斯',
        'korea': '韩国', 'south korea': '韩国',
        'czech republic': '捷克', 'czechia': '捷克',
        'hong kong': '香港', 'taiwan': '台湾',
        'holland': '荷兰', 'deutschland': '德国',
        'saudi arabia': '沙特',
        'ivory coast': '科特迪瓦', "cote d'ivoire": '科特迪瓦',
        'viet nam': '越南',
        'slovak republic': '斯洛伐克',
        # babel返回的正式名→惯用简称
        '阿拉伯联合酋长国': '阿联酋', '俄罗斯联邦': '俄罗斯',
        '大韩民国': '韩国', '朝鲜': '朝鲜',
    }
    # pycountry+babel 翻译缓存（类级别，进程内只初始化一次）
    _babel_locale = None
    _country_cache = {}

    def _country(self, c):
        if not c: return ''
        import re as _re
        raw = _re.sub(r'[\n\r]+', ' ', str(c)).strip()
        # 取逗号前段、去尾点、规范空格
        lookup = raw.split(',')[0].strip().rstrip('.').strip()
        cl = lookup.lower()

        # 1. 覆盖表（优先级最高：处理缩写、别名、babel长名→短名）
        if cl in self._COUNTRY_OVERRIDES:
            return self._COUNTRY_OVERRIDES[cl]
        # 去掉所有点再查一次（处理 U.S.A → usa）
        cl_nodot = cl.replace('.', '').strip()
        if cl_nodot in self._COUNTRY_OVERRIDES:
            return self._COUNTRY_OVERRIDES[cl_nodot]

        # 2. 缓存命中
        if cl in self._country_cache:
            return self._country_cache[cl]

        # 3. pycountry + babel 全量翻译（195个国家全覆盖）
        try:
            import pycountry
            if PDFParser._babel_locale is None:
                from babel import Locale
                PDFParser._babel_locale = Locale.parse('zh_Hans_CN')
            locale = PDFParser._babel_locale

            # 先尝试 ISO alpha-2 直查（两位代码如 US/GB/CN）
            country = None
            if len(cl_nodot) == 2:
                country = pycountry.countries.get(alpha_2=cl_nodot.upper())
            # 再模糊搜索（去点号后）
            if not country:
                results = pycountry.countries.search_fuzzy(cl_nodot)
                country = results[0] if results else None

            if country:
                cn = locale.territories.get(country.alpha_2, '')
                # 再过一遍覆盖表（处理 babel 返回正式长名）
                cn = self._COUNTRY_OVERRIDES.get(cn, cn) or cn
                if cn:
                    PDFParser._country_cache[cl] = cn
                    return cn
        except Exception:
            pass

        # 4. 兜底：返回原始首段（至少不写英文全名进去）
        result = lookup
        PDFParser._country_cache[cl] = result
        return result

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
