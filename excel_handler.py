# -*- coding: utf-8 -*-
"""排期Excel自动处理 v5 - SKU映射 + 缓存 + 模糊搜索 + 并发锁 + 进度条"""
import os, re, shutil, json, threading, time as _time, logging
from datetime import datetime, timedelta, date
from contextlib import contextmanager
import openpyxl
from openpyxl.utils import column_index_from_string

DESKTOP = os.path.join(os.environ.get('USERPROFILE', r'C:\Users\Administrator'), 'Desktop')
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
HISTORY_FILE = os.path.join(DATA_DIR, 'history.json')
RETRY_FILE = os.path.join(DATA_DIR, 'pending_retries.json')
BATCH_DIR = os.path.join(DESKTOP, 'batch_temp')
CFG_FILE = os.path.join(DATA_DIR, 'config.json')

# COM 颜色常量 (RGB: R + G*256 + B*65536)
BLUE_COM = 15773696     # RGB(0, 176, 240) = 0xF0B000 → 浅蓝 FF00B0F0
RED_COM = 255           # RGB(255, 0, 0) → 红色字体
BLACK_COM = 0           # RGB(0, 0, 0) → 黑色字体
YELLOW_COM = 65535      # RGB(255, 255, 0) → 黄色填充
UNDO_DIR = os.path.join(BATCH_DIR, 'undo')
UNDO_HISTORY = os.path.join(DATA_DIR, 'undo_history.json')

# =================== 全局缓存 ===================
_sku_map_cache = {}          # SKU→排期文件关键词映射
_sku_map_mtime = 0           # 总排期文件修改时间
_yellow_cache = {}           # {filepath: {'mtime': float, 'rows': [...]}}
_yellow_cache_time = 0       # 上次全量扫描时间

# =================== 文件操作锁 ===================
_file_locks = {}             # {filepath: threading.Lock()}
_file_locks_lock = threading.Lock()

# =================== 批量进度 ===================
_batch_progress = {'running': False, 'current': '', 'done': 0, 'total': 0, 'details': []}


def _get_file_lock(filepath):
    """获取文件级别的锁"""
    with _file_locks_lock:
        if filepath not in _file_locks:
            _file_locks[filepath] = threading.Lock()
        return _file_locks[filepath]


@contextmanager
def file_lock(filepath):
    """文件操作锁上下文管理器"""
    lock = _get_file_lock(filepath)
    acquired = lock.acquire(timeout=30)
    if not acquired:
        raise TimeoutError(f'文件 {os.path.basename(filepath)} 正在被系统处理，请稍后重试')
    try:
        yield
    finally:
        lock.release()


def _is_yellow_fill(cell):
    """检查单元格是否有黄色填充"""
    try:
        fill = cell.fill
        if not fill or fill.patternType is None or fill.patternType == 'none':
            return False
        fg = fill.fgColor
        if fg and fg.rgb and str(fg.rgb) not in ('00000000', '0'):
            rgb_str = str(fg.rgb).upper()
            if len(rgb_str) == 8:
                r, g, b = int(rgb_str[2:4], 16), int(rgb_str[4:6], 16), int(rgb_str[6:8], 16)
            elif len(rgb_str) == 6:
                r, g, b = int(rgb_str[0:2], 16), int(rgb_str[2:4], 16), int(rgb_str[4:6], 16)
            else:
                return False
            # 黄色：R高, G高, B低
            if r > 200 and g > 180 and b < 100:
                return True
    except:
        pass
    return False


def _sku_key(sku):
    return re.sub(r'[^0-9]', '', str(sku))[:5]


def _is_ma_sheet(sn):
    """判断是否为MA材料sheet（应跳过）。
    跳过：纯"MA"、"彩盒MA"、"布料MA"、"MA包装"等
    保留：带产品前缀的排期sheet如"游水MA彩盒"（去掉MA+材料词后还有内容）"""
    snu = sn.upper()
    if 'MA' not in snu:
        return False
    cleaned = sn.replace('MA', '').replace('ma', '').strip()
    if not cleaned:
        return True  # 纯"MA"名 → 材料sheet
    temp = cleaned
    for _kw in ('彩盒', '半成品', '包装', '产品', '客版', '布料', '成品'):
        temp = temp.replace(_kw, '')
    if not temp.strip():
        return True  # 纯材料sheet（如"布料MA"、"MA包装"、"彩盒MA"）
    return False  # 有产品前缀（如"游水MA彩盒"→"游水"），保留


# =================== 繁体→简体转换（表头检测用）===================
_TRAD_TO_SIMP = str.maketrans(
    '貨驗國備註辦單數產業務號額價條碼總內類據際計劃種組廠區',
    '货验国备注办单数产业务号额价条码总内类据际计划种组厂区'
)

def _t2s(text):
    """繁体转简体（仅常用字，用于表头检测）"""
    return text.translate(_TRAD_TO_SIMP)


def _item_code(s):
    """提取基础商品代码: '125160H-S001' → '125160H', '15760UQ1' → '15760UQ1'"""
    if not s:
        return ''
    s = re.sub(r'[\s\n]+', '', str(s).strip())
    # 取第一段（'-'之前）的数字+字母部分
    base = s.split('-')[0]
    m = re.match(r'(\d+[A-Za-z]*\d*)', base)
    return m.group(1).upper() if m else ''


def _sku_spec(s):
    """提取完整SKU规格码（含-SXXX后缀）: '92105-S001' → '92105-S001', '125160H-S001' → '125160H-S001'
    用于区分同基础货号的不同变体（如92105-S001/S004有不同包装）"""
    if not s:
        return ''
    s = re.sub(r'[\s\n]+', '', str(s).strip()).upper()
    # 匹配 数字[字母][数字]-SXXX 格式
    m = re.match(r'(\d+[A-Za-z]*\d*(?:-S\d+)?)', s, re.IGNORECASE)
    return m.group(1).upper() if m else _item_code(s)


def _normalize_date(s):
    """将各种日期格式统一为YYYY-MM-DD，支持YYYY-MM-DD、DD-MM-YYYY、MM-DD-YYYY"""
    if not s:
        return ''
    s = str(s).strip().replace('/', '-')
    # 已经是YYYY-MM-DD
    m = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    # DD-MM-YYYY 或 MM-DD-YYYY
    m = re.match(r'(\d{1,2})-(\d{1,2})-(\d{4})', s)
    if m:
        a, b, year = int(m.group(1)), int(m.group(2)), m.group(3)
        if a > 12:   # a一定是日，格式DD-MM-YYYY
            return f"{year}-{b:02d}-{a:02d}"
        elif b > 12:  # b一定是日，格式MM-DD-YYYY
            return f"{year}-{a:02d}-{b:02d}"
        else:         # 都<=12，默认MM-DD-YYYY（商业惯用）
            return f"{year}-{a:02d}-{b:02d}"
    return s


def _parse_date(s):
    """解析日期字符串为datetime，无法解析时返回None（不再默认返回now()）"""
    if isinstance(s, datetime):
        return s
    if isinstance(s, date):
        return datetime(s.year, s.month, s.day)
    if not s:
        return None
    ns = _normalize_date(str(s))
    try:
        return datetime.strptime(ns, '%Y-%m-%d')
    except:
        return None


def _date_serial(dt):
    """将datetime转为Excel序列号（整数天数），避免pywin32时区偏移问题
    pywin32写datetime时会自动做UTC转换（CST-8h），导致日期差一天。
    直接写序列号不经过时区转换，结果精确。"""
    if not dt or not hasattr(dt, 'year'):
        return None
    # Excel序列号: 天数从1899/12/30起算（兼容Lotus 1-2-3 bug）
    return (dt - datetime(1899, 12, 30)).days


def _sv_com(ws, r, c, v, d=False):
    """通过COM设置单元格值"""
    cell = ws.Cells(r, c)
    if d and v:
        if isinstance(v, str) and v:
            try:
                dt = datetime.strptime(v.replace('/', '-'), '%Y-%m-%d')
                cell.Value = _date_serial(dt)
                return
            except:
                pass
        if isinstance(v, datetime):
            cell.Value = _date_serial(v)
            return
    if v is not None and v != '':
        cell.Value = v


def _col_num(letter):
    """列字母→数字: A=1, B=2, ..., Z=26, AA=27"""
    r = 0
    for c in letter.upper():
        r = r * 26 + (ord(c) - ord('A') + 1)
    return r


class ExcelHandler:

    def __init__(self, config):
        self.z_path = config.get('z_drive_path', r'Z:\各客排期\ZURU生产排期')

    # =================== 只读搜索 (openpyxl read_only) ===================

    _xlsx_list_cache = None
    _xlsx_list_path = None

    def _list_xlsx(self):
        # 缓存文件列表，同一路径不重复遍历
        if ExcelHandler._xlsx_list_cache is not None and ExcelHandler._xlsx_list_path == self.z_path:
            return ExcelHandler._xlsx_list_cache
        files = []
        if not os.path.isdir(self.z_path):
            return files
        for root, dirs, fnames in os.walk(self.z_path):
            for item in fnames:
                if item.endswith('.xlsx') and not item.startswith('~$'):
                    files.append(os.path.join(root, item))
        ExcelHandler._xlsx_list_cache = files
        ExcelHandler._xlsx_list_path = self.z_path
        return files

    @classmethod
    def clear_cache(cls):
        """清除文件列表缓存（路径切换或刷新时调用）"""
        cls._xlsx_list_cache = None
        cls._xlsx_list_path = None

    # =================== SKU→排期映射 ===================

    def _get_sku_mapping(self):
        """获取SKU→排期映射（优先JSON文件，带缓存）"""
        global _sku_map_cache, _sku_map_mtime
        # 优先从 data/sku_mapping.json 读取
        json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'sku_mapping.json')
        if os.path.exists(json_path):
            try:
                mtime = os.path.getmtime(json_path)
                if _sku_map_cache and mtime == _sku_map_mtime:
                    return _sku_map_cache
                _sku_map_cache = self._load_sku_mapping_json(json_path)
                _sku_map_mtime = mtime
                logging.info(f"[SKU映射] 从JSON加载 {len(_sku_map_cache)} 个映射项")
                return _sku_map_cache
            except Exception as e:
                logging.warning(f"[SKU映射] JSON读取失败: {e}")
        # 回退：从总排期Excel读取
        master = self.find_master_schedule()
        if not master:
            return _sku_map_cache or {}
        try:
            mtime = os.path.getmtime(master)
        except:
            return _sku_map_cache or {}
        if _sku_map_cache and mtime == _sku_map_mtime:
            return _sku_map_cache
        _sku_map_cache = self._load_sku_mapping_excel(master)
        _sku_map_mtime = mtime
        logging.info(f"[SKU映射] 从Excel加载 {len(_sku_map_cache)} 个映射项")
        return _sku_map_cache

    def _load_sku_mapping_json(self, json_path):
        """从 data/sku_mapping.json 加载映射"""
        import json as _json
        with open(json_path, 'r', encoding='utf-8') as f:
            data = _json.load(f)
        return data.get('mapping', {})

    def _get_sheet_mapping(self):
        """获取货号→工作簿名称映射（从sku_mapping.json的sheet_mapping段）"""
        json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'sku_mapping.json')
        if not os.path.exists(json_path):
            return {}
        try:
            import json as _json
            with open(json_path, 'r', encoding='utf-8') as f:
                data = _json.load(f)
            return data.get('sheet_mapping', {})
        except:
            return {}

    def _load_sku_mapping_excel(self, master_fp):
        """从总排期"对应排期-货号"Sheet加载 SKU→排期文件关键词 映射"""
        try:
            wb = openpyxl.load_workbook(master_fp, read_only=True, data_only=True)
        except:
            return {}
        target = None
        for sn in wb.sheetnames:
            if '对应' in sn and '货号' in sn:
                target = sn
                break
        if not target:
            wb.close()
            return {}
        ws = wb[target]
        mapping = {}
        current_keywords = []

        for row in ws.iter_rows(min_row=2, max_col=20):
            first_val = str(row[0].value or '').strip()
            if re.search(r'[Zz][Uu][Rr][Uu]', first_val) and re.search(r'20\d{2}', first_val):
                current_keywords = re.findall(r'(\d{4,5})', first_val)
                continue
            if not current_keywords:
                continue
            for cell in row:
                val = str(cell.value or '').strip()
                if not val:
                    continue
                val = re.sub(r'.*?明[细細][:：]\s*', '', val)
                for token in re.split(r'[\s,;，；]+', val):
                    token = token.strip()
                    if not token:
                        continue
                    if re.match(r'^\d{4,6}$', token):
                        mapping[token] = current_keywords
                    elif re.match(r'^[A-Za-z]+\d+$', token):
                        mapping[token.upper()] = current_keywords
                    elif re.match(r'^\d+[A-Za-z]+\d*$', token):
                        mapping[token.upper()] = current_keywords
        wb.close()
        return mapping

    def get_sku_mapping_info(self):
        """返回SKU映射信息（供API调用）"""
        mapping = self._get_sku_mapping()
        # 按排期文件分组
        grouped = {}
        for sku, keywords in mapping.items():
            key = ','.join(keywords)
            if key not in grouped:
                grouped[key] = {'keywords': keywords, 'skus': []}
            grouped[key]['skus'].append(sku)
        return {
            'total': len(mapping),
            'groups': len(grouped),
            'detail': [{'keywords': g['keywords'], 'skus': sorted(g['skus']),
                         'count': len(g['skus'])} for g in grouped.values()]
        }

    def auto_find(self, sku):
        """自动查找SKU对应的排期文件和工作表（优先使用映射表）"""
        num = _sku_key(sku)
        sku_upper = re.sub(r'[^A-Za-z0-9]', '', str(sku)).upper()
        item = _item_code(sku)  # 基础货号，如 '125160H'
        spec = _sku_spec(sku)   # 完整规格码，如 '125160H-S001'（区分变体）
        # 提取item的前导数字部分作为备选文件名匹配（解决9548-S001→95480不匹配9548文件名）
        _m = re.match(r'\d+', item) if item else None
        item_digits = _m.group() if _m else ''

        # ===== 第1步：通过SKU映射表查找 =====
        mapping = self._get_sku_mapping()
        sheet_map = self._get_sheet_mapping()
        file_keywords = None
        target_sheet = None  # 从sheet_mapping获取的目标工作簿
        lookup_keys = [sku_upper, num, item.upper() if item else '', item_digits]
        if mapping:
            for key in lookup_keys:
                if key and mapping.get(key):
                    file_keywords = mapping[key]
                    break
        # 查找目标工作簿名称
        if sheet_map:
            for key in lookup_keys:
                if key and key in sheet_map and not key.startswith('_'):
                    target_sheet = sheet_map[key]
                    break

        if file_keywords:
            # 有target_sheet时，优先搜索含"排期"的文件（生产排期），再搜其他
            all_matched = []
            for fp in self._list_xlsx():
                fn = os.path.basename(fp)
                if '总' in fn or '样板' in fn:
                    continue
                # 有target_sheet时，跳过旧排期目录（优先当年排期）
                if target_sheet and ('旧排期' in fp or '旧排期' in fn):
                    continue
                if not any(kw in fn for kw in file_keywords):
                    continue
                all_matched.append((fp, fn))
            # 排序：含"排期"的文件排前面
            if target_sheet:
                all_matched.sort(key=lambda x: (0 if '排期' in x[1] else 1, x[1]))
            for fp, fn in all_matched:
                result = self._search_sku_in_file(fp, fn, num, sku_upper, item,
                                                  target_sheet=target_sheet, sku_spec=spec)
                if result:
                    return result

        # ===== 第2步：原有逻辑（按文件名包含数字匹配）=====
        # 构建候选匹配词：num + item纯数字（去重）
        candidates = set()
        if num and len(num) >= 4:
            candidates.add(num)
        if item_digits and len(item_digits) >= 4:
            candidates.add(item_digits)

        # 检查SKU是否含年份版本（如"9298-2025-S001-NB"中的"2025"）
        sku_year = ''
        for part in str(sku).split('-'):
            if re.match(r'^20\d{2}$', part):
                sku_year = part
                break

        if not candidates:
            # 无候选词时直接进入兜底搜索
            pass
        else:
            best = None
            for fp in self._list_xlsx():
                fn = os.path.basename(fp)
                if '总' in fn or '样板' in fn:
                    continue
                if not any(c in fn for c in candidates):
                    continue
                # 含年份版本时，优先匹配同时含年份的文件
                if sku_year and sku_year not in fn:
                    continue
                result = self._search_sku_in_file(fp, fn, num, sku_upper, item, sku_spec=spec)
                if result:
                    if not best or result['cnt'] > best['cnt']:
                        best = result
            # 若年份过滤太严格未匹配到，放宽年份限制重试
            if not best and sku_year:
                for fp in self._list_xlsx():
                    fn = os.path.basename(fp)
                    if '总' in fn or '样板' in fn:
                        continue
                    if not any(c in fn for c in candidates):
                        continue
                    result = self._search_sku_in_file(fp, fn, num, sku_upper, item, sku_spec=spec)
                    if result:
                        if not best or result['cnt'] > best['cnt']:
                            best = result
            if not best and not sku_year:
                # 无年份也无匹配，正常逻辑
                for fp in self._list_xlsx():
                    fn = os.path.basename(fp)
                    if '总' in fn or '样板' in fn:
                        continue
                    if not any(c in fn for c in candidates):
                        continue
                    result = self._search_sku_in_file(fp, fn, num, sku_upper, item, sku_spec=spec)
                    if result:
                        if not best or result['cnt'] > best['cnt']:
                            best = result
            if best:
                return best

        # ===== 第3步：兜底搜索（前两步均未匹配，搜索所有排期文件内容）=====
        logging.info(f"[auto_find] SKU '{sku}' 前两步未匹配，启动全文件搜索")
        best = None
        for fp in self._list_xlsx():
            fn = os.path.basename(fp)
            if '总' in fn or '样板' in fn:
                continue
            result = self._search_sku_in_file(fp, fn, num, sku_upper, item, sku_spec=spec)
            if result:
                if not best or result['cnt'] > best['cnt']:
                    best = result
        if not best:
            logging.warning(f"[auto_find] SKU '{sku}' (num={num}, item={item}, spec={spec}) 未匹配到任何排期文件")
        return best

    def _search_sku_in_file(self, fp, fn, num, sku_upper, item_code='', target_sheet=None, sku_spec=''):
        """在指定文件中搜索SKU，返回匹配结果
        只在ITEM#列（G列=index6）搜索货号，匹配优先级：
        1. SKU-SPEC精确匹配（如92105-S001），区分同货号不同变体
        2. 基础货号精确匹配（如92105）
        3. 前缀匹配（如9548匹配9548G4，长度差≤3）
        target_sheet: 从sheet_mapping指定的目标工作簿名称，有则只搜该sheet
        sku_spec: 完整规格码（含-SXXX后缀），用于精确区分变体"""
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception as e:
            logging.warning(f"[auto_find] 无法打开 {fn}: {e}")
            return None
        best = None
        # 当有target_sheet时，优先精确匹配sheet名称
        sheets_to_search = wb.sheetnames
        matched_target_sheets = set()
        if target_sheet:
            # 1.精确匹配
            matched = [sn for sn in wb.sheetnames if sn == target_sheet]
            if not matched:
                # 2.双向包含匹配（如"15701明细"匹配"15701"，或"恐龙"匹配"恐龙明细"）
                matched = [sn for sn in wb.sheetnames if target_sheet in sn or sn in target_sheet]
            if not matched:
                # 3.前导数字匹配（如target_sheet="9548明细"匹配sheet名含"9548"的）
                ts_digits = re.match(r'\d+', target_sheet)
                if ts_digits:
                    ts_num = ts_digits.group()
                    matched = [sn for sn in wb.sheetnames if ts_num in sn and '取消' not in sn and not _is_ma_sheet(sn)]
            if matched:
                sheets_to_search = matched
                matched_target_sheets = set(matched)
                logging.info(f"[auto_find] 使用sheet_mapping定位: {fn} → {matched[0]}")
            else:
                # target_sheet指定但此文件中找不到 → 跳过此文件，不搜索其他sheet
                logging.info(f"[auto_find] {fn} 中未找到目标sheet '{target_sheet}'，跳过此文件")
                try:
                    wb.close()
                except:
                    pass
                return None

        for sn in sheets_to_search:
            if '取消' in sn or '对应' in sn:
                continue
            if '总' in sn:
                continue
            if '旧' in sn:
                continue
            if '样板' in sn:
                continue
            if _is_ma_sheet(sn):
                continue
            try:
                ws = wb[sn]
                # 三级匹配：spec精确 > base精确 > 前缀
                ref_spec_named = None    # SKU-SPEC精确匹配（有产品名）
                ref_spec_any = None      # SKU-SPEC精确匹配（无产品名）
                cnt_spec = 0
                ref_exact_named = None   # 基础货号精确匹配（有产品名）
                ref_exact_any = None     # 基础货号精确匹配（无产品名）
                cnt_exact = 0
                ref_prefix_named = None
                ref_prefix_any = None
                cnt_prefix = 0
                last_data_row = None
                for row in ws.iter_rows(min_row=2, max_col=10):
                    row_num = getattr(row[0], 'row', None)
                    if row_num is None:
                        for c in row[1:]:
                            row_num = getattr(c, 'row', None)
                            if row_num is not None:
                                break
                    if row_num is None:
                        continue
                    has_any_data = any(ci < len(row) and row[ci].value for ci in range(min(8, len(row))))
                    if has_any_data:
                        last_data_row = row_num
                    # 只在ITEM#列（G=6）和备选列（F=5, H=7）搜索货号
                    # 不搜索D(3)、E(4)列（PO号/客户PO，会产生子串误匹配）
                    for ci in [6, 5, 7]:
                        if ci >= len(row) or not row[ci].value:
                            continue
                        cv = str(row[ci].value).strip()
                        cv_item = _item_code(cv)
                        if not cv_item:
                            continue
                        has_name = len(row) > 7 and row[7].value and str(row[7].value).strip()
                        cv_spec = _sku_spec(cv)
                        # 第1级：SKU-SPEC精确匹配（如92105-S001精确匹配92105-S001）
                        if sku_spec and cv_spec == sku_spec:
                            ref_spec_any = row_num
                            cnt_spec += 1
                            if has_name:
                                ref_spec_named = row_num
                            break
                        # 第2级：基础货号精确匹配（如92105精确匹配92105）
                        elif item_code and cv_item == item_code:
                            ref_exact_any = row_num
                            cnt_exact += 1
                            if has_name:
                                ref_exact_named = row_num
                            break
                        # 第3级：前缀匹配（如"9548"匹配"9548G4"，长度差≤3防止误匹配）
                        elif item_code and cv_item and \
                             abs(len(cv_item) - len(item_code)) <= 3 and \
                             (cv_item.startswith(item_code) or item_code.startswith(cv_item)):
                            ref_prefix_any = row_num
                            cnt_prefix += 1
                            if has_name:
                                ref_prefix_named = row_num
                            break
                # 优先级：spec精确 > base精确 > 前缀
                ref = ref_spec_named or ref_spec_any or ref_exact_named or ref_exact_any or ref_prefix_named or ref_prefix_any
                cnt = cnt_spec or cnt_exact or cnt_prefix
                if ref and sku_spec and ref_spec_named:
                    logging.info(f"[auto_find] {fn}/{sn} SKU-SPEC精确匹配: spec={sku_spec}, ref={ref}")
                elif ref and not ref_spec_named and not ref_spec_any and item_code:
                    logging.info(f"[auto_find] {fn}/{sn} 仅基础货号匹配: item={item_code}(spec={sku_spec}), ref={ref}")
                if ref:
                    result = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': ref,
                              'cnt': cnt, 'mcol': ws.max_column or 30}
                    if not best or cnt > best['cnt']:
                        best = result
                elif target_sheet and sn in matched_target_sheets:
                    if last_data_row:
                        logging.info(f"[auto_find] {fn}/{sn} 无精确匹配行，使用最后数据行{last_data_row}作参考")
                        best = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': last_data_row,
                                'cnt': 1, 'mcol': ws.max_column or 30}
            except Exception as e:
                logging.debug(f"[SKU搜索] {fn}/{sn} read_only模式跳过: {e}")
                # 如果是目标sheet且read_only模式失败，用非只读模式重试
                if target_sheet and sn in matched_target_sheets:
                    logging.info(f"[auto_find] {fn}/{sn} read_only失败，尝试非只读模式重试")
                    try:
                        wb2 = openpyxl.load_workbook(fp, read_only=False, data_only=True)
                        ws2 = wb2[sn]
                        ref_spec2 = None
                        ref_retry = None
                        cnt_retry = 0
                        last_row2 = None
                        for r in range(2, ws2.max_row + 1):
                            has_data = any(ws2.cell(r, c).value for c in range(1, min(9, ws2.max_column + 1)))
                            if has_data:
                                last_row2 = r
                            for ci in [7, 6, 8]:  # G, F, H (1-based)
                                cv = ws2.cell(r, ci).value
                                if not cv:
                                    continue
                                cv_str = str(cv).strip()
                                cv_item = _item_code(cv_str)
                                if not cv_item:
                                    continue
                                cv_sp = _sku_spec(cv_str)
                                # SKU-SPEC精确匹配（最高优先级）
                                if sku_spec and cv_sp == sku_spec:
                                    ref_spec2 = r
                                    cnt_retry += 1
                                    break
                                elif item_code and cv_item == item_code:
                                    ref_retry = r
                                    cnt_retry += 1
                                    break
                                elif item_code and cv_item and \
                                     abs(len(cv_item) - len(item_code)) <= 3 and \
                                     (cv_item.startswith(item_code) or item_code.startswith(cv_item)):
                                    if not ref_retry:
                                        ref_retry = r
                                    cnt_retry += 1
                                    break
                        ref_retry = ref_spec2 or ref_retry
                        if ref_retry:
                            best = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': ref_retry,
                                    'cnt': cnt_retry, 'mcol': ws2.max_column or 30}
                            logging.info(f"[auto_find] 非只读重试成功: {fn}/{sn} 行{ref_retry}")
                        elif last_row2:
                            best = {'file': fp, 'fname': fn, 'sheet': sn, 'ref': last_row2,
                                    'cnt': 1, 'mcol': ws2.max_column or 30}
                            logging.info(f"[auto_find] 非只读重试: {fn}/{sn} 使用最后数据行{last_row2}")
                        wb2.close()
                    except Exception as e2:
                        logging.warning(f"[auto_find] {fn}/{sn} 非只读重试也失败: {e2}")
                        # openpyxl完全无法读取，使用WPS COM后备搜索
                        try:
                            logging.info(f"[auto_find] {fn}/{sn} 启用WPS COM后备搜索")
                            com_result = self._search_sku_com(fp, fn, sn, item_code, sku_spec=sku_spec)
                            if com_result:
                                best = com_result
                        except Exception as e3:
                            logging.warning(f"[auto_find] {fn}/{sn} COM后备也失败: {e3}")
                continue
        try:
            wb.close()
        except:
            pass
        return best

    def _search_sku_com(self, fp, fn, sheet_name, item_code, sku_spec=''):
        """WPS COM后备搜索：当openpyxl无法读取时用COM打开文件搜索ITEM#列
        sku_spec: 完整规格码（含-SXXX后缀），用于精确区分变体"""
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = None
        wb = None
        try:
            for pid in ['Ket.Application', 'Et.Application', 'Excel.Application']:
                try:
                    app = win32com.client.DispatchEx(pid)
                    break
                except:
                    continue
            if not app:
                return None
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(fp, ReadOnly=True)
            ws = None
            for i in range(1, wb.Sheets.Count + 1):
                if wb.Sheets(i).Name == sheet_name:
                    ws = wb.Sheets(i)
                    break
            if not ws:
                return None
            max_row = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
            max_col = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
            ref_spec = None
            ref = None
            cnt = 0
            last_data_row = None
            for r in range(2, min(max_row + 1, 2000)):
                has_data = False
                for c in range(1, min(9, max_col + 1)):
                    v = ws.Cells(r, c).Value
                    if v:
                        has_data = True
                        break
                if has_data:
                    last_data_row = r
                # 搜索G(7), F(6), H(8)列
                for ci in [7, 6, 8]:
                    cv = ws.Cells(r, ci).Value
                    if not cv:
                        continue
                    cv_str = str(cv).strip()
                    cv_item = _item_code(cv_str)
                    if not cv_item:
                        continue
                    cv_sp = _sku_spec(cv_str)
                    # SKU-SPEC精确匹配（最高优先级）
                    if sku_spec and cv_sp == sku_spec:
                        ref_spec = r
                        cnt += 1
                        break
                    elif item_code and cv_item == item_code:
                        ref = r
                        cnt += 1
                        break
                    elif item_code and cv_item and \
                         abs(len(cv_item) - len(item_code)) <= 3 and \
                         (cv_item.startswith(item_code) or item_code.startswith(cv_item)):
                        if not ref:
                            ref = r
                        cnt += 1
                        break
            ref = ref_spec or ref
            if ref:
                logging.info(f"[auto_find] COM后备成功: {fn}/{sheet_name} 行{ref}")
                return {'file': fp, 'fname': fn, 'sheet': sheet_name, 'ref': ref,
                        'cnt': cnt, 'mcol': max_col or 30}
            elif last_data_row:
                logging.info(f"[auto_find] COM后备: {fn}/{sheet_name} 使用最后数据行{last_data_row}")
                return {'file': fp, 'fname': fn, 'sheet': sheet_name, 'ref': last_data_row,
                        'cnt': 1, 'mcol': max_col or 30}
            return None
        except Exception as e:
            logging.warning(f"[auto_find] COM后备搜索异常: {e}")
            return None
        finally:
            try:
                if wb:
                    wb.Close(False)
            except:
                pass
            try:
                if app:
                    app.Quit()
            except:
                pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def search_po(self, po_number):
        po_s = str(po_number)
        po_i = int(po_s) if po_s.isdigit() else None
        results = []
        for fp in self._list_xlsx():
            fn = os.path.basename(fp)
            if '总' in fn:
                continue
            try:
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            except:
                continue
            for sn in wb.sheetnames:
                if '取消' in sn:
                    continue
                if _is_ma_sheet(sn):
                    continue
                try:
                    ws = wb[sn]
                    for row in ws.iter_rows(min_row=2, max_col=30):
                        row_num = getattr(row[0], 'row', None)
                        if row_num is None:
                            for c in row[1:]:
                                row_num = getattr(c, 'row', None)
                                if row_num is not None:
                                    break
                        if row_num is None:
                            continue
                        d = row[3].value if len(row) > 3 else None
                        e = row[4].value if len(row) > 4 else None
                        hit = False
                        if d and (str(d) == po_s or (po_i and d == po_i) or po_s in str(d)):
                            hit = True
                        elif e and (str(e) == po_s or (po_i and e == po_i) or po_s in str(e)):
                            hit = True
                        if hit:
                            data = {}
                            for c in row:
                                try:
                                    col_num = getattr(c, 'column', None)
                                    if col_num is None:
                                        continue
                                    cl = openpyxl.utils.get_column_letter(col_num)
                                    v = c.value
                                    if isinstance(v, datetime):
                                        v = v.strftime('%Y-%m-%d')
                                    data[cl] = v
                                except:
                                    pass
                            results.append({'file': fp, 'fname': fn,
                                            'sheet': sn, 'row': row_num, 'data': data})
                except Exception as e:
                    logging.debug(f"[PO搜索] {fn}/{sn} 跳过: {e}")
                    continue
            try:
                wb.close()
            except:
                continue
        return results

    def batch_search_pos(self, po_list):
        """并行搜索多个PO号，一次遍历所有文件。返回 {po_str: [results]}"""
        from concurrent.futures import ThreadPoolExecutor, as_completed
        po_set = set()
        po_int_set = {}
        for p in po_list:
            if not p:
                continue
            ps = str(p).strip()
            po_set.add(ps)
            if ps.isdigit():
                po_int_set[int(ps)] = ps
        if not po_set:
            return {}
        all_results = {p: [] for p in po_set}
        files = [(fp, os.path.basename(fp)) for fp in self._list_xlsx()
                 if '总' not in os.path.basename(fp)]

        def _scan_file(args):
            fp, fn = args
            hits = []
            try:
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            except:
                return hits
            for sn in wb.sheetnames:
                if '取消' in sn:
                    continue
                if _is_ma_sheet(sn):
                    continue
                try:
                    ws = wb[sn]
                    for row in ws.iter_rows(min_row=2, max_col=30):
                        row_num = getattr(row[0], 'row', None)
                        if row_num is None:
                            for c in row[1:]:
                                row_num = getattr(c, 'row', None)
                                if row_num is not None:
                                    break
                        if row_num is None:
                            continue
                        d = row[3].value if len(row) > 3 else None
                        e = row[4].value if len(row) > 4 else None
                        matched = None
                        for val in [d, e]:
                            if val is None:
                                continue
                            vs = str(val).strip()
                            if vs in po_set:
                                matched = vs
                                break
                            if isinstance(val, (int, float)) and int(val) in po_int_set:
                                matched = po_int_set[int(val)]
                                break
                            for ps in po_set:
                                if ps in vs:
                                    matched = ps
                                    break
                            if matched:
                                break
                        if matched:
                            data = {}
                            for c in row:
                                try:
                                    cn = getattr(c, 'column', None)
                                    if cn is None:
                                        continue
                                    cl = openpyxl.utils.get_column_letter(cn)
                                    v = c.value
                                    if isinstance(v, datetime):
                                        v = v.strftime('%Y-%m-%d')
                                    data[cl] = v
                                except:
                                    pass
                            hits.append((matched, {'file': fp, 'fname': fn,
                                                   'sheet': sn, 'row': row_num, 'data': data}))
                except:
                    continue
            try:
                wb.close()
            except:
                pass
            return hits

        workers = min(6, len(files))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            for result in pool.map(_scan_file, files):
                for po_str, rec in result:
                    all_results[po_str].append(rec)
        return all_results

    def search_by_skus(self, lines):
        """当PO号搜不到时，通过SKU/商品代码在排期中搜索现有记录
        1. 用auto_find定位排期文件和工作表
        2. 在该工作表中搜索所有包含PDF商品代码的行"""
        results = []
        # 收集所有item code
        code_set = set()
        for ln in lines:
            for field in ('sku', 'item_code'):
                v = ln.get(field, '')
                code = _item_code(v)
                if code:
                    code_set.add(code)
        if not code_set:
            return results

        # 用第一个SKU定位排期文件/工作表
        target_files = {}  # file_path → set of sheet names
        for ln in lines:
            sku = ln.get('item_code') or ln.get('sku', '')
            found = self.auto_find(sku)
            if found:
                fp = found['file']
                if fp not in target_files:
                    target_files[fp] = set()
                target_files[fp].add(found['sheet'])
        if not target_files:
            return results

        # 在目标文件的目标sheet中搜索包含任一商品代码的行
        for fp, sheets in target_files.items():
            fn = os.path.basename(fp)
            if '总' in fn:
                continue
            try:
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            except:
                continue
            for sn in sheets:
                if sn not in wb.sheetnames or '取消' in sn:
                    continue
                if _is_ma_sheet(sn):
                    continue
                try:
                    ws = wb[sn]
                    for row in ws.iter_rows(min_row=2, max_col=30):
                        row_num = getattr(row[0], 'row', None)
                        if row_num is None:
                            for c in row[1:]:
                                row_num = getattr(c, 'row', None)
                                if row_num is not None:
                                    break
                        if row_num is None:
                            continue
                        hit = False
                        for c in row[:10]:
                            if c.value:
                                code = _item_code(str(c.value))
                                if code and code in code_set:
                                    hit = True
                                    break
                        if hit:
                            data = {}
                            for c in row:
                                try:
                                    col_num = getattr(c, 'column', None)
                                    if col_num is None:
                                        continue
                                    cl = openpyxl.utils.get_column_letter(col_num)
                                    v = c.value
                                    if isinstance(v, datetime):
                                        v = v.strftime('%Y-%m-%d')
                                    data[cl] = v
                                except:
                                    pass
                            results.append({'file': fp, 'fname': fn,
                                            'sheet': sn, 'row': row_num, 'data': data})
                except:
                    continue
            try:
                wb.close()
            except:
                pass
        return results

    def fuzzy_search(self, keyword):
        """模糊搜索：支持PO号、SKU、客户名等"""
        kw = str(keyword).strip()
        kw_lower = kw.lower()
        kw_num = re.sub(r'[^0-9]', '', kw)
        results = []
        for fp in self._list_xlsx():
            fn = os.path.basename(fp)
            if '总' in fn:
                continue
            try:
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            except:
                continue
            for sn in wb.sheetnames:
                if '取消' in sn:
                    continue
                if _is_ma_sheet(sn):
                    continue
                try:
                    ws = wb[sn]
                    for row in ws.iter_rows(min_row=2, max_col=30):
                        row_num = getattr(row[0], 'row', None)
                        if row_num is None:
                            for c in row[1:]:
                                row_num = getattr(c, 'row', None)
                                if row_num is not None:
                                    break
                        if row_num is None:
                            continue
                        hit = False
                        hit_col = ''
                        for c in row[:10]:
                            if c.value:
                                cv = str(c.value)
                                col_num = getattr(c, 'column', 0)
                                if kw_lower in cv.lower() or (kw_num and len(kw_num) >= 4 and kw_num in cv):
                                    hit = True
                                    col_names = {1:'接单日期', 2:'客户', 3:'目的地', 4:'PO号',
                                                 5:'客户PO', 6:'SKU', 7:'品名', 9:'数量', 13:'出货日期'}
                                    hit_col = col_names.get(col_num, f'列{col_num}')
                                    break
                        if hit:
                            data = {}
                            for c in row:
                                try:
                                    cn = getattr(c, 'column', None)
                                    if cn is None:
                                        continue
                                    cl = openpyxl.utils.get_column_letter(cn)
                                    v = c.value
                                    if isinstance(v, datetime):
                                        v = v.strftime('%Y-%m-%d')
                                    data[cl] = v
                                except:
                                    pass
                            results.append({
                                'file': fp, 'fname': fn, 'sheet': sn,
                                'row': row_num, 'data': data, 'hit_col': hit_col
                            })
                            if len(results) >= 100:
                                try:
                                    wb.close()
                                except:
                                    pass
                                return results
                except:
                    continue
            try:
                wb.close()
            except:
                pass
        return results

    # =================== 智能对比（纯逻辑） ===================

    def smart_diff(self, pdf_data, existing_records):
        """比对PDF与现有记录，生成新增/修改操作
        原则：
        1. 同PO记录：检测变化→修改；无变化→跳过
        2. 不同PO记录：只标记该商品代码已存在→不重复添加，不修改
        3. 未匹配记录：不自动取消（取消由用户手动操作）
        匹配策略：PO-line精确匹配 → 完整商品代码匹配"""
        actions = []
        new_lines = pdf_data.get('lines', [])
        po = pdf_data.get('po_number', '')

        # ===== 1. 构建PDF行查找表 =====
        new_by_code = {}       # 完整商品代码 → line (如 "125160H" → line)
        new_by_poline = {}     # "PO-lineNo" → line

        for ln in new_lines:
            sku = ln.get('sku', '')
            code = _item_code(sku)
            if code:
                new_by_code[code] = ln
            ic = ln.get('item_code', '')
            ic_code = _item_code(ic)
            if ic_code and ic_code not in new_by_code:
                new_by_code[ic_code] = ln
            line_no = ln.get('line_no', '')
            if po and line_no:
                new_by_poline[f"{po}-{line_no}"] = ln

        matched_codes = set()    # 已匹配的完整商品代码
        matched_polines = set()  # 已匹配的PO-line键

        # ===== 2. 逐条匹配已有记录 =====
        for rec in existing_records:
            rd = rec['data']
            matched_ln = None
            match_code = None

            # 检查D-J列寻找匹配
            for col in 'FGHDEIJ':
                v = rd.get(col)
                if not v:
                    continue
                vs = str(v).strip()
                if not vs:
                    continue
                # 策略1：PO-line精确匹配（如 "4500193745-10"）
                if vs in new_by_poline:
                    matched_ln = new_by_poline[vs]
                    match_code = _item_code(matched_ln.get('sku', ''))
                    matched_polines.add(vs)
                    break
                # 策略2：完整商品代码匹配（如 "125160H" from "125160H-S001"）
                code = _item_code(vs)
                if code and code in new_by_code:
                    matched_ln = new_by_code[code]
                    match_code = code
                    break

            if matched_ln:
                if match_code:
                    matched_codes.add(match_code)

                # 判断是否同一PO：只有同PO才比较和修改
                rec_po = str(rd.get('D', '') or '').strip()
                same_po = bool(po and rec_po and (po in rec_po or rec_po in po))

                if not same_po:
                    # 不同PO → 只标记已存在，不做任何修改
                    continue

                # 同PO → 检测实际变化（不检查接单日期，接单日期不做修改）
                changes = {}
                new_qty = matched_ln.get('qty', 0)
                new_price = matched_ln.get('price', 0)
                new_ship = _normalize_date(pdf_data.get('ship_date', ''))
                new_cpo = matched_ln.get('customer_po', '')

                # 数量：在I-L列中找到第一个数字列比较
                if new_qty:
                    for c in 'IJKL':
                        ov = rd.get(c)
                        if ov is not None:
                            try:
                                if int(float(ov)) > 0:
                                    if int(float(ov)) != int(new_qty):
                                        changes[c] = new_qty
                                    break
                            except:
                                continue

                # 单价：在R-AF列中找到第一个小数（0<x<100）比较
                if new_price:
                    for ci in range(18, 33):
                        c = chr(64 + ci) if ci <= 26 else 'A' + chr(64 + ci - 26)
                        ov = rd.get(c)
                        if ov is not None:
                            try:
                                if 0 < float(ov) < 100:
                                    if abs(float(ov) - new_price) > 0.001:
                                        changes[c] = new_price
                                    break
                            except:
                                continue

                # 出货日期：在M-P列中找到日期列比较
                if new_ship:
                    for c in 'MNOP':
                        ov = rd.get(c)
                        if ov is not None:
                            old_str = ''
                            if hasattr(ov, 'year'):
                                old_str = ov.strftime('%Y-%m-%d')
                            elif isinstance(ov, str) and ov:
                                old_str = _normalize_date(ov)
                            if old_str:
                                if old_str != new_ship:
                                    changes[c] = new_ship
                                break

                # 客户PO：在D-G列中查找匹配，只有确实不同才标记修改
                # 跳过小数值（防止CBM等误判为客PO）
                if new_cpo and not re.match(r'^\d+\.\d+$', new_cpo):
                    cpo_changed = True
                    for c in 'DEFG':
                        ov = str(rd.get(c, '') or '').strip()
                        if ov == new_cpo:
                            cpo_changed = False
                            break
                    # 只有PDF有明确客PO、且排期所有候选列都不匹配时才标修改
                    if cpo_changed:
                        # 找到排期中最可能的客PO列（E或F）
                        # 跳过PO-line格式值（如"4500193080-20"），这是SKU/系统货号列
                        for c in 'EF':
                            ov = str(rd.get(c, '') or '').strip()
                            if not ov:
                                continue
                            # 跳过PO-line格式（PO号-行号）：这是SKU列不能覆写
                            if re.match(r'^\d{7,}-\d+$', ov):
                                continue
                            # 跳过包含当前PO号的值（可能是PO列或SKU列）
                            if po and po in ov:
                                continue
                            if ov != new_cpo:
                                changes[c] = new_cpo
                                break

                if changes:
                    actions.append({
                        'type': 'modify', 'record': rec, 'changes': changes,
                        'sku': matched_ln.get('sku', ''),
                        'detail': f"修改 {matched_ln.get('sku','')} " +
                                  ', '.join([f"{k}:{v}" for k, v in changes.items()])
                    })
                # else: 同PO无变化 → 不生成任何操作（跳过）
            # else: 未匹配 → 不自动取消，保持原样

        # ===== 3. PDF中有但排期中没有的行 → 新增 =====
        for ln in new_lines:
            code = _item_code(ln.get('sku', ''))
            line_no = ln.get('line_no', '')
            po_line = f"{po}-{line_no}" if po and line_no else ''
            already = (code and code in matched_codes) or (po_line and po_line in matched_polines)
            if not already:
                sched = self.auto_find(ln.get('sku_spec', '') or ln.get('item_code', '') or ln.get('sku', ''))
                actions.append({
                    'type': 'new', 'line': ln, 'schedule': sched,
                    'sku': ln.get('sku', ''),
                    'detail': f"新增 {ln.get('sku','')} {ln.get('qty',0)}pcs"
                })

        return actions

    # =================== COM 启动 ===================

    @staticmethod
    def _com_app():
        """启动WPS/Excel COM进程"""
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        for pid in ['Ket.Application', 'Et.Application', 'Excel.Application']:
            try:
                app = win32com.client.DispatchEx(pid)
                app.Visible = False
                app.DisplayAlerts = False
                return app
            except:
                continue
        raise RuntimeError("无法启动WPS或Excel，请确认已安装WPS Office")

    @staticmethod
    def _com_quit(app):
        """安全退出COM"""
        if app:
            try:
                app.DisplayAlerts = False
                app.Quit()
            except:
                pass
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except:
            pass

    # =================== 批量处理 (COM写入) ===================

    @staticmethod
    def get_batch_progress():
        """获取批量处理进度"""
        return dict(_batch_progress)

    def batch_process(self, orders):
        global _batch_progress
        os.makedirs(BATCH_DIR, exist_ok=True)
        os.makedirs(UNDO_DIR, exist_ok=True)
        results = []
        failed = []
        _batch_progress = {'running': True, 'current': '分析中...', 'done': 0, 'total': 0, 'details': []}

        # 生成批次ID
        batch_id = datetime.now().strftime('%Y%m%d-%H%M%S')

        # 按排期文件分组
        file_ops = {}
        for order in orders:
            for act in order.get('actions', []):
                if act['type'] == 'new' and act.get('schedule'):
                    fkey = act['schedule']['file']
                    if fkey not in file_ops:
                        file_ops[fkey] = {'file': fkey, 'new': [], 'modify': [], 'cancel': [],
                                          'header': order.get('header', {})}
                    file_ops[fkey]['new'].append(act)
                elif act['type'] == 'modify':
                    fkey = act['record']['file']
                    if fkey not in file_ops:
                        file_ops[fkey] = {'file': fkey, 'new': [], 'modify': [], 'cancel': [],
                                          'header': order.get('header', {})}
                    file_ops[fkey]['modify'].append(act)
                elif act['type'] == 'cancel':
                    fkey = act['record']['file']
                    if fkey not in file_ops:
                        file_ops[fkey] = {'file': fkey, 'new': [], 'modify': [], 'cancel': [],
                                          'header': order.get('header', {})}
                    file_ops[fkey]['cancel'].append(act)

        # 收集每个订单的PO和客户信息，用于按文件构建撤销记录
        order_info = {}
        for order in orders:
            po = order.get('header', {}).get('po_number', '')
            customer = order.get('header', {}).get('customer', '')
            order_info[po] = customer

        # 逐文件处理
        _batch_progress['total'] = len(file_ops)
        app = None
        try:
            app = self._com_app()

            for file_idx, (fkey, ops) in enumerate(file_ops.items()):
                fname = os.path.basename(fkey)
                _batch_progress['current'] = fname
                _batch_progress['done'] = file_idx
                local = os.path.join(BATCH_DIR, fname)
                try:
                    # 只备份当前修改的排期文件（不全量备份）
                    undo_fp = os.path.join(UNDO_DIR, f"{batch_id}_{fname}")
                    shutil.copy2(fkey, undo_fp)
                    shutil.copy2(fkey, local)
                    wb = app.Workbooks.Open(os.path.abspath(local))
                    msg_parts = []

                    # --- 1) 先取消（从大行号往小删，避免行号偏移）---
                    cancel_ops = sorted(ops['cancel'],
                                        key=lambda x: x['record']['row'], reverse=True)
                    deleted_rows = []  # (sheet_name, row) 元组列表，区分不同sheet
                    for act in cancel_ops:
                        sn = act['record']['sheet']
                        rn = act['record']['row']
                        ws = wb.Sheets(sn)
                        mc = min(ws.UsedRange.Columns.Count + ws.UsedRange.Column, 100)
                        self._do_cancel_com(wb, ws, rn, mc)
                        deleted_rows.append((sn, rn))  # 跟踪sheet名
                        msg_parts.append(f"取消{act['sku']}")

                    # --- 2) 修改（调整行号，只对同sheet的删除做偏移）---
                    for act in ops['modify']:
                        sn = act['record']['sheet']
                        orig_row = act['record']['row']
                        # 只对同sheet的删除行做偏移调整（不同sheet行号独立）
                        shift = sum(1 for dsn, d in deleted_rows if dsn == sn and d < orig_row)
                        adj_row = orig_row - shift
                        ws = wb.Sheets(sn)
                        mc = min(ws.UsedRange.Columns.Count + ws.UsedRange.Column, 100)
                        self._do_modify_com(ws, adj_row, mc, act['changes'])
                        msg_parts.append(f"修改{act['sku']}")

                    # --- 3) 新增（动态查找插入位置，按sheet独立跟踪偏移）---
                    inserted_positions = {}  # {sheet_name: [pos_list]}，按sheet独立跟踪
                    last_insert_pos = {}     # {sheet_name: last_pos}，按sheet独立跟踪
                    for act in ops['new']:
                        if not act.get('schedule'):
                            continue
                        sn = act['schedule']['sheet']
                        ref = act['schedule']['ref']
                        mc = min(act['schedule'].get('mcol', 100), 100)
                        # 只对同sheet的删除行做偏移调整
                        shift_del = sum(1 for dsn, d in deleted_rows if dsn == sn and d < ref)
                        adj_ref = ref - shift_del
                        # 只对同sheet的已插入行做偏移调整
                        for p in inserted_positions.get(sn, []):
                            if p <= adj_ref:
                                adj_ref += 1
                        ws = wb.Sheets(sn)
                        pos, w = self._do_new_com(ws, adj_ref, mc, ops['header'], act['line'],
                                               start_after=last_insert_pos.get(sn, 0))
                        inserted_positions.setdefault(sn, []).append(pos)
                        last_insert_pos[sn] = pos
                        warn_tag = ''
                        if w:
                            warn_tag = f" [空字段: {', '.join(w)}]"
                        msg_parts.append(f"新增{act['sku']}{warn_tag}")

                    wb.Save()
                    wb.Close(False)

                    # 尝试保存到Z盘
                    z_ok = False
                    try:
                        self._try_save_z(local, fkey)
                        z_ok = True
                    except:
                        pass

                    r = {'file': fname, 'local': local, 'z': fkey, 'z_saved': z_ok,
                         'msg': ' | '.join(msg_parts),
                         'counts': {'new': len(ops['new']), 'modify': len(ops['modify']),
                                    'cancel': len(ops['cancel'])}}
                    if z_ok:
                        results.append(r)
                        # 每个排期文件单独保存一条撤销记录
                        type_names = {'new': '新增', 'modify': '修改', 'cancel': '取消'}
                        file_ops_list = []
                        for t in ('new', 'modify', 'cancel'):
                            for act in ops[t]:
                                op = {'type': t, 'sku': act.get('sku', ''),
                                      'detail': act.get('detail', '')}
                                if t == 'new' and act.get('line'):
                                    op['qty'] = act['line'].get('qty', 0)
                                file_ops_list.append(op)
                        labels = []
                        for op in file_ops_list:
                            tn = type_names.get(op['type'], op['type'])
                            labels.append(f"{tn} {op['sku']}" +
                                          (f" {op.get('qty','')}pcs" if op.get('qty') else ''))
                        self._save_undo_entry({
                            'id': f"{batch_id}_{fname}",
                            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'operations': file_ops_list,
                            'files': [{'name': fname, 'backup': undo_fp, 'z_path': fkey}],
                            'label': f"[{fname}] " + (' | '.join(labels[:3]) +
                                     (f' 等{len(labels)}项' if len(labels) > 3 else ''))
                        })
                    else:
                        r['reason'] = '文件被占用（只读）'
                        failed.append(r)

                except Exception as e:
                    logging.error(f"[batch] 处理 {fname} 异常: {e}")
                    # 确保workbook关闭
                    try:
                        wb.Close(False)
                    except:
                        pass
                    err_str = str(e)
                    # 区分文件占用和COM处理异常
                    if 'Permission' in err_str or '占用' in err_str or '只读' in err_str:
                        reason = f'文件被占用: {err_str[:80]}'
                    else:
                        reason = f'处理异常: {err_str[:120]}'
                    failed.append({'file': fname, 'local': local, 'z': fkey,
                                   'z_saved': False, 'reason': reason,
                                   'msg': f'处理失败: {e}',
                                   'counts': {'new': len(ops['new']),
                                              'modify': len(ops['modify']),
                                              'cancel': len(ops['cancel'])}})
        finally:
            self._com_quit(app)
            _batch_progress = {'running': False, 'current': '完成', 'done': len(file_ops),
                               'total': len(file_ops), 'details': []}

        return {'results': results, 'failed': failed}

    def _save_undo_entry(self, entry):
        """保存撤销历史条目"""
        os.makedirs(DATA_DIR, exist_ok=True)
        history = []
        if os.path.exists(UNDO_HISTORY):
            try:
                with open(UNDO_HISTORY, 'r', encoding='utf-8') as f:
                    history = json.load(f)
            except:
                history = []
        history.append(entry)
        # 只保留最近30条
        history = history[-30:]
        with open(UNDO_HISTORY, 'w', encoding='utf-8') as f:
            json.dump(history, f, ensure_ascii=False, indent=1)

    # =================== COM内部操作 ===================

    def _detect_cols(self, ws, max_col):
        """从表头行自动检测所有关键列位置，适配不同排期文件布局
        不同排期文件列顺序不同（如15760有额外的系统货号/ITEM#列），
        必须通过表头关键词检测，不能硬编码列号
        支持繁体中文表头（如出貨→出货、驗貨→验货等）"""
        cols = {}
        mc = min(max_col, 100)
        # 扫描前5行找表头（有些文件表头在第4行）
        for r in range(1, 6):
            for c in range(1, mc + 1):
                try:
                    v = ws.Cells(r, c).Value
                    if not v:
                        continue
                    vl_raw = str(v).strip()
                    vl = _t2s(vl_raw)  # 繁体→简体
                    vlu = vl.upper().replace(' ', '')

                    # 接单日期（含"首办"变体）
                    if ('接单' in vl or '首办' in vl) and 'po_date' not in cols:
                        cols['po_date'] = c
                    # 客户名（排除客户PO，支持"第二方"/"第三方"变体）
                    elif ('客户名' in vl or '第三方' in vl or '第二方' in vl) and 'PO' not in vlu and 'customer' not in cols:
                        cols['customer'] = c
                    elif vl == '客户' and 'customer' not in cols:
                        cols['customer'] = c
                    # 走货国
                    elif ('走货' in vl or vl in ('国家', '目的国')) and 'destination' not in cols:
                        cols['destination'] = c
                    # PO号（排除客户PO/小PO）
                    elif 'PO' in vlu and ('号' in vl or '#' in vl or vl == 'PO') and '客户' not in vl and '客' not in vl.split('PO')[0][-1:] and '小' not in vl and '数量' not in vl and 'po_number' not in cols:
                        cols['po_number'] = c
                    # 客户PO / 小PO / 客PO（含"第三方客PO"变体）
                    # 注意："客PO期"是出货日期列，不是客户PO号列
                    elif (('客户' in vl and 'PO' in vlu) or ('小' in vl and 'PO' in vlu) or
                          ('客PO' in vl_raw or '客PO' in vl) or
                          vl in ('小PO号', '小PO', '客PO号', '客PO',
                                 '第三方客户PO NO#', '第三方客PO NO#')) and '期' not in vl and 'customer_po' not in cols:
                        cols['customer_po'] = c
                    # SKU（精确匹配及变体，含"SKU号"）
                    elif ('SKU' in vlu and vlu not in ('SKUCODE',) and
                          'sku' not in cols and 'ITEM' not in vlu):
                        cols['sku'] = c
                    # 系统货号（标记但不写入）
                    elif ('系统' in vl or vlu in ('SYSTEMCODE', 'SYSTEMNO')) and 'system_code' not in cols:
                        cols['system_code'] = c
                    # ITEM#/ITEMS/货号
                    elif 'items' not in cols and (
                        ('ITEM' in vlu and ('#' in vl or vlu.endswith('ITEM'))) or
                        vlu in ('ITEMS', 'ITEM', 'ITEMCODE', 'ITEM#')
                    ):
                        cols['items'] = c
                    elif vl == '货号' and 'items' not in cols:
                        cols['items'] = c
                    # 中文名/品名/产品名/名称
                    elif 'product_name' not in cols and (
                        '中文' in vl or '品名' in vl or vl == '名称' or
                        ('产品' in vl and ('名' in vl or '描述' in vl))
                    ):
                        cols['product_name'] = c
                    # PO数量/数量
                    elif 'qty' not in cols and (
                        ('数量' in vl and '合计' not in vl and '计划' not in vl and
                         '箱' not in vl and '外' not in vl) or
                        vlu in ('QTY', 'QTYPCS', 'PO数量')
                    ):
                        cols['qty'] = c
                    # 内箱
                    elif '内箱' in vl and 'inner_box' not in cols:
                        cols['inner_box'] = c
                    # 外箱（排除"外箱贴纸"等非数量列）
                    elif ('外箱' in vl or ('装箱' in vl and '内箱' not in vl)) and '贴纸' not in vl and 'outer_box' not in cols:
                        cols['outer_box'] = c
                    # 总箱/箱数
                    elif ('总箱' in vl or vl == '箱数') and 'total_box' not in cols:
                        cols['total_box'] = c
                    # 卡板/柜
                    elif '卡板' in vl and 'pallets' not in cols:
                        cols['pallets'] = c
                    # 出货期/出货日期/出期/走货日期（排除"计算走货"）
                    elif ('出货' in vl or '出期' in vl or
                          ('走货' in vl and '计算' not in vl)) and 'ship_date' not in cols:
                        cols['ship_date'] = c
                    # 客PO期（部分排期有独立的客PO期列，填同样的出货日期）
                    elif ('客PO' in vl and '期' in vl) and 'cpo_date' not in cols:
                        cols['cpo_date'] = c
                    # 验货期/验货日期/计划验货期
                    elif '验货' in vl and 'inspection' not in cols:
                        cols['inspection'] = c
                    # 业务/跟单
                    elif ('业务' in vl or '跟单' in vl) and 'from_person' not in cols:
                        cols['from_person'] = c
                    # 单价
                    elif '单价' in vl and 'price' not in cols:
                        cols['price'] = c
                    # 金额
                    elif '金额' in vl and 'total_usd' not in cols:
                        cols['total_usd'] = c
                    # 备注
                    elif vl in ('备注', '备注专栏', 'Remark', 'REMARK') and 'remark' not in cols:
                        cols['remark'] = c
                    # 条码
                    elif ('条码' in vl or 'BARCODE' in vlu or 'UPC' in vlu or 'EAN' in vlu) and 'barcode' not in cols:
                        cols['barcode'] = c
                except:
                    pass
        # 兜底：如果ship_date未检测到但cpo_date存在，用cpo_date作为ship_date
        if 'ship_date' not in cols and 'cpo_date' in cols:
            cols['ship_date'] = cols['cpo_date']
        logging.debug(f"[_detect_cols] 检测到列: {cols}")
        return cols

    def _do_new_com(self, ws, ref_row, max_col, header, ln, start_after=0):
        """通过COM插入新行 — 全部列位置通过表头自动检测，适配不同排期文件布局
        流程：检测列位置 → 按出货期找插入位 → 插入空行 → 复制参考行全部内容
             → 仅覆盖订单特定字段 → 清除系统列
        start_after: 传给_insert_pos_com，同批次多条目时保持插入顺序"""
        # 出货日期：优先header的ship_date，兜底用行数据的delivery
        ship_str = header.get('ship_date', '') or ln.get('delivery', '')
        ship_dt = _parse_date(ship_str)
        mc = min(max_col, 100)

        # 1. 先检测所有列位置（在插入行之前，表头不受影响）
        dcols = self._detect_cols(ws, mc)

        # 2. 用检测到的出货期列查找插入位置（不同文件出货期列不同）
        ship_col = dcols.get('ship_date', 13)
        pos = self._insert_pos_com(ws, ship_dt, col=ship_col, start_after=start_after)

        # 3. 插入空行
        ws.Rows(pos).Insert()

        # 4. 插入后ref_row可能移位
        actual_ref = ref_row + 1 if ref_row >= pos else ref_row

        # 5. 只复制参考行的格式（不复制值和公式，避免公式引用错乱影响其他行）
        # 使用范围复制（仅mc列）而非整行复制，避免16384列超大sheet导致COM崩溃
        try:
            src_rng = ws.Range(ws.Cells(actual_ref, 1), ws.Cells(actual_ref, mc))
            src_rng.Copy()
            dst_rng = ws.Range(ws.Cells(pos, 1), ws.Cells(pos, mc))
            dst_rng.PasteSpecial(Paste=-4122)  # xlPasteFormats only
        except Exception as _fmt_err:
            logging.warning(f"[新录入] 范围格式复制失败，尝试整行: {_fmt_err}")
            try:
                ws.Rows(actual_ref).Copy()
                ws.Rows(pos).PasteSpecial(Paste=-4122)
            except Exception as _row_err:
                logging.warning(f"[新录入] 整行格式复制也失败: {_row_err}")

        # 6. 清除剪贴板
        try:
            ws.Application.CutCopyMode = False
        except:
            pass

        # 7. 逐列复制值和公式（公式用FormulaR1C1自动调整相对引用，不影响其他行）
        for c in range(1, mc + 1):
            try:
                ref_cell = ws.Cells(actual_ref, c)
                if ref_cell.HasFormula:
                    # R1C1格式的公式自动适配新行位置
                    ws.Cells(pos, c).FormulaR1C1 = ref_cell.FormulaR1C1
                else:
                    v = ref_cell.Value
                    if v is not None:
                        ws.Cells(pos, c).Value = v
            except:
                pass

        # 7.5 公式修复：对计算列(总箱数/卡板/金额)，若参考行无公式则搜索附近行的公式
        calc_col_keys = ['total_box', 'pallets', 'total_usd']
        for ck in calc_col_keys:
            fc = dcols.get(ck)
            if not fc:
                continue
            try:
                if not ws.Cells(pos, fc).HasFormula:
                    # 向上搜索有公式的行（跳过新插入行本身）
                    for sr in range(pos - 1, max(3, pos - 50), -1):
                        try:
                            if ws.Cells(sr, fc).HasFormula:
                                ws.Cells(pos, fc).FormulaR1C1 = ws.Cells(sr, fc).FormulaR1C1
                                break
                        except:
                            pass
            except:
                pass

        # 7.8 清除非录入列（参考行复制了旧值，新订单不应继承）
        # A) 跟踪列：验货期和备注之间的列（第三方验货日期、验货结果、船发SO等）
        insp_c = dcols.get('inspection', 0)
        remark_c = dcols.get('remark') or self._note_col_com(ws, pos, mc)
        if insp_c and remark_c and remark_c > insp_c + 1:
            for c in range(insp_c + 1, remark_c):
                try:
                    if not ws.Cells(pos, c).HasFormula:
                        ws.Cells(pos, c).ClearContents()
                except:
                    pass
        # B) 系统列及后续：系统号、BOM状态、CA散大箱、Global散等
        # 注意：有些排期（如9543）系统货号列在数据列之前，必须保护已检测到的数据列
        sys_col = dcols.get('system_code')
        if sys_col:
            protected = set()
            for pk in ('po_date', 'customer', 'destination', 'po_number', 'customer_po',
                       'sku', 'items', 'product_name', 'qty', 'inner_box', 'outer_box',
                       'total_box', 'pallets', 'ship_date', 'inspection', 'from_person',
                       'price', 'total_usd', 'remark', 'barcode'):
                if pk in dcols:
                    protected.add(dcols[pk])
            for c in range(sys_col, mc + 1):
                if c in protected:
                    continue
                try:
                    if not ws.Cells(pos, c).HasFormula:
                        ws.Cells(pos, c).ClearContents()
                except:
                    pass

        # ===== 8. 仅覆盖订单特定字段（全部使用检测到的列位置）=====
        # 核心原则：产品固有属性（中文名、内箱、外箱、总箱公式）全部从参考行继承，
        #          只有订单特定字段从PDF覆盖
        po = header.get('po_number', '')

        # 接单日期（用序列号写入，避免pywin32时区偏移，NumberFormat保留参考行格式）
        po_dt = _parse_date(header.get('po_date', ''))
        if po_dt:
            ws.Cells(pos, dcols.get('po_date', 1)).Value = _date_serial(po_dt)

        # 客户名
        cust = header.get('customer', '')
        if cust:
            _sv_com(ws, pos, dcols.get('customer', 2), cust)

        # 走货国
        dest = header.get('destination_cn', '')
        if dest:
            _sv_com(ws, pos, dcols.get('destination', 3), dest)

        # PO号（先清除公式/旧值，再写入新值，确保不被参考行公式覆盖）
        po_col = dcols.get('po_number', 4)
        if po:
            try:
                ws.Cells(pos, po_col).ClearContents()
            except:
                pass
            _sv_com(ws, pos, po_col, int(po) if po.isdigit() else po)
            logging.info(f"[PO写入] row={pos}, col={po_col}, po={po}")

        # 客户PO（确保数字型PO号写为整数，避免小数问题）
        cpo = ln.get('customer_po', '')
        cpo_col = dcols.get('customer_po', 5)
        if cpo:
            try:
                cpo_v = int(float(cpo))
            except (ValueError, TypeError):
                cpo_v = str(cpo)
            _sv_com(ws, pos, cpo_col, cpo_v)
        else:
            # 客PO未提供时清空（不保留参考行的旧值）
            try:
                if not ws.Cells(pos, cpo_col).HasFormula:
                    ws.Cells(pos, cpo_col).ClearContents()
            except:
                pass

        # SKU (PO-line format)
        c_sku = dcols.get('sku', 6)
        line_no = ln.get('line_no', '')
        if po and line_no:
            _sv_com(ws, pos, c_sku, f"{po}-{line_no}")
        elif 'sku' in dcols:
            ref_sku_val = ''
            try:
                ref_sku_val = str(ws.Cells(actual_ref, c_sku).Value or '')
            except:
                pass
            if re.match(r'^\d{7,}-\d+$', ref_sku_val):
                pass
            elif ln.get('sku_spec') or ln.get('sku'):
                _sv_com(ws, pos, c_sku, ln.get('sku_spec', '') or ln.get('sku', ''))
        else:
            ref_sku_val = ''
            try:
                ref_sku_val = str(ws.Cells(actual_ref, c_sku).Value or '')
            except:
                pass
            if re.match(r'^\d{7,}-\d+$', ref_sku_val):
                pass
            elif ln.get('sku_spec') or ln.get('sku'):
                _sv_com(ws, pos, c_sku, ln.get('sku_spec', '') or ln.get('sku', ''))

        # ITEM#/货号 — 用PDF的完整sku_spec覆写（如"9296-S001"），大小写跟随排期已有条目
        item_col = dcols.get('items')
        sku_spec_val = ln.get('sku_spec', '') or ln.get('sku', '')
        if item_col and sku_spec_val:
            # 检查参考行的货号大小写，跟随排期已有格式
            try:
                ref_item_val = str(ws.Cells(actual_ref, item_col).Value or '').strip()
                if ref_item_val and ref_item_val.upper() == sku_spec_val.upper():
                    sku_spec_val = ref_item_val  # 使用参考行的大小写（如s001 vs S001）
                    logging.info(f"[货号大小写] 跟随参考行: '{ref_item_val}'")
            except:
                pass
            _sv_com(ws, pos, item_col, sku_spec_val)
            logging.info(f"[货号写入] row={pos}, col={item_col}, 货号={sku_spec_val}")
        # 产品名称/中文名 — 不覆写（从参考行同货号复制而来，绝不用PDF英文名）
        # 内箱/外箱 — 不覆写（从参考行同货号复制而来，产品固有属性）
        # 总箱数/卡板/金额 — 不覆写（从参考行复制公式，自动计算）
        # 系统货号 — 不填（明确清空，带"系统"的货号列不填）
        if 'system_code' in dcols:
            try:
                ws.Cells(pos, dcols['system_code']).ClearContents()
            except:
                pass

        # PO数量（唯一从PDF取的数量字段）
        qty = ln.get('qty', 0)
        if qty:
            _sv_com(ws, pos, dcols.get('qty', 9), qty)

        # 出货日期（用序列号写入，避免pywin32时区偏移）
        if ship_dt:
            ws.Cells(pos, ship_col).Value = _date_serial(ship_dt)
            # 客PO期 = 走货日期 = 出货日期（ZURU订单三者相同）
            cpo_date_col = dcols.get('cpo_date')
            if cpo_date_col and cpo_date_col != ship_col:
                ws.Cells(pos, cpo_date_col).Value = _date_serial(ship_dt)

        # 验货日期（出货期 - N天，避开周末）
        insp_col = dcols.get('inspection')
        if insp_col and ship_dt:
            try:
                sn = ws.Name if ws.Name else ''
                # 只有15746/河源Sheet才用2天减期，其他一律4天
                is_15746 = any(k in sn for k in ('15746', '河源'))
                days_before = 2 if is_15746 else 4
                insp_dt = ship_dt - timedelta(days=days_before)
                if insp_dt.weekday() == 6:  # Sunday → Friday
                    insp_dt -= timedelta(days=2)
                elif insp_dt.weekday() == 5:  # Saturday → Friday
                    insp_dt -= timedelta(days=1)
                ws.Cells(pos, insp_col).Value = _date_serial(insp_dt)
            except:
                pass

        # 条码（仅检测到条码列且有数据时写入）
        barcode = ln.get('barcode', '')
        if barcode and 'barcode' in dcols:
            _sv_com(ws, pos, dcols['barcode'], barcode)

        # 备注（只保留当前item相关的包装信息，不包含其他item的信息）
        item_num = _sku_key(ln.get('sku_spec', '') or ln.get('sku', ''))
        note = self._build_note(header, item_num=item_num)
        nc = dcols.get('remark') or self._note_col_com(ws, pos, mc)
        if note and nc:
            ws.Cells(pos, nc).Value = note

        # 跟单人/业务（仅检测到时写入）
        if 'from_person' in dcols:
            fp = header.get('from_person', '')
            if fp:
                _sv_com(ws, pos, dcols['from_person'], fp.split('/')[0].strip())

        # 单价USD（仅检测到且有数据时写入）
        if 'price' in dcols and ln.get('price', 0) > 0:
            _sv_com(ws, pos, dcols['price'], ln['price'])

        # 9. 新录入行：蓝色填充 + 黑色字体（所有新录入统一格式）
        try:
            new_rng = ws.Range(ws.Cells(pos, 1), ws.Cells(pos, mc))
            new_rng.Interior.Color = BLUE_COM
            new_rng.Font.Color = 0  # 黑色字体
        except Exception as e:
            logging.warning(f"[新录入] 格式设置失败: {e}")

        # 10. 验证：检查所有必填字段是否成功写入
        _warnings = []
        _check_fields = [
            ('items', '货号(ITEM#)'),
            ('product_name', '中文名/货名'),
            ('qty', 'PO数量'),
            ('inner_box', '内箱装箱数'),
            ('outer_box', '外箱装箱数'),
            ('ship_date', '出货日期/走货日期'),
            ('customer', '客户名'),
            ('destination', '国家'),
            ('po_number', 'PO号'),
        ]
        for col_key, label in _check_fields:
            if col_key in dcols:
                try:
                    cv = ws.Cells(pos, dcols[col_key]).Value
                    if cv is None or (isinstance(cv, str) and not cv.strip()):
                        _warnings.append(f"{label}为空(col={dcols[col_key]})")
                except:
                    pass
            else:
                _warnings.append(f"{label}列未检测到")
        if _warnings:
            logging.warning(f"[新录入验证] row={pos}, item={ln.get('sku','')}: {', '.join(_warnings)}")

        return pos, _warnings

    def _do_modify_com(self, ws, row, max_col, changes):
        """通过COM修改指定单元格（只更新值，不改变格式/颜色）"""
        for cl, nv in changes.items():
            cn = _col_num(cl)
            cell = ws.Cells(row, cn)

            # 设置值（按值内容检测日期，不限制列号）
            is_date = False
            if isinstance(nv, str) and re.match(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', str(nv)):
                is_date = True
            if is_date:
                try:
                    dt = _parse_date(nv)
                    if dt:
                        cell.Value = _date_serial(dt)
                    else:
                        cell.Value = nv
                except:
                    cell.Value = nv
            else:
                try:
                    cell.Value = int(nv) if str(nv).isdigit() else float(nv)
                except:
                    cell.Value = nv

        # 修改单不改变行格式（保持原有底色和字体颜色）

    def _do_cancel_com(self, wb, ws, row, max_col):
        """通过COM取消行：复制到取消订单Sheet → 标红+蓝 → 删原行
        注意：总排期文件只删原行，不复制到取消Sheet"""
        mc = min(max_col, 100)

        # 判断是否为总排期文件（总排期只删不复制）
        wb_name = wb.Name if wb.Name else ''
        is_summary = '总' in wb_name

        if not is_summary:
            # 查找或创建取消订单Sheet
            cancel_ws = None
            for i in range(1, wb.Sheets.Count + 1):
                if '取消' in wb.Sheets(i).Name:
                    cancel_ws = wb.Sheets(i)
                    break
            if cancel_ws is None:
                cancel_ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
                cancel_ws.Name = '取消订单'

            # 找取消Sheet下一空行
            try:
                cr = cancel_ws.Cells(cancel_ws.Rows.Count, 1).End(-4162).Row + 1  # xlUp
                if cr < 1:
                    cr = 1
            except:
                cr = 1

            # 复制整行到取消Sheet（保留原格式，超范围时自动缩小）
            try:
                src_rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, mc))
                src_rng.Copy(Destination=cancel_ws.Range(cancel_ws.Cells(cr, 1),
                                                          cancel_ws.Cells(cr, mc)))
            except Exception as e:
                logging.warning(f"[取消] 复制{mc}列失败: {e}, 回退到50列")
                mc = 50
                src_rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, mc))
                src_rng.Copy(Destination=cancel_ws.Range(cancel_ws.Cells(cr, 1),
                                                          cancel_ws.Cells(cr, mc)))

            # 取消Sheet中：红色字体 + 浅蓝底色
            dest_rng = cancel_ws.Range(cancel_ws.Cells(cr, 1), cancel_ws.Cells(cr, mc))
            dest_rng.Font.Color = RED_COM
            dest_rng.Interior.Color = BLUE_COM

            # 清剪贴板
            try:
                ws.Application.CutCopyMode = False
            except:
                pass

        # 删除原行
        ws.Rows(row).Delete()

    def _insert_pos_com(self, ws, ship_dt, col=13, start_after=0):
        """通过COM查找按出货日期的插入位置
        start_after: 从此行之后开始搜索（同批次多条目时保持插入顺序）
        使用 >= 比较: 新条目插在同日期已有条目之前（最新PO排在最上面）"""
        start_row = max(4, start_after + 1)
        last = start_row - 1
        try:
            used_rows = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        except:
            used_rows = 100

        # 若无出货日期，直接插到末尾
        if ship_dt is None:
            for r in range(start_row, min(used_rows + 1, 5000)):
                v = ws.Cells(r, col).Value
                if v is not None and hasattr(v, 'year'):
                    last = r
            return last + 1

        for r in range(start_row, min(used_rows + 1, 5000)):
            v = ws.Cells(r, col).Value
            if v is None:
                continue
            try:
                if hasattr(v, 'year'):
                    dt = datetime(v.year, v.month, v.day)
                    last = r
                    if dt >= ship_dt:
                        return r
            except:
                continue
        return last + 1

    def _note_col_com(self, ws, row, max_col):
        """通过COM查找备注列"""
        for c in [23, 22, 26]:
            for r in range(max(2, row - 10), row):
                try:
                    v = ws.Cells(r, c).Value
                    if v and ('日期码' in str(v) or 'Remark' in str(v)):
                        return c
                except:
                    pass
        return 23

    # =================== 兼容旧接口（单条操作，也用COM）===================

    def enter_new(self, sched, header, lines):
        local = os.path.join(DESKTOP, 'schedule_temp.xlsx')
        shutil.copy2(sched['file'], local)
        app = self._com_app()
        try:
            wb = app.Workbooks.Open(os.path.abspath(local))
            ws = wb.Sheets(sched['sheet'])
            ref = sched['ref']
            mc = min(sched.get('mcol', 50), 50)
            inserted_positions = []
            last_insert_pos = 0  # 同批次保持插入顺序
            for ln in lines:
                # 累计偏移：之前插入的行会使ref下移
                adj_ref = ref
                for p in inserted_positions:
                    if p <= adj_ref:
                        adj_ref += 1
                pos, _w = self._do_new_com(ws, adj_ref, mc, header, ln,
                                       start_after=last_insert_pos)
                inserted_positions.append(pos)
                last_insert_pos = pos
            wb.Save()
            wb.Close(False)
        finally:
            self._com_quit(app)
        return {'ok': True, 'local': local, 'z': sched['file'],
                'msg': f'已录入{len(lines)}行'}

    def modify(self, record, changes):
        local = os.path.join(DESKTOP, 'schedule_temp.xlsx')
        shutil.copy2(record['file'], local)
        app = self._com_app()
        try:
            wb = app.Workbooks.Open(os.path.abspath(local))
            ws = wb.Sheets(record['sheet'])
            mc = min(ws.UsedRange.Columns.Count + ws.UsedRange.Column, 100)
            self._do_modify_com(ws, record['row'], mc, changes)
            wb.Save()
            wb.Close(False)
        finally:
            self._com_quit(app)
        return {'ok': True, 'local': local, 'z': record['file'],
                'msg': f"第{record['row']}行已修改，修改字段已标红"}

    def cancel(self, record):
        local = os.path.join(DESKTOP, 'schedule_temp.xlsx')
        shutil.copy2(record['file'], local)
        app = self._com_app()
        try:
            wb = app.Workbooks.Open(os.path.abspath(local))
            ws = wb.Sheets(record['sheet'])
            mc = min(ws.UsedRange.Columns.Count + ws.UsedRange.Column, 100)
            self._do_cancel_com(wb, ws, record['row'], mc)
            wb.Save()
            wb.Close(False)
        finally:
            self._com_quit(app)
        return {'ok': True, 'local': local, 'z': record['file'],
                'msg': f"第{record['row']}行已移至取消订单Sheet"}

    def save_z(self, local, z):
        self._try_save_z(local, z)
        return {'ok': True, 'msg': '已保存到Z盘'}

    def _try_save_z(self, local, z):
        try:
            with open(z, 'r+b'):
                pass
        except:
            raise PermissionError('Z盘文件被占用或只读')
        shutil.copy2(local, z)

    # =================== 重试保存 ===================

    def retry_save(self, items):
        ok = []
        still_failed = []
        for item in items:
            try:
                self._try_save_z(item['local'], item['z'])
                ok.append(item)
            except:
                still_failed.append(item)
        return {'ok': ok, 'failed': still_failed}

    # =================== 备份系统 ===================

    def create_backup(self, modified_files=None):
        today = date.today()
        monday = today - timedelta(days=today.weekday())
        saturday = monday + timedelta(days=5)
        folder_name = f"排期备份 {monday.month}.{monday.day}-{saturday.month}.{saturday.day}"
        backup_dir = os.path.join(DESKTOP, folder_name)
        os.makedirs(backup_dir, exist_ok=True)

        date_suffix = f"{today.month}.{today.day}"
        backed_up = []
        files_to_backup = modified_files or self._list_xlsx()

        for fp in files_to_backup:
            fname = os.path.basename(fp)
            base, ext = os.path.splitext(fname)
            new_name = f"{base} {date_suffix}{ext}"
            dest = os.path.join(backup_dir, new_name)
            counter = 0
            while os.path.exists(dest):
                counter += 1
                new_name = f"{base} {date_suffix}-{counter}{ext}"
                dest = os.path.join(backup_dir, new_name)
            try:
                shutil.copy2(fp, dest)
                backed_up.append(new_name)
            except Exception as e:
                backed_up.append(f"{fname} (备份失败: {e})")

        return {'ok': True, 'folder': backup_dir, 'files': backed_up,
                'msg': f'已备份{len(backed_up)}个文件到 {folder_name}'}

    # =================== 历史记录 ===================

    @staticmethod
    def add_history(po, action, detail, files=''):
        os.makedirs(DATA_DIR, exist_ok=True)
        records = []
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    records = json.load(f)
            except:
                records = []
        records.append({
            'time': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'po': str(po), 'action': action,
            'detail': detail, 'files': files
        })
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=1)

    @staticmethod
    def get_history():
        if not os.path.exists(HISTORY_FILE):
            return []
        try:
            with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []

    @staticmethod
    def export_history_excel():
        """将历史记录导出为Excel文件，返回文件路径"""
        records = ExcelHandler.get_history()
        if not records:
            return None
        export_dir = os.path.join(DATA_DIR, 'exports')
        os.makedirs(export_dir, exist_ok=True)
        fname = f"操作历史_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        fpath = os.path.join(export_dir, fname)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '操作历史'
        # 表头
        headers = ['时间', '操作类型', 'PO号', '详情', '文件']
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = openpyxl.styles.Font(bold=True)
        # 数据
        for idx, rec in enumerate(reversed(records), 2):
            ws.cell(row=idx, column=1, value=rec.get('time', ''))
            ws.cell(row=idx, column=2, value=rec.get('action', ''))
            ws.cell(row=idx, column=3, value=rec.get('po', ''))
            ws.cell(row=idx, column=4, value=rec.get('detail', ''))
            ws.cell(row=idx, column=5, value=rec.get('files', ''))
        # 自动列宽
        for col in ws.columns:
            max_len = max(len(str(c.value or '')) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
        wb.save(fpath)
        return fpath

    # =================== 文件占用检测 ===================

    def check_all_file_status(self):
        """扫描所有排期文件的锁定/占用状态"""
        results = []
        if not os.path.isdir(self.z_path):
            return results

        # 先收集所有~$锁文件
        lock_files = {}
        for item in os.listdir(self.z_path):
            if item.startswith('~$') and item.endswith('.xlsx'):
                # ~$替换了原文件名前2个字符
                suffix = item[2:]
                lock_path = os.path.join(self.z_path, item)
                user = self._read_lock_user(lock_path)
                lock_files[suffix] = user

        for fp in self._list_xlsx():
            fname = os.path.basename(fp)
            suffix = fname[2:] if len(fname) > 2 else fname
            status = 'available'
            user = ''
            lock_type = ''

            # 方法1：检查~$锁文件（最可靠，能识别使用者）
            if suffix in lock_files:
                status = 'locked'
                user = lock_files[suffix]
                lock_type = '正在编辑'
            else:
                # 方法2：尝试写入权限测试
                try:
                    with open(fp, 'r+b'):
                        pass
                except PermissionError:
                    status = 'locked'
                    lock_type = '写入被拒'
                except:
                    pass

            results.append({
                'file': fp, 'fname': fname,
                'status': status, 'user': user, 'lock_type': lock_type
            })

        return results

    def _read_lock_user(self, lock_path):
        """从~$锁文件读取使用者用户名"""
        try:
            with open(lock_path, 'rb') as f:
                raw = f.read(200)
            if not raw:
                return '未知用户'

            # 尝试方法1：第一个字节是长度，后面是用户名
            try:
                name_len = raw[0]
                if 1 <= name_len <= 50:
                    # 尝试GBK（中文Windows常用）
                    name = raw[1:1+name_len*2].decode('utf-16-le', errors='ignore')
                    name = name.split('\x00')[0].strip()
                    if name and len(name) >= 1:
                        return name
            except:
                pass

            # 尝试方法2：直接UTF-16LE
            try:
                text = raw[:54].decode('utf-16-le', errors='ignore')
                name = text.split('\x00')[0].strip()
                if name and len(name) >= 2:
                    return name
            except:
                pass

            # 尝试方法3：GBK/ASCII
            try:
                text = raw[:54].decode('gbk', errors='ignore')
                name = ''.join(c for c in text if c.isprintable()).strip()
                if name:
                    return name
            except:
                pass

            return '未知用户'
        except:
            return '未知用户'

    # =================== 待重试队列 ===================

    @staticmethod
    def get_pending_retries():
        if not os.path.exists(RETRY_FILE):
            return []
        try:
            with open(RETRY_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []

    @staticmethod
    def save_pending_retries(items):
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(RETRY_FILE, 'w', encoding='utf-8') as f:
            json.dump(items, f, ensure_ascii=False, indent=1)

    def auto_retry_pending(self):
        """自动重试所有待保存的文件"""
        items = self.get_pending_retries()
        if not items:
            return {'ok': [], 'failed': [], 'msg': '无待重试项'}

        ok = []
        still_failed = []
        for item in items:
            try:
                self._try_save_z(item['local'], item['z'])
                ok.append(item)
                self.add_history(
                    item.get('po', ''), 'auto_retry',
                    f"自动重试成功: {item['file']}", item['file']
                )
            except:
                item['retries'] = item.get('retries', 0) + 1
                item['last_retry'] = datetime.now().strftime('%H:%M')
                still_failed.append(item)

        self.save_pending_retries(still_failed)
        return {'ok': ok, 'failed': still_failed,
                'msg': f'成功{len(ok)}个，仍有{len(still_failed)}个待重试'}

    def delete_entries_com(self, entries):
        """删除指定的排期行（支持撤销）
        entries: [{file: z盘路径, sheet: 工作表名, row: 行号, sku: 显示用}]
        返回: {ok: True/False, deleted: [...], failed: [...], undo_id: ...}
        """
        if not entries:
            return {'ok': False, 'error': '没有要删除的条目'}

        os.makedirs(UNDO_DIR, exist_ok=True)
        batch_id = datetime.now().strftime('%Y%m%d-%H%M%S') + '_del'

        # 按文件分组，行号从大到小排列（先删大行号避免偏移）
        file_groups = {}
        for e in entries:
            fkey = e['file']
            if fkey not in file_groups:
                file_groups[fkey] = []
            file_groups[fkey].append(e)
        for fk in file_groups:
            file_groups[fk].sort(key=lambda x: x.get('row', 0), reverse=True)

        deleted = []
        failed = []
        undo_files = []
        app = None
        try:
            app = self._com_app()
            for fkey, group in file_groups.items():
                fname = os.path.basename(fkey)
                # 备份
                undo_fp = os.path.join(UNDO_DIR, f"{batch_id}_{fname}")
                try:
                    import shutil
                    shutil.copy2(fkey, undo_fp)
                except Exception as e:
                    for g in group:
                        failed.append({'sku': g.get('sku', ''), 'reason': f'备份失败: {e}'})
                    continue

                try:
                    wb = app.Workbooks.Open(os.path.abspath(fkey))
                    for e in group:
                        sn = e.get('sheet', '')
                        rn = e.get('row', 0)
                        try:
                            ws = wb.Sheets(sn)
                            # 记录被删行的内容（用于反馈）
                            row_info = {}
                            for ci in range(1, min(20, ws.UsedRange.Columns.Count + 1)):
                                v = ws.Cells(rn, ci).Value
                                if v is not None:
                                    row_info[f'col{ci}'] = str(v)[:50]
                            ws.Rows(rn).Delete()
                            deleted.append({
                                'sku': e.get('sku', ''),
                                'file': fname,
                                'sheet': sn,
                                'row': rn,
                                'row_info': row_info
                            })
                        except Exception as ex:
                            failed.append({'sku': e.get('sku', ''), 'reason': str(ex)[:100]})
                    wb.Save()
                    wb.Close(False)
                    undo_files.append({'name': fname, 'backup': undo_fp, 'z_path': fkey})
                except Exception as ex:
                    for g in group:
                        if not any(f.get('sku') == g.get('sku') for f in failed):
                            failed.append({'sku': g.get('sku', ''), 'reason': str(ex)[:100]})
        finally:
            if app:
                try:
                    app.Quit()
                except:
                    pass

        # 保存撤销记录
        if deleted:
            ops = [{'type': 'delete', 'sku': d['sku'],
                    'detail': f"删除 {d['sheet']} 行{d['row']}"} for d in deleted]
            self._save_undo_entry({
                'id': batch_id,
                'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'operations': ops,
                'files': undo_files,
                'label': f"删除 {len(deleted)} 行",
            })

        return {
            'ok': len(deleted) > 0,
            'deleted': deleted,
            'failed': failed,
            'undo_id': batch_id if deleted else None,
            'msg': f'已删除 {len(deleted)} 行' + (f'，{len(failed)} 行失败' if failed else '')
        }

    def reentry_batch(self, orders):
        """重新入单（删除旧行+重新写入），返回与batch_process相同格式的详细结果"""
        return self.batch_process(orders)

    # =================== 定时重试 ===================

    SCHEDULED_FILE = os.path.join(DATA_DIR, 'scheduled_retries.json')

    @staticmethod
    def get_scheduled_retries():
        fp = os.path.join(DATA_DIR, 'scheduled_retries.json')
        if os.path.exists(fp):
            try:
                with open(fp, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return []

    @staticmethod
    def save_scheduled_retries(items):
        os.makedirs(DATA_DIR, exist_ok=True)
        fp = os.path.join(DATA_DIR, 'scheduled_retries.json')
        with open(fp, 'w', encoding='utf-8') as f:
            json.dump(items, f, ensure_ascii=False, indent=1)

    # =================== 工具 ===================

    def _build_note(self, h, item_num=''):
        """构建备注文本，若提供item_num则只保留该item相关的包装信息"""
        p = []
        if h.get('tracking_code'):
            p.append(h['tracking_code'])
        if h.get('packaging_info'):
            pkg = h['packaging_info']
            if item_num and len(item_num) >= 3:
                # 过滤packaging_info：只保留当前item相关行 + 通用行
                filtered = []
                for line in pkg.split('\n'):
                    line_s = line.strip()
                    if not line_s:
                        continue
                    # 检查行是否以某个item编号开头（数字开头的行视为item特定行）
                    m = re.match(r'^(\d{3,})', line_s)
                    if m:
                        line_item = m.group(1)
                        # 只保留当前item的行（item_num前缀匹配）
                        if item_num in line_item or line_item in item_num:
                            filtered.append(line_s)
                    else:
                        # 不以数字开头的通用行（如"出加拿大"、"有客供价格贴"等）保留
                        filtered.append(line_s)
                pkg = '\n'.join(filtered)
            if pkg.strip():
                p.append(pkg.strip())
        if h.get('remark'):
            p.append(h['remark'])
        return '\n'.join(p) if p else ''

    # =================== 撤回操作 ===================

    def undo_selected(self, batch_ids):
        """撤回指定批次的操作：根据batch_id恢复对应备份文件"""
        history = self._load_undo_history()
        if not history:
            return {'error': '没有可撤回的操作'}

        # 找到要撤回的批次
        to_undo = [h for h in history if h['id'] in batch_ids]
        if not to_undo:
            return {'error': '未找到指定的操作记录'}

        restored = []
        failed = []
        undone_ids = []

        for entry in to_undo:
            for finfo in entry.get('files', []):
                backup = finfo.get('backup', '')
                z_path = finfo.get('z_path', '')
                fname = finfo.get('name', '')

                if not backup or not os.path.exists(backup):
                    # 兼容旧格式：尝试不带前缀的备份
                    old_backup = os.path.join(UNDO_DIR, fname)
                    if os.path.exists(old_backup):
                        backup = old_backup
                    else:
                        failed.append({'file': fname, 'reason': '备份文件不存在'})
                        continue

                if not z_path or not os.path.exists(z_path):
                    # 尝试从z_path构建
                    z_path = os.path.join(self.z_path, fname)
                    if not os.path.exists(z_path):
                        failed.append({'file': fname, 'reason': '目标文件不存在'})
                        continue

                try:
                    with open(z_path, 'r+b'):
                        pass
                    shutil.copy2(backup, z_path)
                    restored.append(fname)
                except:
                    failed.append({'file': fname, 'reason': '文件被占用'})

            if not any(f['file'] == fi.get('name') for fi in entry.get('files', []) for f in failed):
                undone_ids.append(entry['id'])

        # 从历史中移除已成功撤回的批次，并清理备份文件
        if undone_ids:
            new_history = []
            for h in history:
                if h['id'] in undone_ids:
                    # 清理备份文件
                    for finfo in h.get('files', []):
                        bp = finfo.get('backup', '')
                        if bp and os.path.exists(bp):
                            try:
                                os.remove(bp)
                            except:
                                pass
                else:
                    new_history.append(h)
            self._write_undo_history(new_history)

        return {
            'ok': True, 'restored': restored, 'failed': failed,
            'undone_ids': undone_ids,
            'msg': f'已撤回 {len(restored)} 个文件' +
                   (f'，{len(failed)} 个失败' if failed else '')
        }

    def undo_last_batch(self):
        """兼容旧接口：撤回最近一次操作"""
        history = self._load_undo_history()
        if history:
            return self.undo_selected([history[-1]['id']])
        # 兼容旧格式undo目录
        if not os.path.isdir(UNDO_DIR):
            return {'error': '没有可撤回的操作'}
        undo_files = [f for f in os.listdir(UNDO_DIR)
                      if f.endswith('.xlsx') and not f.startswith('~$')
                      and not re.match(r'\d{8}-\d{6}_', f)]
        if not undo_files:
            return {'error': '没有可撤回的操作'}
        restored = []
        failed = []
        for fname in undo_files:
            undo_fp = os.path.join(UNDO_DIR, fname)
            z_fp = os.path.join(self.z_path, fname)
            if not os.path.exists(z_fp):
                failed.append({'file': fname, 'reason': '目标文件不存在'})
                continue
            try:
                with open(z_fp, 'r+b'):
                    pass
                shutil.copy2(undo_fp, z_fp)
                restored.append(fname)
            except:
                failed.append({'file': fname, 'reason': '文件被占用'})
        if restored and not failed:
            for f in os.listdir(UNDO_DIR):
                try:
                    os.remove(os.path.join(UNDO_DIR, f))
                except:
                    pass
        return {'ok': True, 'restored': restored, 'failed': failed,
                'msg': f'已撤回 {len(restored)} 个文件' + (f'，{len(failed)} 个失败' if failed else '')}

    def _load_undo_history(self):
        if not os.path.exists(UNDO_HISTORY):
            return []
        try:
            with open(UNDO_HISTORY, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []

    def _write_undo_history(self, history):
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(UNDO_HISTORY, 'w', encoding='utf-8') as f:
            json.dump(history, f, ensure_ascii=False, indent=1)

    @staticmethod
    def get_undo_info():
        """获取可撤回的操作列表（含历史详情）"""
        history = []
        if os.path.exists(UNDO_HISTORY):
            try:
                with open(UNDO_HISTORY, 'r', encoding='utf-8') as f:
                    history = json.load(f)
            except:
                pass

        # 检查每个批次的备份文件是否仍然存在
        valid = []
        for h in history:
            has_backup = False
            for finfo in h.get('files', []):
                bp = finfo.get('backup', '')
                if bp and os.path.exists(bp):
                    has_backup = True
                    break
            if has_backup:
                valid.append(h)

        if not valid:
            # 兼容旧格式
            if os.path.isdir(UNDO_DIR):
                old_files = [f for f in os.listdir(UNDO_DIR)
                             if f.endswith('.xlsx') and not f.startswith('~$')
                             and not re.match(r'\d{8}-\d{6}_', f)]
                if old_files:
                    latest = max(os.path.getmtime(os.path.join(UNDO_DIR, f)) for f in old_files)
                    return {
                        'available': True,
                        'batches': [{
                            'id': 'legacy',
                            'time': datetime.fromtimestamp(latest).strftime('%Y-%m-%d %H:%M:%S'),
                            'label': f'上次操作（{len(old_files)}个文件）',
                            'operations': [],
                            'files': [{'name': f} for f in old_files],
                        }],
                        'count': 1,
                        # 兼容旧UI
                        'files': old_files,
                        'time': datetime.fromtimestamp(latest).strftime('%Y-%m-%d %H:%M'),
                    }
            return {'available': False, 'batches': [], 'count': 0}

        return {
            'available': True,
            'batches': list(reversed(valid)),  # 最新的在前
            'count': len(valid),
            # 兼容旧UI
            'files': [f['name'] for h in valid for f in h.get('files', [])],
            'time': valid[-1]['time'][:16] if valid else '',
        }

    # =================== 总排期操作 ===================

    def find_master_schedule(self):
        """查找总排期文件"""
        for fp in self._list_xlsx():
            fn = os.path.basename(fp)
            if '总' in fn:
                return fp
        return None

    def scan_yellow_rows(self, use_cache=True, progress_callback=None):
        """扫描所有分排期文件中的黄色填充行（带缓存 + 进度回调）"""
        global _yellow_cache
        results = []
        files = [fp for fp in self._list_xlsx() if '总' not in os.path.basename(fp)]
        total = len(files)

        for idx, fp in enumerate(files):
            fn = os.path.basename(fp)
            if progress_callback:
                progress_callback(idx + 1, total, fn)

            # 缓存检查：文件未修改则复用
            try:
                mtime = os.path.getmtime(fp)
            except:
                continue
            if use_cache and fp in _yellow_cache and _yellow_cache[fp]['mtime'] == mtime:
                results.extend(_yellow_cache[fp]['rows'])
                continue

            # 需要扫描此文件
            file_rows = []
            try:
                wb = openpyxl.load_workbook(fp, data_only=True)
                for sn in wb.sheetnames:
                    sn_lower = sn.lower()
                    if '取消' in sn or '明细' in sn or 'ma' in sn_lower or '对应' in sn:
                        continue
                    ws = wb[sn]
                    row_count = 0
                    for row in ws.iter_rows(min_row=2, max_col=30):
                        row_count += 1
                        if row_count > 2000:
                            break
                        if not any(c.value for c in row[:6]):
                            continue
                        is_yellow = False
                        for c in row[:6]:
                            if c.value is not None:
                                is_yellow = _is_yellow_fill(c)
                                break
                        if not is_yellow:
                            is_yellow = _is_yellow_fill(row[0])
                        if is_yellow:
                            data = {}
                            for c in row:
                                cl = openpyxl.utils.get_column_letter(c.column)
                                v = c.value
                                if isinstance(v, datetime):
                                    v = v.strftime('%Y-%m-%d')
                                data[cl] = v
                            file_rows.append({
                                'file': fp, 'fname': fn, 'sheet': sn,
                                'row': row[0].row, 'data': data
                            })
                wb.close()
            except:
                continue
            # 更新缓存
            _yellow_cache[fp] = {'mtime': mtime, 'rows': file_rows}
            results.extend(file_rows)
        return results

    def _read_headers(self, fp, sheet_name=None):
        """读取文件表头行，返回 {列字母: 表头名称}"""
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            sn = sheet_name or wb.sheetnames[0]
            ws = wb[sn]
            headers = {}
            for row in ws.iter_rows(min_row=1, max_row=4, max_col=30):
                for cell in row:
                    if cell.value and str(cell.value).strip():
                        cl = openpyxl.utils.get_column_letter(cell.column)
                        if cl not in headers:
                            headers[cl] = str(cell.value).strip()
                if len(headers) >= 5:
                    break
            wb.close()
            return headers
        except:
            return {}

    def _build_column_mapping(self, src_headers, dst_headers):
        """构建 分排期列→总排期列 的映射"""
        dst_name_col = {v: k for k, v in dst_headers.items()}
        used_dst = set()
        mapping = {}

        # 特殊别名表
        ALIASES = {
            'ZURU PO NO#': 'PO号', 'PO NO.': 'PO号', 'PO NUMBER': 'PO号',
            'SKU': '系统货号', 'ITEM CODE': '系统货号',
            'ITEM#': '货号#', 'ITEM NO.': '货号#',
            '货品名称': '中文名', '品名': '中文名',
            '出货日期': '预计船期', '船期': '预计船期', '走货日期': '预计船期',
            '验货日期': '预计验货日期',
        }

        for src_col, src_name in src_headers.items():
            sn = src_name.strip()
            # 1. 精确匹配
            if sn in dst_name_col and dst_name_col[sn] not in used_dst:
                mapping[src_col] = dst_name_col[sn]
                used_dst.add(dst_name_col[sn])
                continue
            # 2. 别名匹配
            alias_target = ALIASES.get(sn, '')
            if alias_target and alias_target in dst_name_col and dst_name_col[alias_target] not in used_dst:
                mapping[src_col] = dst_name_col[alias_target]
                used_dst.add(dst_name_col[alias_target])
                continue
            # 3. 包含匹配
            for dn, dc in dst_name_col.items():
                if dc in used_dst:
                    continue
                if sn in dn or dn in sn:
                    mapping[src_col] = dc
                    used_dst.add(dc)
                    break

        return mapping

    def copy_to_master(self, yellow_rows=None):
        """将黄色填充行复制到总排期"""
        master_fp = self.find_master_schedule()
        if not master_fp:
            return {'error': '未找到总排期文件（文件名需包含"总"字）'}
        try:
            with open(master_fp, 'r+b'):
                pass
        except:
            return {'error': '总排期文件被占用或只读，无法写入'}

        if yellow_rows is None:
            yellow_rows = self.scan_yellow_rows()
        if not yellow_rows:
            return {'error': '未找到黄色填充的行'}

        master_headers = self._read_headers(master_fp)
        if not master_headers:
            return {'error': '无法读取总排期表头'}

        src_cache = {}
        map_cache = {}

        app = None
        try:
            app = self._com_app()
            wb = app.Workbooks.Open(os.path.abspath(master_fp))
            ws = wb.Sheets(1)

            try:
                last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                if last_row < 2:
                    last_row = 2
            except:
                last_row = 2

            insert_row = last_row + 1
            copied = 0
            mc = max(_col_num(c) for c in master_headers.keys()) if master_headers else 20

            for yr in yellow_rows:
                src_key = yr['file'] + '|' + yr['sheet']
                if src_key not in src_cache:
                    src_cache[src_key] = self._read_headers(yr['file'], yr['sheet'])
                    map_cache[src_key] = self._build_column_mapping(
                        src_cache[src_key], master_headers)

                col_map = map_cache[src_key]
                has_data = False

                for src_col, value in yr['data'].items():
                    if value is None:
                        continue
                    dst_col = col_map.get(src_col)
                    if not dst_col:
                        continue
                    dst_num = _col_num(dst_col)
                    cell = ws.Cells(insert_row, dst_num)
                    if isinstance(value, str) and re.match(r'\d{4}-\d{2}-\d{2}', value):
                        try:
                            dt = datetime.strptime(value, '%Y-%m-%d')
                            cell.Value = dt
                            cell.NumberFormat = 'yyyy/m/d'
                        except:
                            cell.Value = value
                    else:
                        cell.Value = value
                    has_data = True

                if has_data:
                    rng = ws.Range(ws.Cells(insert_row, 1), ws.Cells(insert_row, mc))
                    rng.Interior.Color = YELLOW_COM
                    insert_row += 1
                    copied += 1

            wb.Save()
            wb.Close(False)
        finally:
            self._com_quit(app)

        return {
            'ok': True, 'copied': copied,
            'master_file': os.path.basename(master_fp),
            'msg': f'已复制 {copied} 行到总排期（{os.path.basename(master_fp)}）'
        }

    def clear_master_yellow(self):
        """清除总排期中的黄色填充"""
        master_fp = self.find_master_schedule()
        if not master_fp:
            return {'error': '未找到总排期文件'}
        try:
            with open(master_fp, 'r+b'):
                pass
        except:
            return {'error': '总排期文件被占用或只读'}

        app = None
        try:
            app = self._com_app()
            wb = app.Workbooks.Open(os.path.abspath(master_fp))
            ws = wb.Sheets(1)

            try:
                last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            except:
                last_row = 100

            mc = min(ws.UsedRange.Columns.Count + ws.UsedRange.Column, 100)
            cleared = 0

            for r in range(2, last_row + 1):
                try:
                    color = ws.Cells(r, 1).Interior.Color
                    rc = color % 256
                    gc = (color // 256) % 256
                    bc = (color // 65536) % 256
                    if rc > 200 and gc > 180 and bc < 100:
                        rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, mc))
                        rng.Interior.Pattern = -4142  # xlNone
                        cleared += 1
                except:
                    continue

            wb.Save()
            wb.Close(False)
        finally:
            self._com_quit(app)

        return {
            'ok': True, 'cleared': cleared,
            'msg': f'已清除 {cleared} 行的黄色填充'
        }

    # =================== 排期文件列表（手动选择用）===================

    def list_schedule_files(self):
        """列出所有排期文件及其Sheet，供手动选择"""
        result = []
        for fp in self._list_xlsx():
            fn = os.path.basename(fp)
            if '总' in fn:
                continue
            sheets = []
            try:
                wb = openpyxl.load_workbook(fp, read_only=True)
                for sn in wb.sheetnames:
                    if '取消' not in sn and '明细' not in sn:
                        sheets.append(sn)
                wb.close()
            except:
                sheets = ['Sheet1']
            result.append({'file': fp, 'fname': fn, 'sheets': sheets})
        return result

    def manual_find_ref(self, filepath, sheet_name):
        """在指定文件+Sheet中查找参考行（最后一个有数据的行）"""
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb[sheet_name]
            ref = 2
            for row in ws.iter_rows(min_row=2, max_col=10):
                if any(c.value for c in row[:6]):
                    ref = row[0].row
            wb.close()
            return {'file': filepath, 'fname': os.path.basename(filepath),
                    'sheet': sheet_name, 'ref': ref, 'cnt': 0, 'mcol': 30}
        except Exception as e:
            return {'error': str(e)}


# =================== 邮件集成 ===================

class EmailHandler:
    """通过IMAP读取邮箱附件PDF"""

    def __init__(self, config):
        self.server = config.get('email_server', '')
        self.port = int(config.get('email_port', 993))
        self.user = config.get('email_user', '')
        self.password = config.get('email_password', '')
        self.ssl = config.get('email_ssl', True)

    def check_new_emails(self, folder='INBOX', limit=20):
        """检查邮箱中的新邮件，返回含PDF附件的邮件列表"""
        if not self.server or not self.user:
            return {'error': '邮箱未配置，请在设置页面配置IMAP信息'}
        import imaplib
        import email
        from email.header import decode_header

        results = []
        try:
            if self.ssl:
                mail = imaplib.IMAP4_SSL(self.server, self.port)
            else:
                mail = imaplib.IMAP4(self.server, self.port)
            mail.login(self.user, self.password)
            mail.select(folder)

            # 搜索未读邮件
            status, messages = mail.search(None, 'UNSEEN')
            if status != 'OK':
                mail.logout()
                return {'emails': [], 'msg': '无新邮件'}

            msg_ids = messages[0].split()[-limit:]  # 只取最近的
            for mid in reversed(msg_ids):
                status, msg_data = mail.fetch(mid, '(RFC822)')
                if status != 'OK':
                    continue
                msg = email.message_from_bytes(msg_data[0][1])

                # 解码主题
                subject = ''
                raw_subject = msg.get('Subject', '')
                if raw_subject:
                    decoded = decode_header(raw_subject)
                    subject = ''.join(
                        part.decode(enc or 'utf-8') if isinstance(part, bytes) else part
                        for part, enc in decoded
                    )

                from_addr = msg.get('From', '')
                date_str = msg.get('Date', '')

                # 查找PDF附件
                attachments = []
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    fname = part.get_filename()
                    if fname:
                        decoded_fname = decode_header(fname)
                        fname = ''.join(
                            p.decode(enc or 'utf-8') if isinstance(p, bytes) else p
                            for p, enc in decoded_fname
                        )
                        if fname.lower().endswith('.pdf'):
                            attachments.append({
                                'filename': fname,
                                'size': len(part.get_payload(decode=True) or b''),
                                'msg_id': mid.decode() if isinstance(mid, bytes) else str(mid),
                            })

                if attachments:
                    results.append({
                        'subject': subject, 'from': from_addr, 'date': date_str,
                        'msg_id': mid.decode() if isinstance(mid, bytes) else str(mid),
                        'attachments': attachments,
                    })

            mail.logout()
            return {'emails': results, 'count': len(results)}

        except Exception as e:
            return {'error': f'邮箱连接失败: {str(e)}'}

    def download_attachment(self, msg_id, filename, save_dir):
        """下载指定邮件的PDF附件"""
        if not self.server or not self.user:
            return None
        import imaplib
        import email
        from email.header import decode_header

        os.makedirs(save_dir, exist_ok=True)
        try:
            if self.ssl:
                mail = imaplib.IMAP4_SSL(self.server, self.port)
            else:
                mail = imaplib.IMAP4(self.server, self.port)
            mail.login(self.user, self.password)
            mail.select('INBOX')

            status, msg_data = mail.fetch(msg_id.encode() if isinstance(msg_id, str) else msg_id,
                                          '(RFC822)')
            if status != 'OK':
                mail.logout()
                return None

            msg = email.message_from_bytes(msg_data[0][1])
            for part in msg.walk():
                fname = part.get_filename()
                if fname:
                    decoded_fname = decode_header(fname)
                    fname = ''.join(
                        p.decode(enc or 'utf-8') if isinstance(p, bytes) else p
                        for p, enc in decoded_fname
                    )
                    if fname == filename:
                        content = part.get_payload(decode=True)
                        save_path = os.path.join(save_dir, fname)
                        with open(save_path, 'wb') as f:
                            f.write(content)
                        mail.logout()
                        return save_path

            mail.logout()
            return None
        except Exception as e:
            logging.error(f"[邮件] 下载附件失败: {e}")
            return None
