# -*- coding: utf-8 -*-
"""排期录入系统 v5 - SKU映射 + 缓存 + 模糊搜索 + 邮件 + 进度条"""
import os, json, logging, threading, time, re
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file, Response
from excel_handler import ExcelHandler, EmailHandler
from pdf_parser import PDFParser
from excel_po_parser import ExcelPOParser

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


# =================== 反向代理前缀支持 ===================
class ReverseProxyMiddleware:
    """支持反向代理URL前缀（如 /schedule/）
    优先读取HTTP头 X-Script-Name / X-Forwarded-Prefix，
    否则使用环境变量 APP_PREFIX"""
    def __init__(self, wsgi_app, default_prefix=''):
        self.wsgi_app = wsgi_app
        self.default_prefix = default_prefix.rstrip('/')

    def __call__(self, environ, start_response):
        prefix = (environ.get('HTTP_X_SCRIPT_NAME', '') or
                  environ.get('HTTP_X_FORWARDED_PREFIX', '') or
                  self.default_prefix)
        if prefix:
            prefix = prefix.rstrip('/')
            environ['SCRIPT_NAME'] = prefix
            path_info = environ.get('PATH_INFO', '')
            if path_info.startswith(prefix):
                environ['PATH_INFO'] = path_info[len(prefix):] or '/'
        return self.wsgi_app(environ, start_response)


app.wsgi_app = ReverseProxyMiddleware(
    app.wsgi_app,
    default_prefix=os.environ.get('APP_PREFIX', '')
)

CFG_FILE = os.path.join(os.path.dirname(__file__), 'data', 'config.json')
os.makedirs(os.path.dirname(CFG_FILE), exist_ok=True)

LOG_FILE = os.path.join(os.path.dirname(__file__), 'data', 'ops.log')
ISSUES_FILE = os.path.join(os.path.dirname(__file__), 'data', 'issues.json')
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format='%(asctime)s %(message)s', encoding='utf-8')


def _load_issues():
    if not os.path.exists(ISSUES_FILE):
        return []
    try:
        with open(ISSUES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return []


def _save_issues(items):
    os.makedirs(os.path.dirname(ISSUES_FILE), exist_ok=True)
    with open(ISSUES_FILE, 'w', encoding='utf-8') as f:
        json.dump(items[-100:], f, ensure_ascii=False, indent=1)


def _add_issues(new_items):
    existing = _load_issues()
    existing.extend(new_items)
    _save_issues(existing)


def cfg():
    d = {'z_drive_path': r'Z:\各客排期\ZURU生产排期'}
    if os.path.exists(CFG_FILE):
        with open(CFG_FILE, 'r', encoding='utf-8') as f:
            d.update(json.load(f))
    return d


# =================== 页面 ===================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/history')
def history_page():
    return render_template('history.html')


@app.route('/settings')
def settings():
    return render_template('settings.html', config=cfg())


@app.route('/statistics')
def statistics_page():
    return render_template('statistics.html')


# =================== 仪表盘统计 ===================

@app.route('/api/dashboard')
def dashboard():
    """返回仪表盘统计数据"""
    handler = ExcelHandler(cfg())
    history = handler.get_history()

    today = datetime.now().strftime('%Y-%m-%d')
    today_ops = [h for h in history if h['time'].startswith(today)]
    week_ops = history[-50:] if len(history) > 50 else history

    # 统计
    today_new = sum(1 for h in today_ops if h['action'] == 'new')
    today_mod = sum(1 for h in today_ops if h['action'] == 'modify')
    today_can = sum(1 for h in today_ops if h['action'] == 'cancel')

    # Z盘文件数量
    files = handler._list_xlsx()
    file_count = len(files)

    # 检查哪些文件被锁定
    locked = []
    for fp in files[:20]:  # 只检查前20个，避免太慢
        try:
            with open(fp, 'r+b'):
                pass
        except:
            locked.append(os.path.basename(fp))

    return jsonify({
        'today': {'new': today_new, 'modify': today_mod, 'cancel': today_can,
                  'total': len(today_ops)},
        'total_history': len(history),
        'file_count': file_count,
        'locked_files': locked,
        'locked_count': len(locked),
    })


# =================== 批量上传解析 ===================

@app.route('/api/batch-upload', methods=['POST'])
def batch_upload():
    import time as _t
    _t0 = _t.time()
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': '未上传文件'}), 400

    pdf_parser = PDFParser()
    excel_po_parser = ExcelPOParser()
    handler = ExcelHandler(cfg())
    ExcelHandler.clear_cache()  # 刷新文件列表缓存
    orders = []
    all_issues = []
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M')

    # ===== 第1步：解析所有文件（支持PDF和Excel） =====
    parsed_list = []  # [(filename, data, warnings)]
    for f in files:
        path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
        f.save(path)
        ext = os.path.splitext(f.filename)[1].lower()
        try:
            if ext in ('.xlsx', '.xls'):
                data = excel_po_parser.parse(path)
            else:
                data = pdf_parser.parse(path)
        except Exception as e:
            issue = PDFParser.classify_error(f.filename, e)
            issue['filename'] = f.filename
            issue['time'] = now_str
            all_issues.append(issue)
            orders.append({'filename': f.filename, 'type': 'error',
                          'error': str(e), 'issue': issue})
            continue
        warnings = PDFParser.validate(data, f.filename)
        for w in warnings:
            w['filename'] = f.filename
            w['time'] = now_str
        all_issues.extend(warnings)
        parsed_list.append((f.filename, data, warnings))

    logging.info(f"[性能] 文件解析 {len(parsed_list)} 个文件耗时 {_t.time()-_t0:.1f}s")

    # ===== 第2步：批量并行搜索所有PO号（一次遍历所有文件） =====
    _t1 = _t.time()
    all_pos = [d.get('po_number', '') for _, d, _ in parsed_list if d.get('po_number')]
    po_results = handler.batch_search_pos(all_pos) if all_pos else {}
    # 排除总排期
    for k in po_results:
        po_results[k] = [r for r in po_results[k] if '总' not in r.get('fname', '')]
    logging.info(f"[性能] 批量PO搜索 {len(all_pos)} 个PO耗时 {_t.time()-_t1:.1f}s")

    # ===== 第3步：SKU匹配+auto_find（使用映射表快速定位） =====
    _t2 = _t.time()
    # 预缓存所有需要auto_find的SKU
    sku_cache = {}
    for filename, data, warnings in parsed_list:
        po = data.get('po_number', '')
        existing = po_results.get(po, [])
        if not existing and not data.get('is_cancel'):
            for ln in data.get('lines', []):
                sku = ln.get('item_code') or ln.get('sku', '')
                if sku and sku not in sku_cache:
                    sku_cache[sku] = handler.auto_find(sku)

    logging.info(f"[性能] SKU匹配 {len(sku_cache)} 个SKU耗时 {_t.time()-_t2:.1f}s")

    # ===== 第4步：组装订单结果 =====
    for filename, data, warnings in parsed_list:
        po = data.get('po_number', '')

        if data.get('is_cancel'):
            existing = po_results.get(po, [])
            actions = [{'type': 'cancel', 'record': rec,
                        'sku': rec['data'].get('F', '') or rec['data'].get('G', ''),
                        'detail': f"取消 {rec['data'].get('F','')} 行{rec['row']}"}
                       for rec in existing]
            orders.append({
                'filename': filename, 'type': 'cancel',
                'po_number': po, 'customer': data.get('customer', ''),
                'ship_date': data.get('ship_date', ''),
                'lines': data.get('lines', []),
                'actions': actions,
                'header': _build_header(data),
                'line_count': len(data.get('lines', [])),
                'existing_count': len(existing),
                'warnings': warnings,
            })
            continue

        existing = po_results.get(po, [])

        if existing:
            actions = handler.smart_diff(data, existing)
            order_type = 'modify'
        else:
            actions = []
            for ln in data.get('lines', []):
                sku = ln.get('item_code') or ln.get('sku', '')
                sched = sku_cache.get(sku)
                if not sched:
                    issue = {
                        'category': 'sku_not_found',
                        'title': f'无相同货号(item)，无法写入 · {sku} · PO {po}',
                        'icon': 'bi-exclamation-triangle',
                        'color': 'danger',
                        'filename': filename,
                        'sku': sku,
                        'time': now_str,
                        'tip': (f'PO {po} 的货号 "{sku}" 在Z盘所有排期文件中未找到相同item的参考行。\n'
                                '因为中文名、内箱、外箱、总箱数等产品固有属性必须从同货号参考行复制，\n'
                                '无参考行则无法保证数据正确性，此行跳过不写入。\n'
                                '请手动到对应排期文件中录入，或检查货号是否正确。')
                    }
                    all_issues.append(issue)
                actions.append({
                    'type': 'new', 'line': ln, 'schedule': sched,
                    'sku': ln.get('sku', ''),
                    'detail': f"新增 {ln.get('sku','')} {ln.get('qty',0)}pcs"
                })
            order_type = 'new'

        orders.append({
            'filename': filename, 'type': order_type,
            'po_number': po, 'customer': data.get('customer', ''),
            'ship_date': data.get('ship_date', ''),
            'from_person': data.get('from_person', ''),
            'lines': data.get('lines', []),
            'actions': actions,
            'header': _build_header(data),
            'line_count': len(data.get('lines', [])),
            'existing_count': len(existing),
            'tracking_code': data.get('tracking_code', ''),
            'packaging_info': data.get('packaging_info', ''),
            'remark': data.get('remark', ''),
            'warnings': warnings,
        })

    if all_issues:
        _add_issues(all_issues)
        logging.info(f"[异常提醒] 发现 {len(all_issues)} 个问题")

    total_time = round(_t.time() - _t0, 1)
    logging.info(f"[性能] 批量上传总耗时 {total_time}s (PDF×{len(files)}, PO×{len(all_pos)}, SKU×{len(sku_cache)})")
    return jsonify({'orders': orders, 'issues': all_issues, 'parse_seconds': total_time})


@app.route('/api/batch-execute', methods=['POST'])
def batch_execute():
    import time as _time
    _t0 = _time.time()
    orders = request.json.get('orders', [])
    handler = ExcelHandler(cfg())
    result = handler.batch_process(orders)
    result['elapsed_seconds'] = round(_time.time() - _t0, 1)

    # 只记录成功的操作到历史（失败的不记录）
    ok_files = {r.get('file', '') for r in (result.get('results', []))}
    failed_files = {f.get('file', '') for f in (result.get('failed', []))}
    for order in orders:
        po = order.get('header', {}).get('po_number', '')
        for act in order.get('actions', []):
            # 跳过未实际执行的新单（auto_find未匹配到排期文件，schedule为空）
            if act['type'] == 'new' and not act.get('schedule'):
                continue
            fname = (act.get('record', {}).get('fname', '') or
                     (act.get('schedule', {}) or {}).get('fname', ''))
            # 检查该文件是否在成功列表中（不在失败列表中）
            if fname and fname in failed_files:
                continue
            ExcelHandler.add_history(po, act['type'], act.get('detail', ''), fname)
        logging.info(f"批量处理 PO={po} 动作数={len(order.get('actions', []))}")

    # 将失败项加入待重试队列
    if result.get('failed'):
        pending = ExcelHandler.get_pending_retries()
        for f in result['failed']:
            f['time'] = datetime.now().strftime('%Y-%m-%d %H:%M')
            f['retries'] = 0
            # 提取PO号用于显示
            for order in orders:
                f['po'] = order.get('header', {}).get('po_number', '')
                break
            pending.append(f)
        ExcelHandler.save_pending_retries(pending)

    return jsonify(result)


@app.route('/api/retry-failed', methods=['POST'])
def retry_failed():
    items = request.json.get('items', [])
    handler = ExcelHandler(cfg())
    result = handler.retry_save(items)
    for item in result.get('ok', []):
        logging.info(f"重试成功: {item['file']}")
    # 更新待重试队列
    if result.get('failed'):
        pending = ExcelHandler.get_pending_retries()
        for f in result['failed']:
            if not any(p['local'] == f['local'] for p in pending):
                f['time'] = datetime.now().strftime('%Y-%m-%d %H:%M')
                f['retries'] = f.get('retries', 0) + 1
                pending.append(f)
        ExcelHandler.save_pending_retries(pending)
    return jsonify(result)


# =================== 备份 ===================

@app.route('/api/schedule-files')
def list_schedule_files():
    """列出所有排期文件（带修改时间），供备份选择"""
    handler = ExcelHandler(cfg())
    files = handler._list_xlsx()
    result = []
    for fp in files:
        fn = os.path.basename(fp)
        if '总' in fn:
            continue
        try:
            mtime = os.path.getmtime(fp)
            mt = datetime.fromtimestamp(mtime)
            result.append({
                'path': fp, 'name': fn,
                'mtime': mt.strftime('%Y-%m-%d %H:%M'),
                'mtime_ts': mtime,
                'recent': (datetime.now() - mt).total_seconds() < 86400  # 24h内修改
            })
        except:
            result.append({'path': fp, 'name': fn, 'mtime': '', 'mtime_ts': 0, 'recent': False})
    result.sort(key=lambda x: -x['mtime_ts'])
    return jsonify(result)


@app.route('/api/backup', methods=['POST'])
def create_backup():
    files = request.json.get('files', [])
    handler = ExcelHandler(cfg())
    result = handler.create_backup(files if files else None)
    logging.info(f"备份: {result['msg']}")
    return jsonify(result)


# =================== 历史记录 ===================

@app.route('/api/history')
def get_history():
    return jsonify(ExcelHandler.get_history())


# =================== 兼容旧接口 ===================

@app.route('/api/process-pdf', methods=['POST'])
def process_pdf():
    if 'file' not in request.files:
        return jsonify({'error': '未上传文件'}), 400
    f = request.files['file']
    path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
    f.save(path)
    try:
        data = PDFParser().parse(path)
    except Exception as e:
        return jsonify({'error': f'PDF解析失败: {e}'}), 500
    handler = ExcelHandler(cfg())
    line_schedules = []
    for ln in data.get('lines', []):
        sku = ln.get('item_code') or ln.get('sku', '')
        found = handler.auto_find(sku)
        line_schedules.append({
            'line': ln,
            'schedule': {'file': found['file'], 'fname': found['fname'],
                         'sheet': found['sheet'], 'ref': found['ref'],
                         'mcol': found['mcol']} if found else None
        })
    po = data.get('po_number', '')
    existing = handler.search_po(po) if po else []
    data['line_schedules'] = line_schedules
    data['existing'] = existing
    data['order_type'] = 'modify' if existing else 'new'
    return jsonify(data)


@app.route('/api/execute', methods=['POST'])
def execute():
    d = request.json
    action = d.get('action')
    handler = ExcelHandler(cfg())
    try:
        if action == 'new':
            r = handler.enter_new(d['schedule'], d['header'], d['lines'])
            ExcelHandler.add_history(d['header'].get('po_number', ''), 'new',
                                    f"新单 → {d['schedule']['sheet']}", d['schedule']['fname'])
            return jsonify(r)
        elif action == 'modify':
            r = handler.modify(d['record'], d['changes'])
            ExcelHandler.add_history('', 'modify',
                                    f"修改 → {d['record']['sheet']} 行{d['record']['row']}")
            return jsonify(r)
        elif action == 'cancel':
            r = handler.cancel(d['record'])
            ExcelHandler.add_history('', 'cancel',
                                    f"取消 → {d['record']['sheet']} 行{d['record']['row']}")
            return jsonify(r)
        return jsonify({'error': '未知操作'}), 400
    except Exception as e:
        logging.error(f"失败: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/save-z', methods=['POST'])
def save_z():
    d = request.json
    try:
        return jsonify(ExcelHandler(cfg()).save_z(d['local'], d['z']))
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/search-po', methods=['POST'])
def search_po():
    po = request.json.get('po', '')
    if not po:
        return jsonify({'error': '请输入PO号'}), 400
    return jsonify(ExcelHandler(cfg()).search_po(po))


@app.route('/api/config', methods=['POST'])
def save_cfg():
    with open(CFG_FILE, 'w', encoding='utf-8') as f:
        json.dump(request.json, f, ensure_ascii=False, indent=2)
    return jsonify({'ok': True})


@app.route('/api/logs')
def get_logs():
    if not os.path.exists(LOG_FILE):
        return jsonify([])
    with open(LOG_FILE, 'r', encoding='utf-8') as f:
        lines = f.readlines()[-100:]
    return jsonify([l.strip() for l in lines if l.strip()])


def _build_header(data):
    return {
        'po_number': data.get('po_number', ''),
        'customer': data.get('customer', ''),
        'po_date': data.get('po_date', ''),
        'ship_date': data.get('ship_date', ''),
        'destination_cn': data.get('destination_cn', '') or data.get('destination', ''),
        'from_person': data.get('from_person', ''),
        'customer_po_header': data.get('customer_po_header', ''),
        'tracking_code': data.get('tracking_code', ''),
        'packaging_info': data.get('packaging_info', ''),
        'remark': data.get('remark', ''),
    }


# =================== 异常提醒 ===================

@app.route('/api/issues')
def get_issues():
    return jsonify(_load_issues())


@app.route('/api/clear-issues', methods=['POST'])
def clear_issues():
    _save_issues([])
    return jsonify({'ok': True})


# =================== 排期分类导航 ===================

@app.route('/api/schedule-dirs')
def schedule_dirs():
    """扫描各客排期目录，按客户分类返回"""
    base = r'Z:\各客排期'
    if not os.path.isdir(base):
        return jsonify({'error': '各客排期目录不存在', 'groups': []})
    groups = []
    try:
        for name in sorted(os.listdir(base)):
            fp = os.path.join(base, name)
            if not os.path.isdir(fp) or name.startswith(('~', '.')):
                continue
            # 统计xlsx文件数
            xlsx = [f for f in os.listdir(fp)
                    if f.endswith('.xlsx') and not f.startswith('~$')]
            sub_dirs = [d for d in os.listdir(fp)
                        if os.path.isdir(os.path.join(fp, d)) and not d.startswith(('~', '.'))]
            groups.append({
                'name': name,
                'path': fp,
                'file_count': len(xlsx),
                'sub_count': len(sub_dirs),
            })
    except Exception as e:
        return jsonify({'error': str(e), 'groups': []})
    return jsonify({'groups': groups})


@app.route('/api/open-folder', methods=['POST'])
def open_folder():
    """在Windows资源管理器中打开目录"""
    path = request.json.get('path', '')
    if not path or not os.path.exists(path):
        return jsonify({'error': '路径不存在'}), 400
    import subprocess
    try:
        subprocess.Popen(['explorer', os.path.normpath(path)])
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# =================== 路径浏览 ===================

@app.route('/api/browse-dirs', methods=['POST'])
def browse_dirs():
    """浏览目录，返回子目录和xlsx文件列表；path为空或不存在时列出所有驱动器"""
    import string as _str
    path = (request.json.get('path', '') or '').strip()

    # 路径为空或不存在 → 列出所有可用驱动器（Windows）
    if not path or not os.path.isdir(path):
        if os.name == 'nt':
            drives = []
            for letter in _str.ascii_uppercase:
                drive = f'{letter}:\\'
                if os.path.exists(drive):
                    drives.append({'name': f'{letter}: 盘', 'path': drive})
            return jsonify({'current': '', 'parent': '', 'dirs': drives, 'files': [], 'is_drives': True})
        path = '/'

    dirs, files = [], []
    try:
        for item in sorted(os.listdir(path)):
            fp = os.path.join(path, item)
            if os.path.isdir(fp) and not item.startswith(('~', '.')):
                dirs.append({'name': item, 'path': fp})
            elif item.endswith('.xlsx') and not item.startswith('~$'):
                files.append({'name': item, 'path': fp})
    except PermissionError:
        return jsonify({'error': '无权限访问', 'dirs': [], 'files': []}), 403
    # 到达驱动器根目录时 parent 为空，前端可据此返回驱动器列表
    drive_root = os.path.splitdrive(path)[0] + '\\'
    parent = '' if path.rstrip('\\') == drive_root.rstrip('\\') else os.path.dirname(path.rstrip('\\'))
    return jsonify({'current': path, 'parent': parent, 'dirs': dirs, 'files': files})


@app.route('/api/set-schedule-path', methods=['POST'])
def set_schedule_path():
    """设置排期文件路径"""
    new_path = request.json.get('path', '')
    if not new_path or not os.path.isdir(new_path):
        return jsonify({'error': '无效路径'}), 400
    c = cfg()
    c['z_drive_path'] = new_path
    with open(CFG_FILE, 'w', encoding='utf-8') as f:
        json.dump(c, f, ensure_ascii=False, indent=2)
    return jsonify({'ok': True, 'path': new_path})


@app.route('/api/current-path')
def current_path():
    return jsonify({'path': cfg().get('z_drive_path', '')})


# =================== 排期同步 ===================

@app.route('/api/sync-schedules', methods=['POST'])
def sync_schedules():
    """从Z盘正式排期目录同步到本地目录"""
    import shutil
    source = request.json.get('source', r'Z:\各客排期\ZURU生产排期')
    dest = request.json.get('dest', cfg().get('z_drive_path', ''))
    if not dest:
        return jsonify({'error': '未设置目标路径'}), 400
    if not os.path.isdir(source):
        return jsonify({'error': f'源目录不存在: {source}'}), 400
    os.makedirs(dest, exist_ok=True)

    src_files = {f for f in os.listdir(source)
                 if f.endswith('.xlsx') and not f.startswith('~$')}
    dst_files = {f for f in os.listdir(dest)
                 if f.endswith('.xlsx') and not f.startswith('~$')}

    copied, updated, deleted, skipped = 0, 0, 0, 0
    # 复制/更新
    for fname in src_files:
        sp = os.path.join(source, fname)
        dp = os.path.join(dest, fname)
        if not os.path.exists(dp):
            shutil.copy2(sp, dp)
            copied += 1
        else:
            src_mt = os.path.getmtime(sp)
            dst_mt = os.path.getmtime(dp)
            if src_mt > dst_mt:
                shutil.copy2(sp, dp)
                updated += 1
            else:
                skipped += 1
    # 删除目标中源已不存在的旧文件
    for fname in dst_files - src_files:
        os.remove(os.path.join(dest, fname))
        deleted += 1

    msg = f'同步完成：新增{copied} 更新{updated} 删除{deleted} 跳过{skipped}'
    logging.info(f"[排期同步] {msg} ({source} → {dest})")
    return jsonify({
        'ok': True, 'msg': msg,
        'copied': copied, 'updated': updated,
        'deleted': deleted, 'skipped': skipped,
        'total': len(src_files)
    })


# =================== 总排期操作 ===================

@app.route('/api/master-info')
def master_info():
    """获取总排期信息"""
    handler = ExcelHandler(cfg())
    master = handler.find_master_schedule()
    if not master:
        return jsonify({'exists': False})
    fn = os.path.basename(master)
    locked = False
    try:
        with open(master, 'r+b'):
            pass
    except:
        locked = True
    return jsonify({'exists': True, 'file': fn, 'path': master, 'locked': locked})


@app.route('/api/scan-yellow', methods=['POST'])
def scan_yellow():
    """扫描分排期中的黄色填充行"""
    handler = ExcelHandler(cfg())
    rows = handler.scan_yellow_rows()
    return jsonify({'count': len(rows), 'rows': rows[:500]})


@app.route('/api/copy-to-master', methods=['POST'])
def copy_to_master():
    """将黄色填充行复制到总排期"""
    handler = ExcelHandler(cfg())
    result = handler.copy_to_master()
    if result.get('ok'):
        logging.info(f"[总排期] 复制 {result['copied']} 行到 {result['master_file']}")
    return jsonify(result)


@app.route('/api/clear-master-yellow', methods=['POST'])
def clear_master_yellow():
    """清除总排期中的黄色填充"""
    handler = ExcelHandler(cfg())
    result = handler.clear_master_yellow()
    if result.get('ok'):
        logging.info(f"[总排期] 清除 {result['cleared']} 行黄色填充")
    return jsonify(result)


# =================== 模糊搜索 ===================

@app.route('/api/fuzzy-search', methods=['POST'])
def fuzzy_search():
    """模糊搜索：PO号、SKU、客户名等"""
    keyword = request.json.get('keyword', '')
    if not keyword or len(keyword) < 2:
        return jsonify({'error': '请输入至少2个字符'}), 400
    handler = ExcelHandler(cfg())
    results = handler.fuzzy_search(keyword)
    return jsonify({'count': len(results), 'results': results[:50]})


# =================== SKU映射 ===================

SKU_MAP_FILE = os.path.join(os.path.dirname(__file__), 'data', 'sku_mapping.json')

@app.route('/api/sku-mapping')
def sku_mapping():
    """查看SKU→排期文件映射"""
    handler = ExcelHandler(cfg())
    return jsonify(handler.get_sku_mapping_info())


@app.route('/api/refresh-sku-mapping', methods=['POST'])
def refresh_sku_mapping():
    """强制刷新SKU映射缓存"""
    import excel_handler
    excel_handler._sku_map_cache = {}
    excel_handler._sku_map_mtime = 0
    handler = ExcelHandler(cfg())
    info = handler.get_sku_mapping_info()
    return jsonify({'ok': True, 'msg': f'已刷新，共 {info["total"]} 个映射项', **info})


@app.route('/api/sku-mapping/edit', methods=['POST'])
def edit_sku_mapping():
    """增删改SKU映射条目"""
    action = request.json.get('action')  # add, delete, update
    sku = request.json.get('sku', '').strip().upper()
    keywords = request.json.get('keywords', [])

    if not sku:
        return jsonify({'error': '请输入SKU'}), 400

    # 读取当前JSON
    data = {'_说明': 'SKU/货号→排期文件关键词映射表', '_更新时间': '', 'mapping': {}}
    if os.path.exists(SKU_MAP_FILE):
        with open(SKU_MAP_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)

    mapping = data.get('mapping', {})

    if action == 'add' or action == 'update':
        if not keywords:
            return jsonify({'error': '请输入关键词'}), 400
        mapping[sku] = keywords
    elif action == 'delete':
        if sku in mapping:
            del mapping[sku]
        else:
            return jsonify({'error': f'SKU {sku} 不存在'}), 404
    else:
        return jsonify({'error': '未知操作'}), 400

    data['mapping'] = mapping
    data['_更新时间'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    with open(SKU_MAP_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # 清除缓存
    import excel_handler
    excel_handler._sku_map_cache = {}
    excel_handler._sku_map_mtime = 0

    logging.info(f"[SKU映射] {action} SKU={sku} keywords={keywords}")
    return jsonify({'ok': True, 'total': len(mapping),
                    'msg': f'已{"添加" if action=="add" else "更新" if action=="update" else "删除"} {sku}'})


# =================== 历史导出 ===================

@app.route('/api/export-history')
def export_history():
    """导出操作历史为Excel"""
    fpath = ExcelHandler.export_history_excel()
    if not fpath:
        return jsonify({'error': '无历史记录'}), 400
    return send_file(fpath, as_attachment=True,
                     download_name=os.path.basename(fpath),
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# =================== 批量处理进度 ===================

@app.route('/api/batch-progress')
def batch_progress():
    """获取批量处理进度"""
    return jsonify(ExcelHandler.get_batch_progress())


# =================== 手动选择排期 ===================

@app.route('/api/list-schedules')
def list_schedules():
    """列出所有排期文件供手动选择"""
    handler = ExcelHandler(cfg())
    return jsonify(handler.list_schedule_files())


@app.route('/api/manual-assign', methods=['POST'])
def manual_assign():
    """手动指定SKU对应的排期文件"""
    filepath = request.json.get('file', '')
    sheet = request.json.get('sheet', '')
    if not filepath or not sheet:
        return jsonify({'error': '请选择文件和工作表'}), 400
    handler = ExcelHandler(cfg())
    result = handler.manual_find_ref(filepath, sheet)
    if 'error' in result:
        return jsonify(result), 500
    return jsonify(result)


# =================== 邮件集成 ===================

@app.route('/api/email/check', methods=['POST'])
def email_check():
    """检查邮箱新邮件"""
    c = cfg()
    handler = EmailHandler(c)
    return jsonify(handler.check_new_emails())


@app.route('/api/email/download', methods=['POST'])
def email_download():
    """下载邮件附件并解析"""
    msg_id = request.json.get('msg_id', '')
    filename = request.json.get('filename', '')
    if not msg_id or not filename:
        return jsonify({'error': '参数缺失'}), 400

    c = cfg()
    handler = EmailHandler(c)
    save_dir = app.config['UPLOAD_FOLDER']
    path = handler.download_attachment(msg_id, filename, save_dir)
    if not path:
        return jsonify({'error': '下载失败'}), 500

    # 自动解析PDF
    try:
        data = PDFParser().parse(path)
        return jsonify({'ok': True, 'path': path, 'filename': filename, 'data': data})
    except Exception as e:
        return jsonify({'ok': True, 'path': path, 'filename': filename,
                       'error': f'PDF解析失败: {e}'})


@app.route('/api/email/settings', methods=['GET', 'POST'])
def email_settings():
    """获取/保存邮箱设置"""
    if request.method == 'GET':
        c = cfg()
        return jsonify({
            'server': c.get('email_server', ''),
            'port': c.get('email_port', 993),
            'user': c.get('email_user', ''),
            'password': '***' if c.get('email_password') else '',
            'ssl': c.get('email_ssl', True),
        })
    # POST: save
    data = request.json
    c = cfg()
    c['email_server'] = data.get('server', '')
    c['email_port'] = data.get('port', 993)
    c['email_user'] = data.get('user', '')
    if data.get('password') and data['password'] != '***':
        c['email_password'] = data['password']
    c['email_ssl'] = data.get('ssl', True)
    with open(CFG_FILE, 'w', encoding='utf-8') as f:
        json.dump(c, f, ensure_ascii=False, indent=2)
    return jsonify({'ok': True})


# =================== 统计仪表盘（高级）===================

@app.route('/api/statistics')
def statistics():
    """按月/客户/产品线统计录入量"""
    history = ExcelHandler.get_history()
    if not history:
        return jsonify({'monthly': [], 'by_customer': [], 'by_action': {}, 'total': 0})

    from collections import Counter, defaultdict
    monthly = defaultdict(lambda: {'new': 0, 'modify': 0, 'cancel': 0})
    by_customer = Counter()
    by_action = Counter()
    by_file = Counter()  # 产品线≈排期文件

    for h in history:
        t = h.get('time', '')
        ym = t[:7] if len(t) >= 7 else 'unknown'
        act = h.get('action', 'new')
        monthly[ym][act] = monthly[ym].get(act, 0) + 1
        by_action[act] += 1

        detail = h.get('detail', '')
        files = h.get('files', '')
        if files:
            # 从文件名提取产品线
            fname = os.path.basename(files)
            fname_short = re.sub(r'^20\d{2}年ZURU', '', fname).replace('.xlsx', '').strip()
            if fname_short:
                by_file[fname_short] += 1

        # 从detail提取客户/SKU信息
        m = re.search(r'→\s*(.+?)(?:\s|$)', detail)
        if m:
            by_customer[m.group(1).strip()] += 1

    # 格式化月度数据
    monthly_list = []
    for ym in sorted(monthly.keys(), reverse=True)[:12]:
        d = monthly[ym]
        monthly_list.append({
            'month': ym,
            'new': d.get('new', 0),
            'modify': d.get('modify', 0),
            'cancel': d.get('cancel', 0),
            'total': d.get('new', 0) + d.get('modify', 0) + d.get('cancel', 0)
        })

    return jsonify({
        'total': len(history),
        'monthly': monthly_list,
        'by_action': dict(by_action),
        'by_product': [{'name': k, 'count': v} for k, v in by_file.most_common(20)],
        'by_target': [{'name': k, 'count': v} for k, v in by_customer.most_common(20)],
    })


# =================== 文件变更监控 ===================

_file_snapshots = {}  # {filepath: mtime}

@app.route('/api/file-changes')
def file_changes():
    """检测排期文件变更（对比上次快照）"""
    global _file_snapshots
    handler = ExcelHandler(cfg())
    files = handler._list_xlsx()
    changes = []
    current = {}

    for fp in files:
        fn = os.path.basename(fp)
        try:
            mtime = os.path.getmtime(fp)
            size = os.path.getsize(fp)
        except:
            continue
        current[fp] = mtime

        if fp in _file_snapshots:
            old_mtime = _file_snapshots[fp]
            if mtime > old_mtime:
                # 文件被修改了
                from_time = datetime.fromtimestamp(old_mtime).strftime('%H:%M')
                to_time = datetime.fromtimestamp(mtime).strftime('%H:%M')
                # 检查是否是被其他人修改（看锁文件）
                lock_file = os.path.join(os.path.dirname(fp), '~$' + fn[2:] if len(fn) > 2 else fn)
                user = ''
                if os.path.exists(lock_file):
                    user = handler._read_lock_user(lock_file)
                changes.append({
                    'file': fn, 'path': fp,
                    'old_time': from_time, 'new_time': to_time,
                    'user': user,
                    'size_kb': round(size / 1024, 1),
                })
        else:
            # 新文件
            pass

    # 更新快照
    _file_snapshots = current

    return jsonify({
        'changes': changes,
        'total_files': len(files),
        'scanned_at': datetime.now().strftime('%H:%M:%S'),
    })


# =================== 列映射配置 ===================

@app.route('/api/column-mapping', methods=['GET', 'POST'])
def column_mapping():
    """获取/保存总排期列映射配置"""
    if request.method == 'GET':
        c = cfg()
        return jsonify(c.get('column_mapping', {}))
    c = cfg()
    c['column_mapping'] = request.json
    with open(CFG_FILE, 'w', encoding='utf-8') as f:
        json.dump(c, f, ensure_ascii=False, indent=2)
    return jsonify({'ok': True})


# =================== 撤回操作 ===================

@app.route('/api/undo-info')
def undo_info():
    """获取可撤回信息"""
    return jsonify(ExcelHandler.get_undo_info())


@app.route('/api/undo-last', methods=['POST'])
def undo_last():
    """撤回上次批量操作（兼容旧接口）"""
    handler = ExcelHandler(cfg())
    result = handler.undo_last_batch()
    if result.get('ok'):
        logging.info(f"[撤回] {result['msg']}")
    return jsonify(result)


@app.route('/api/undo-selected', methods=['POST'])
def undo_selected():
    """撤回用户选中的批次操作"""
    batch_ids = request.json.get('batch_ids', [])
    if not batch_ids:
        return jsonify({'error': '未选择要撤回的操作'})
    handler = ExcelHandler(cfg())
    result = handler.undo_selected(batch_ids)
    if result.get('ok'):
        logging.info(f"[选择性撤回] {result['msg']} ids={batch_ids}")
    return jsonify(result)


# =================== 文件占用监控 ===================

@app.route('/api/file-status')
def file_status():
    """返回所有排期文件的占用状态"""
    handler = ExcelHandler(cfg())
    return jsonify(handler.check_all_file_status())


# =================== 待重试队列 + 自动重试 ===================

@app.route('/api/pending-retries')
def pending_retries():
    return jsonify(ExcelHandler.get_pending_retries())


@app.route('/api/auto-retry', methods=['POST'])
def auto_retry():
    handler = ExcelHandler(cfg())
    result = handler.auto_retry_pending()
    if result.get('ok'):
        logging.info(f"自动重试: {result['msg']}")
    return jsonify(result)


@app.route('/api/clear-pending', methods=['POST'])
def clear_pending():
    """清空待重试队列"""
    ExcelHandler.save_pending_retries([])
    return jsonify({'ok': True})


@app.route('/api/delete-entries', methods=['POST'])
def delete_entries():
    """删除指定的排期行（支持撤销）"""
    entries = request.json.get('entries', [])
    if not entries:
        return jsonify({'error': '没有要删除的条目'})
    handler = ExcelHandler(cfg())
    result = handler.delete_entries_com(entries)
    if result.get('ok'):
        logging.info(f"[删除入单行] {result['msg']}")
    return jsonify(result)


@app.route('/api/reentry', methods=['POST'])
def reentry():
    """一键重新入单（重新执行batch_process），返回详细反馈"""
    import time as _time
    _t0 = _time.time()
    orders = request.json.get('orders', [])
    if not orders:
        return jsonify({'error': '没有订单数据'})
    handler = ExcelHandler(cfg())
    result = handler.reentry_batch(orders)
    result['elapsed_seconds'] = round(_time.time() - _t0, 1)
    # 记录历史
    for order in orders:
        po = order.get('header', {}).get('po_number', '')
        for act in order.get('actions', []):
            if act['type'] == 'new' and not act.get('schedule'):
                continue
            fname = (act.get('record', {}).get('fname', '') or
                     (act.get('schedule', {}) or {}).get('fname', ''))
            ExcelHandler.add_history(po, 'reentry', act.get('detail', ''), fname)
    logging.info(f"[一键重入] 处理完成 耗时{result.get('elapsed_seconds')}s")
    return jsonify(result)


@app.route('/api/schedule-retry', methods=['POST'])
def schedule_retry():
    """安排定时重新入单"""
    scheduled_time = request.json.get('time', '')  # 格式: "HH:MM"
    orders = request.json.get('orders', [])
    label = request.json.get('label', '')
    if not scheduled_time or not orders:
        return jsonify({'error': '缺少时间或订单数据'})
    items = ExcelHandler.get_scheduled_retries()
    item = {
        'id': datetime.now().strftime('%Y%m%d%H%M%S'),
        'time': scheduled_time,
        'orders': orders,
        'label': label or f"PO重入 ({len(orders)}单)",
        'created': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'status': 'pending'
    }
    items.append(item)
    ExcelHandler.save_scheduled_retries(items)
    logging.info(f"[定时入单] 已安排 {scheduled_time} 执行: {item['label']}")
    return jsonify({'ok': True, 'id': item['id'], 'msg': f'已安排在 {scheduled_time} 自动重新入单'})


@app.route('/api/scheduled-retries')
def get_scheduled_retries():
    return jsonify(ExcelHandler.get_scheduled_retries())


@app.route('/api/cancel-scheduled', methods=['POST'])
def cancel_scheduled():
    """取消定时任务"""
    task_id = request.json.get('id', '')
    items = ExcelHandler.get_scheduled_retries()
    items = [i for i in items if i['id'] != task_id]
    ExcelHandler.save_scheduled_retries(items)
    return jsonify({'ok': True})


# 后台自动重试线程
_auto_retry_running = True
_auto_retry_interval = 180  # 3分钟


def _auto_retry_worker():
    """后台线程：每3分钟自动重试保存失败的文件 + 检查定时任务"""
    while _auto_retry_running:
        try:
            handler = ExcelHandler(cfg())
            # 1. 自动重试保存失败的文件
            result = handler.auto_retry_pending()
            ok_count = len(result.get('ok', []))
            if ok_count > 0:
                logging.info(f"[自动重试] 成功保存{ok_count}个文件")

            # 2. 检查定时重入任务
            try:
                now = datetime.now()
                now_hm = now.strftime('%H:%M')
                items = ExcelHandler.get_scheduled_retries()
                remaining = []
                for item in items:
                    if item.get('status') != 'pending':
                        continue
                    sched_time = item.get('time', '')
                    if sched_time and sched_time <= now_hm:
                        # 执行定时任务
                        logging.info(f"[定时入单] 开始执行: {item.get('label', '')}")
                        try:
                            orders = item.get('orders', [])
                            result = handler.batch_process(orders)
                            ok_n = len(result.get('results', []))
                            fail_n = len(result.get('failed', []))
                            item['status'] = 'done'
                            item['result'] = f'成功{ok_n}个' + (f'，失败{fail_n}个' if fail_n else '')
                            item['done_time'] = now.strftime('%H:%M:%S')
                            logging.info(f"[定时入单] 完成: {item['result']}")
                        except Exception as se:
                            item['status'] = 'error'
                            item['result'] = str(se)[:200]
                            logging.error(f"[定时入单] 失败: {se}")
                    remaining.append(item)
                if remaining != items:
                    ExcelHandler.save_scheduled_retries(remaining)
            except Exception as se:
                logging.error(f"[定时检查] 异常: {se}")

        except Exception as e:
            logging.error(f"[自动重试] 异常: {e}")
        # 分段休眠（每秒检查退出信号，每60秒轮询一次）
        for _ in range(60):
            if not _auto_retry_running:
                break
            time.sleep(1)


def start_auto_retry():
    t = threading.Thread(target=_auto_retry_worker, daemon=True, name='auto-retry')
    t.start()
    logging.info("[自动重试] 后台线程已启动，间隔60秒（含定时任务检查）")


if __name__ == '__main__':
    print('=' * 50)
    print('  ZURU 排期录入系统 v4')
    print('  http://localhost:5000')
    print('  自动重试：每3分钟检查待保存文件')
    print('=' * 50)
    start_auto_retry()
    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=True)
