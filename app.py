# -*- coding: utf-8 -*-
"""总排期入单系统 - 只操作Z盘总排期文件"""
import os, sys, json, logging, re
from datetime import datetime
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file
from excel_po_parser import ExcelPOParser
from master_schedule import write_orders, generate_excel, DEFAULT_MASTER_PATH, lookup_schedule_info
from generate_yellow_summary import generate_summary, generate_summary_excel

APP_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=os.path.join(APP_DIR, 'templates'),
            static_folder=os.path.join(APP_DIR, 'static'))
app.config['UPLOAD_FOLDER'] = os.path.join(APP_DIR, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 512 * 1024 * 1024
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

LOG_FILE = os.path.join(APP_DIR, 'data', 'ops.log')
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format='%(asctime)s %(message)s', encoding='utf-8')


@app.route('/')
def index():
    return render_template('master.html')


_ignore_cache = {'mtime': 0, 'items': set()}


def _load_ignore_items():
    """加载黑名单 data/ignore_items.json，带mtime缓存避免重复读"""
    p = os.path.join(APP_DIR, 'data', 'ignore_items.json')
    if not os.path.exists(p):
        return set()
    try:
        mtime = os.path.getmtime(p)
        if _ignore_cache['mtime'] == mtime:
            return _ignore_cache['items']
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        items = set(str(x).strip().upper() for x in data.get('ignore_items', []) if x)
        _ignore_cache['mtime'] = mtime
        _ignore_cache['items'] = items
        return items
    except Exception as e:
        logging.warning(f'[黑名单] 加载失败: {e}')
        return set()


def _filter_ignored(orders):
    """按货号数字前缀匹配黑名单，命中的行从orders中移除
    返回 (过滤后orders, 忽略报告, 忽略行的完整orders列表)
    报告格式: [{'item': '15790', 'count': 3, 'source': 'PO-xxx.xlsx'}]
    lines全空的order会被整体丢弃
    """
    ignore_set = _load_ignore_items()
    if not ignore_set:
        return orders, [], []
    report = []
    new_orders = []
    ignored_orders = []  # 保留被忽略行的完整order结构
    for od in orders:
        kept_lines = []
        ignored_lines = []
        ignored_count = {}
        for ln in od.get('lines', []):
            sku = ln.get('sku_spec') or ln.get('sku', '')
            num_m = re.match(r'^(\d+)', str(sku).strip())
            prefix = num_m.group(1).upper() if num_m else ''
            if prefix and prefix in ignore_set:
                ignored_count[prefix] = ignored_count.get(prefix, 0) + 1
                ignored_lines.append(ln)
                continue
            kept_lines.append(ln)
        for item, cnt in ignored_count.items():
            report.append({
                'item': item,
                'count': cnt,
                'source': od.get('filename', ''),
            })
            logging.info(f'[黑名单忽略] {od.get("filename","")}: 货号{item} x {cnt}行')
        if ignored_lines:
            # 复制order header，替换lines为被忽略的行
            ig_od = dict(od)
            ig_od['lines'] = ignored_lines
            ignored_orders.append(ig_od)
        if kept_lines:
            od['lines'] = kept_lines
            new_orders.append(od)
        else:
            logging.info(f'[黑名单忽略] {od.get("filename","")}: 所有行都在黑名单，整单丢弃')
    return new_orders, report, ignored_orders


def _extract_revision(filename):
    """从文件名提取版本号，用于同PO去重时判断哪个更新
    匹配: R1/R.1/R2/Rev.1/Rev1/V1 等模式
    Windows副本后缀 (1)/(2) 视为版本0
    返回: int 版本号（越大越新，无版本号返回0）
    """
    name = os.path.splitext(filename)[0]
    # 去掉Windows副本后缀 (1) (2) 等
    name = re.sub(r'\(\d+\)\s*$', '', name).strip()
    # 匹配 Rev.2 / Rev2 / R.2 / R2 / V1 等（取最后一个匹配，因为文件名可能有多个数字）
    patterns = [
        r'[Rr]ev\.?\s*(\d+)',   # Rev.2, Rev2, rev.1
        r'[Rr]\.?\s*(\d+)',     # R.1, R1, R2
        r'[Vv]\.?\s*(\d+)',     # V1, V.1
    ]
    best = 0
    for pat in patterns:
        for m in re.finditer(pat, name):
            v = int(m.group(1))
            if v > best:
                best = v
    return best


def _dedup_orders(orders):
    """按PO号去重：同PO号保留文件名版本号最大的，返回 (去重后orders, 去重报告列表)
    去重报告格式: ["PO 4500197032: 保留 兴信R2.xlsx，去掉 兴信R1.xlsx"]
    """
    by_po = defaultdict(list)
    deduped = []
    for od in orders:
        header = od.get('header') or od
        po = str(header.get('po_number', '') or od.get('po_number', '')).strip()
        if po.endswith('.0'):
            po = po[:-2]
        fname = od.get('filename', '')
        rev = _extract_revision(fname)
        if not po:
            deduped.append(od)
            continue
        by_po[po].append((rev, fname, od))

    report = []
    for po, entries in by_po.items():
        if len(entries) <= 1:
            deduped.append(entries[0][2])
            continue
        # 按版本号降序排，同版本按文件名降序（确定性）
        entries.sort(key=lambda x: (x[0], x[1]), reverse=True)
        keep_rev, keep_fname, keep_od = entries[0]
        deduped.append(keep_od)
        removed_names = [e[1] for e in entries[1:]]
        report.append(f'PO {po}: 保留 {keep_fname}，去掉 {", ".join(removed_names)}')
        logging.info(f'[去重] PO {po}: 保留 {keep_fname}(rev={keep_rev}), '
                     f'去掉 {", ".join(removed_names)}')
    return deduped, report


@app.route('/api/master-schedule-upload', methods=['POST'])
def master_schedule_upload():
    """上传PO文件写入Z盘总排期"""
    saved = []
    for f in request.files.getlist('files'):
        if not f.filename:
            continue
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ('.xlsx', '.xls'):
            continue
        path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
        f.save(path)
        saved.append((f.filename, path))

    if not saved:
        return jsonify({'error': '没有有效的Excel文件'}), 400

    orders = []
    errors = []
    for fname, path in saved:
        try:
            data = ExcelPOParser().parse(path)
            data['filename'] = fname
            orders.append(data)
        except Exception as e:
            errors.append(f'{fname}: {str(e)[:80]}')

    if not orders:
        return jsonify({'error': f'所有文件解析失败: {"; ".join(errors)}'}), 400

    # 同PO号去重：保留最新版本
    orders, dedup_report = _dedup_orders(orders)

    # 黑名单过滤：特定货号不入排期
    orders, ignored_report, ignored_orders = _filter_ignored(orders)
    if not orders:
        return jsonify({
            'ok': True,
            'msg': '所有PO行都在黑名单中，已全部忽略',
            'dedup_report': dedup_report,
            'ignored_report': ignored_report,
            'modified': 0, 'new_count': 0,
            'mod_details': [], 'new_details': [],
            'errors': errors, 'warnings': [],
        })

    # 数据异常检测：PDF转Excel可能丢失数据，检测qty/outer为0的行
    data_warnings = []
    for od in orders:
        fname = od.get('filename', '')
        for ln in od.get('lines', []):
            sku = ln.get('sku_spec') or ln.get('sku', '')
            qty = ln.get('qty', 0) or 0
            outer = ln.get('outer_qty', 0) or 0
            if qty <= 0:
                data_warnings.append(f'{fname}: {sku} 数量=0，可能是PDF转Excel时数据丢失，请核对原始PO')
            elif outer <= 0:
                data_warnings.append(f'{fname}: {sku} 外箱装箱数=0，可能是PDF转Excel时数据丢失，请核对原始PO')

    master_path = _get_master_path()
    export_dir = os.path.join(APP_DIR, 'exports')
    try:
        result = write_orders(master_path, orders, export_dir=export_dir,
                              ignored_orders=ignored_orders)
        if not result['ok']:
            return jsonify({'error': result['msg']}), 500
        resp = {
            'ok': True,
            'msg': result['msg'],
            'modified': result.get('modified', 0),
            'new_count': result.get('new_count', 0),
            'mod_details': result.get('mod_details', []),
            'new_details': result.get('new_details', []),
            'errors': errors,
            'warnings': result.get('warnings', []),
            'dedup_report': dedup_report,
            'ignored_report': ignored_report,
        }
        if data_warnings:
            resp['data_warnings'] = data_warnings
        if result.get('export_file'):
            resp['export_file'] = result['export_file']
        # 分排期归属查找（纯附加信息，不影响主流程）
        try:
            all_items = []
            for o in orders:
                for ln in o.get('lines', []):
                    s = ln.get('sku_spec', '') or ln.get('sku', '')
                    if s: all_items.append(s)
            sch_info = lookup_schedule_info(all_items)
            resp['schedule_info'] = sch_info
        except Exception:
            pass
        return jsonify(resp)
    except PermissionError:
        return jsonify({'error': '总排期文件被占用，请关闭WPS中的总排期后重试'}), 500
    except Exception as e:
        return jsonify({'error': f'处理失败: {e}'}), 500


@app.route('/api/master-export-download/<filename>')
def master_export_download(filename):
    """下载生成的入单Excel"""
    export_dir = os.path.join(APP_DIR, 'exports')
    filepath = os.path.join(export_dir, filename)
    if not os.path.abspath(filepath).startswith(os.path.abspath(export_dir)):
        return jsonify({'error': '非法路径'}), 403
    if not os.path.exists(filepath):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(filepath, as_attachment=True, download_name=filename)


@app.route('/api/master-schedule-download')
def master_schedule_download():
    """下载总排期文件"""
    mp = _get_master_path()
    if not os.path.exists(mp):
        return jsonify({'error': '总排期文件不存在'}), 404
    return send_file(mp, as_attachment=True,
                     download_name=os.path.basename(mp))


_custom_path = {'path': ''}  # 用户自定义路径（运行时修改）


def _get_master_path():
    return _custom_path['path'] or DEFAULT_MASTER_PATH


@app.route('/api/yellow-summary', methods=['POST'])
def yellow_summary():
    """扫描总排期有填充行汇总"""
    mp = _get_master_path()
    try:
        result = generate_summary(mp)
        return jsonify(result)
    except PermissionError as e:
        return jsonify({'error': str(e)}), 500
    except Exception as e:
        return jsonify({'error': f'生成汇总失败: {e}'}), 500


@app.route('/api/yellow-summary-download', methods=['POST'])
def yellow_summary_download():
    """生成有填充行汇总Excel并下载"""
    mp = _get_master_path()
    export_dir = os.path.join(APP_DIR, 'exports')
    try:
        fname = generate_summary_excel(mp, export_dir)
        if not fname:
            return jsonify({'error': '无整行填充行，无法生成'}), 400
        filepath = os.path.join(export_dir, fname)
        return send_file(filepath, as_attachment=True, download_name=fname)
    except Exception as e:
        return jsonify({'error': f'生成汇总Excel失败: {e}'}), 500


@app.route('/api/master-schedule-info')
def master_schedule_info():
    """总排期文件状态"""
    mp = _get_master_path()
    exists = os.path.exists(mp)
    locked = False
    if exists:
        try:
            fh = open(mp, 'r+b')
            fh.close()
        except (PermissionError, OSError):
            locked = True
    return jsonify({
        'exists': exists,
        'locked': locked,
        'path': mp,
    })


@app.route('/api/master-schedule-set-path', methods=['POST'])
def master_schedule_set_path():
    """切换总排期文件路径（只接受.xlsx文件）"""
    new_path = request.json.get('path', '').strip()
    if not new_path:
        _custom_path['path'] = ''
        return jsonify({'ok': True, 'path': DEFAULT_MASTER_PATH, 'msg': '已恢复默认Z盘路径'})
    if not os.path.exists(new_path):
        return jsonify({'error': f'路径不存在: {new_path}'}), 400
    if not new_path.lower().endswith('.xlsx'):
        return jsonify({'error': '请选择.xlsx文件'}), 400
    _custom_path['path'] = new_path
    return jsonify({'ok': True, 'path': new_path, 'msg': f'已切换到: {new_path}'})


@app.route('/api/ignore-items', methods=['GET'])
def get_ignore_items():
    """获取黑名单列表"""
    p = os.path.join(APP_DIR, 'data', 'ignore_items.json')
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return jsonify({'items': data.get('ignore_items', [])})
    except Exception as e:
        return jsonify({'items': [], 'error': str(e)})


@app.route('/api/ignore-items', methods=['POST'])
def update_ignore_items():
    """更新黑名单列表"""
    items = request.json.get('items', [])
    # 清洗：去空白、去重、纯数字排序
    cleaned = sorted(set(str(x).strip() for x in items if str(x).strip()),
                     key=lambda x: int(x) if x.isdigit() else float('inf'))
    p = os.path.join(APP_DIR, 'data', 'ignore_items.json')
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        data = {}
    data['ignore_items'] = cleaned
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # 清除缓存让下次加载生效
    _ignore_cache['mtime'] = None
    return jsonify({'ok': True, 'items': cleaned, 'count': len(cleaned)})


@app.route('/api/dual-map', methods=['GET'])
def get_dual_map():
    """获取双排期货号映射"""
    p = os.path.join(APP_DIR, 'data', 'dual_schedule_map.json')
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # 过滤掉_开头的说明字段
        items = {k: v for k, v in data.items() if not k.startswith('_')}
        return jsonify({'items': items})
    except Exception as e:
        return jsonify({'items': {}, 'error': str(e)})


@app.route('/api/dual-map', methods=['POST'])
def update_dual_map():
    """更新双排期货号映射"""
    items = request.json.get('items', {})
    p = os.path.join(APP_DIR, 'data', 'dual_schedule_map.json')
    # 保留原有的_说明字段
    try:
        with open(p, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception:
        data = {}
    meta = {k: v for k, v in data.items() if k.startswith('_')}
    meta.update(items)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    # 刷新内存缓存
    from master_schedule import _load_dual_map, _DUAL_MAP
    import master_schedule
    master_schedule._DUAL_MAP = _load_dual_map()
    return jsonify({'ok': True, 'count': len(items)})


@app.route('/api/master-schedule-upload-file', methods=['POST'])
def master_schedule_upload_file():
    """上传总排期文件（本地Excel上传到服务器使用）"""
    if 'file' not in request.files:
        return jsonify({'error': '未选择文件'}), 400
    f = request.files['file']
    if not f.filename or not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': '请选择 .xlsx 或 .xls 文件'}), 400
    save_path = os.path.join(APP_DIR, 'data', 'uploaded_master.xlsx')
    try:
        f.save(save_path)
    except Exception as e:
        return jsonify({'error': f'保存失败: {e}'}), 500
    _custom_path['path'] = save_path
    return jsonify({'ok': True, 'path': save_path, 'msg': f'已上传并切换到: {f.filename}'})


@app.route('/api/browse-directory', methods=['POST'])
def browse_directory():
    """浏览目录，返回子目录和xlsx文件列表"""
    dir_path = os.path.normpath(request.json.get('path', '').strip()) if request.json.get('path', '').strip() else ''
    if not dir_path:
        # 默认起始目录：总排期所在文件夹
        dir_path = os.path.dirname(DEFAULT_MASTER_PATH)
    if not os.path.exists(dir_path) or not os.path.isdir(dir_path):
        return jsonify({'error': f'目录不存在: {dir_path}'}), 400
    try:
        entries = os.listdir(dir_path)
    except PermissionError:
        return jsonify({'error': f'无权访问: {dir_path}'}), 403
    dirs = sorted([e for e in entries if os.path.isdir(os.path.join(dir_path, e))
                   and not e.startswith('.')])
    files = sorted([e for e in entries if e.lower().endswith('.xlsx')
                    and not e.startswith('~$')])
    parent = os.path.dirname(dir_path)
    return jsonify({
        'current': dir_path,
        'parent': parent if parent != dir_path else '',
        'dirs': dirs,
        'files': files,
    })


if __name__ == '__main__':
    CFG_FILE = os.path.join(APP_DIR, 'data', 'config.json')
    port = 5003
    if os.path.exists(CFG_FILE):
        try:
            with open(CFG_FILE, encoding='utf-8') as f:
                port = json.load(f).get('port', 5003)
        except:
            pass
    print('=' * 50)
    print('  ZURU 总排期入单系统')
    print(f'  http://localhost:{port}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=True, use_reloader=False, threaded=True)
