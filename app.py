import csv
import io
import json
import os
import sys
import re
import sqlite3
import time
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from flask import Flask, render_template, request, jsonify
import requests
from requests_ntlm import HttpNtlmAuth
from urllib.parse import quote

# ── PyInstaller 路徑處理 ────────────────────────────────────────
# 執行為 EXE 時：_MEIPASS 為解壓目錄（含 templates）；
# EXE 同目錄為使用者可編輯的 config.py / pdm_search.db
if getattr(sys, 'frozen', False):
    _BASE = sys._MEIPASS          # 唯讀資源（templates 等）
    _APP_DIR = os.path.dirname(sys.executable)  # EXE 所在目錄
    sys.path.insert(0, _APP_DIR)  # 讓 import config 找到 EXE 旁的 config.py
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))
    _APP_DIR = _BASE

import config

# 自主巡檢表路徑
INSPECTION_ROOT = r"\\192.168.1.99\共用區\品保加工共用平台\加工課自主巡檢表(空白)"

# PDM 圖面索引資料庫（可在 config.py 覆寫）
PDM_DB_PATH = getattr(config, 'PDM_DB_PATH',
    r"C:\Users\user\AppData\Local\PDMSearch\pdm_search.db")

# Fallback：若設定路徑不存在，嘗試同目錄下的 pdm_search.db（打包時隨包）
if not os.path.exists(PDM_DB_PATH):
    _bundled_db = os.path.join(_APP_DIR, 'pdm_search.db')
    if os.path.exists(_bundled_db):
        PDM_DB_PATH = _bundled_db

app = Flask(__name__, template_folder=os.path.join(_BASE, 'templates'))
app.config['TEMPLATES_AUTO_RELOAD'] = True

@app.after_request
def set_no_cache(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    # 動態注入版本號到所有 HTML 頁面（繞過模板快取）
    if response.content_type and 'text/html' in response.content_type:
        try:
            html = response.get_data(as_text=True)
            if '</head>' in html and APP_VERSION not in html:
                inject = f'<script>document.title="製令查詢 {APP_VERSION}";'
                inject += 'document.addEventListener("DOMContentLoaded",function(){'
                inject += f'var v=document.getElementById("app-ver");if(v)v.textContent="版本:{APP_VERSION}";'
                inject += '});</script>\n</head>'
                html = html.replace('</head>', inject)
                response.set_data(html)
        except Exception:
            pass
    return response

# ── TTL 快取（避免每次搜尋都重新呼叫 SSRS）──────────────────
CACHE_TTL = 120  # 快取有效期 120 秒

_cache = {}
_cache_lock = threading.Lock()


def cache_get(key):
    """取得快取資料，過期回傳 None"""
    with _cache_lock:
        entry = _cache.get(key)
        if entry and (time.time() - entry['ts']) < CACHE_TTL:
            return entry['data']
    return None


def cache_set(key, data):
    """設定快取資料"""
    with _cache_lock:
        _cache[key] = {'data': data, 'ts': time.time()}


def match_and_not(query, text):
    """AND/NOT 模糊搜尋語法
    空格分隔 = AND（全部都要符合）
    -前綴 = NOT（排除）
    例: 'FTB 夾管座'   → 包含 FTB 且包含 夾管座
        'FTB -夾管座'  → 包含 FTB 但不包含 夾管座
    """
    if not query:
        return True
    text_upper = text.upper()
    for token in query.strip().split():
        if token.startswith('-') and len(token) > 1:
            if token[1:].upper() in text_upper:
                return False
        else:
            if token.upper() not in text_upper:
                return False
    return True


def fetch_all_unfinished():
    """取得所有未完工製令（帶 TTL 快取）"""
    cached = cache_get('unfinished')
    if cached is not None:
        return cached
    params = {
        '發放情況': '已發放',
        'SFT完工碼': '尚未',
        '加工單位': '*',
    }
    csv_text = fetch_report_csv(config.REPORT_PATHS['unfinished'], params)
    data = parse_csv(csv_text)
    cache_set('unfinished', data)
    return data


_ssrs_session = None
_ssrs_session_lock = threading.Lock()


def get_ssrs_session():
    """建立/重用 SSRS NTLM 認證連線（連線池，避免每次建立新連線）"""
    global _ssrs_session
    with _ssrs_session_lock:
        if _ssrs_session is None:
            _ssrs_session = requests.Session()
            _ssrs_session.auth = HttpNtlmAuth(config.SSRS_USERNAME, config.SSRS_PASSWORD)
        return _ssrs_session


def fetch_report_csv(report_path, params=None):
    """從 SSRS 取得報表 CSV 資料"""
    session = get_ssrs_session()
    url = f'{config.SSRS_BASE_URL}?{quote(report_path)}&rs:Command=Render&rs:Format=CSV'
    if params:
        param_str = '&'.join(f'{quote(k)}={quote(str(v))}' for k, v in params.items())
        url += f'&{param_str}'
    resp = session.get(url, timeout=30)
    resp.raise_for_status()
    return resp.content.decode('utf-8-sig')


def parse_csv(csv_text):
    """解析 CSV 為 dict list"""
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)
    if not rows:
        return []
    header = rows[0]
    results = []
    for row in rows[1:]:
        if len(row) < len(header):
            continue
        results.append(dict(zip(header, row)))
    return results


UNIT_OPTIONS = {
    '000-1': '加工課',
    '000-3': '製三課',
    '000-2': '成品部',
}
KNOWN_UNITS = set(UNIT_OPTIONS.keys())


def search_orders(order_id='', product_name='', unit='000-1'):
    """從未完工製令報表搜尋，支援製令號碼、品名模糊搜尋和生產線篩選
    unit: '000-1'/'000-2'/'000-3'/'other'/'*'
    """
    all_records = fetch_all_unfinished()
    filtered = []
    for r in all_records:
        if order_id and order_id.upper() not in r.get('單別', '').upper():
            continue
        if product_name and not match_and_not(product_name, r.get('品名', '')):
            continue
        # 生產線篩選
        if unit and unit != '*':
            rec_unit = r.get('加工單位', '').strip()
            if unit == 'other':
                if rec_unit in KNOWN_UNITS:
                    continue
            else:
                if rec_unit != unit:
                    continue
        filtered.append(r)
    return filtered


def fetch_std_time(order_type, order_num):
    """從加工課-生產日報表取得標準工時和製程名稱（帶 TTL 快取）"""
    cache_key = f'std_{order_type}-{order_num}'
    cached = cache_get(cache_key)
    if cached is not None:
        return cached
    try:
        params = {'p1': order_type, 'p2': order_num}
        csv_text = fetch_report_csv(config.REPORT_PATHS['daily'], params)
        reader = csv.reader(io.StringIO(csv_text))
        rows = list(reader)
        result = {}
        for row in rows[1:]:
            if len(row) < 6:
                continue
            std_text = row[1]       # "標準工時：44 PCS/H"
            proc_code_text = row[3] # "製程代號：N19"
            proc_name_text = row[5] # "製程名稱：鑽+攻+倒"
            m_std = re.search(r'標準工時：(.+)', std_text)
            m_proc = re.search(r'製程代號：(.+)', proc_code_text)
            m_name = re.search(r'製程名稱：(.+)', proc_name_text)
            if m_proc:
                code = m_proc.group(1).strip()
                result[code] = {
                    'std_time': m_std.group(1).strip() if m_std else '',
                    'process_name': m_name.group(1).strip() if m_name else '',
                }
        cache_set(cache_key, result)
        return result
    except Exception:
        return {}


def fetch_efficiency_data():
    """從加工部總效率報表取得 SFT 預計開工日/完工日/標工/即時生產量（帶 TTL 快取）
    回傳 dict: { '製令號碼': { '製程代號': {...} } }
    """
    cached = cache_get('efficiency')
    if cached is not None:
        return cached
    try:
        csv_text = fetch_report_csv(config.REPORT_PATHS['efficiency'], {})
        records = parse_csv(csv_text)
        result = {}
        for r in records:
            oid = r.get('製令', '').strip()
            proc = r.get('製程', '').strip()
            if not oid:
                continue
            if oid not in result:
                result[oid] = {}
            result[oid][proc] = {
                'SFT預計開工日': r.get('SFT預計開工日', '').strip(),
                'SFT預計完工日': r.get('SFT預計完工日', '').strip(),
                '標工_秒': r.get('ERP變動人時3', '').strip(),
                '生產機台': r.get('生產機台', '').strip(),
                '即時生產量': r.get('現在產量', '').strip(),
            }
        cache_set('efficiency', result)
        return result
    except Exception:
        return {}


APP_VERSION = 'V26032703'

@app.route('/ver')
def ver_check():
    """版本確認（純文字，繞過模板和快取）"""
    return f'<h1>{APP_VERSION}</h1><p>template_folder={app.template_folder}</p>'

@app.route('/')
def index():
    return render_template('index.html', app_version=APP_VERSION)


@app.route('/drawing')
def drawing_page():
    return render_template('drawing.html')


@app.route('/api/query', methods=['GET'])
def query():
    """製令查詢 API：支援製令號碼和品名模糊搜尋"""
    order_id = request.args.get('order_id', '').strip()
    product_name_q = request.args.get('product_name', '').strip()
    unit = request.args.get('unit', '000-1').strip()

    if not order_id and not product_name_q:
        return jsonify({'success': False, 'error': '請輸入製令號碼或品名'}), 400

    try:
        # ── 第 1 步：取未完工製令 + 效率報表（並行） ──
        with ThreadPoolExecutor(max_workers=2) as pool:
            fut_orders = pool.submit(search_orders, order_id, product_name_q, unit)
            fut_eff = pool.submit(fetch_efficiency_data)
            records = fut_orders.result()
            eff_data = fut_eff.result()

        if not records:
            return jsonify({'success': False, 'error': '查無符合條件的資料'})

        # ── 第 2 步：日報表標工查詢（並行，限最多 20 筆） ──
        std_time_cache = {}
        unique_oids = list({r.get('單別', '') for r in records})
        oids_to_fetch = []
        for oid in unique_oids[:20]:
            oparts = oid.split('-', 1)
            if len(oparts) == 2:
                # 先檢查快取，只查沒快取的
                cache_key = f'std_{oparts[0]}-{oparts[1]}'
                cached = cache_get(cache_key)
                if cached is not None:
                    std_time_cache[oid] = cached
                else:
                    oids_to_fetch.append((oid, oparts[0], oparts[1]))

        if oids_to_fetch:
            with ThreadPoolExecutor(max_workers=min(8, len(oids_to_fetch))) as pool:
                futures = {
                    pool.submit(fetch_std_time, otype, onum): oid
                    for oid, otype, onum in oids_to_fetch
                }
                for fut in as_completed(futures):
                    oid = futures[fut]
                    try:
                        std_time_cache[oid] = fut.result()
                    except Exception:
                        std_time_cache[oid] = {}

        # ── 第 3 步：組合結果 ──
        data = []
        for r in records:
            rid = r.get('單別', '')
            proc_code = r.get('製程代號', '').strip()

            daily_info = std_time_cache.get(rid, {}).get(proc_code, {})
            eff_info = eff_data.get(rid, {}).get(proc_code, {})

            product_name = r.get('品名', '').rstrip('|').strip()

            std_time = daily_info.get('std_time', '')
            if not std_time and eff_info.get('標工_秒'):
                std_time = eff_info['標工_秒'] + ' 秒'
            # 標工：移除 "/H"，若只有 "- PCS" 等無數字則顯示空白
            if std_time:
                std_time = std_time.replace('/H', '').strip()
                if not re.search(r'\d', std_time):
                    std_time = ''

            realtime_qty = eff_info.get('即時生產量', '') or r.get('即時生產量', '')

            data.append({
                '製令': rid,
                '品號': r.get('品號', ''),
                '品名': product_name,
                '加工順序': r.get('加工順序', ''),
                '製程代號': proc_code,
                '製程名稱': daily_info.get('process_name', ''),
                '標工': std_time,
                '預計生產量': r.get('預計生產數', ''),
                '預計開工日': eff_info.get('SFT預計開工日', '') or r.get('製程預計開工日', ''),
                # SFT效率報表只有正在生產的製程；未開工製程用製程預計開工日作為估算完工日（has_sft_finish=False 表示估算值）
                '預計完工日': eff_info.get('SFT預計完工日', '') or r.get('製程預計開工日', ''),
                'has_sft_finish': bool(eff_info.get('SFT預計完工日', '')),
                '即時生產量': realtime_qty,
                '生產機台': eff_info.get('生產機台', ''),
            })

        return jsonify({'success': True, 'data': data})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/detail', methods=['GET'])
def detail():
    """製令詳細資訊 API：用料、製程、移轉單"""
    order_id = request.args.get('order_id', '').strip()
    if not order_id:
        return jsonify({'success': False, 'error': '請輸入製令號碼'}), 400

    try:
        params = {'製令': order_id}
        csv_text = fetch_report_csv(config.REPORT_PATHS['all'], params)

        # 報表回傳三段式 CSV（空行分隔）：材料、製程、移轉單
        csv_text = csv_text.replace('\r\n', '\n').replace('\r', '\n')
        sections = [s.strip() for s in csv_text.split('\n\n') if s.strip()]
        result = {'materials': [], 'processes': [], 'transfers': []}

        for section in sections:
            reader = csv.reader(io.StringIO(section))
            rows = list(reader)
            if len(rows) < 2:
                continue
            header = rows[0]
            data = []
            for row in rows[1:]:
                if len(row) >= len(header) and any(cell.strip() for cell in row):
                    data.append(dict(zip(header, row)))

            if '材料品號' in header:
                result['materials'] = data
            elif '完工否' in header:
                result['processes'] = data
            elif '移轉單' in header:
                result['transfers'] = data

        return jsonify({'success': True, 'data': result})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ── 列印暫存（pywebview 無法用 window.open，改用 URL 方式）──
_print_cache = {}
_print_cache_lock = threading.Lock()


def _render_print_html(records):
    """將列印資料渲染成 HTML"""
    from datetime import datetime
    today = datetime.now().strftime('%Y/%m/%d %H:%M')

    std_time_cache = {}
    for r in records:
        order_id = r.get('order_id', '')
        if order_id and order_id not in std_time_cache:
            parts = order_id.split('-')
            if len(parts) >= 2:
                std_time_cache[order_id] = fetch_std_time(parts[0], parts[1])

    items = []
    for r in records:
        order_id = r.get('order_id', '')
        process_seq = r.get('process_seq', '')
        process_code = r.get('process_code', '')
        process_name = r.get('process_name', '')
        unit = r.get('unit', '').strip()
        product_name = r.get('product_name', '')
        qty = r.get('qty', '')

        proc_info = std_time_cache.get(order_id, {}).get(process_code, {})
        std_time = proc_info.get('std_time', '') if isinstance(proc_info, dict) else ''
        actual_process_name = proc_info.get('process_name', '') if isinstance(proc_info, dict) else ''
        if not actual_process_name:
            actual_process_name = process_name

        barcode = f"{order_id};{process_seq};{process_code};{unit}"

        items.append({
            'order_id': f"{order_id}-{process_seq}",
            'product_name': product_name,
            'process_code': process_code,
            'process_name': actual_process_name,
            'qty': qty,
            'std_time': std_time,
            'barcode': barcode,
            'machine': r.get('machine', ''),
        })

    return render_template('print_report.html',
                           items=items,
                           items_json=json.dumps(items, ensure_ascii=False),
                           today=today)


@app.route('/print', methods=['POST'])
def print_report():
    """生產日報表列印頁面 — POST 暫存 HTML，回傳 key 供 GET 取用"""
    try:
        records = request.get_json() or []
    except Exception:
        return jsonify({'success': False, 'error': '資料格式錯誤'}), 400

    html = _render_print_html(records)
    key = str(uuid.uuid4())
    with _print_cache_lock:
        _print_cache[key] = html
    return jsonify({'success': True, 'key': key})


@app.route('/print/view/<key>', methods=['GET'])
def print_view(key):
    """取得暫存的列印 HTML"""
    with _print_cache_lock:
        html = _print_cache.get(key)
        # 列出目前快取的 keys 用於除錯
        cached_keys = list(_print_cache.keys())
    if html:
        return html
    return f'列印資料不存在。要求 key={key}，目前快取 keys={cached_keys}', 404


@app.route('/api/inspection', methods=['GET'])
def find_inspection():
    """根據品號查找自主巡檢表"""
    product_no = request.args.get('product_no', '').strip()
    if not product_no:
        return jsonify({'success': False, 'error': '請提供品號'}), 400

    try:
        root = Path(INSPECTION_ROOT)
        if not root.exists():
            return jsonify({'success': False, 'error': '無法存取巡檢表網路資料夾'}), 500

        found = []
        for f in root.rglob('*.xls*'):
            fname = f.stem
            if product_no.upper() in fname.upper():
                found.append({
                    'file_path': str(f),
                    'file_name': f.name,
                    'dir': str(f.parent),
                    'need_inspect': '需送檢' in fname,
                })

        if found:
            return jsonify({'success': True, 'files': found, 'count': len(found)})
        else:
            return jsonify({'success': False, 'error': '無對應自主檢查表', 'files': []})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/inspection/open', methods=['POST'])
def open_inspection():
    """開啟自主巡檢表，並將製令、預計產量寫入 Excel 指定欄位
    - 單張製令：G6 寫製令號碼（垂直置中），J6+K6 寫預計產量
    - 多張製令：G6 寫所有製令號碼（換行，垂直置中），J6/K6 不寫
    """
    import time as _time
    data = request.get_json() or {}
    file_path = data.get('file_path', '').strip()
    order_ids = data.get('order_ids', [])       # 製令號碼列表
    qty = data.get('qty', '').strip()            # 預計產量（僅單張時寫入）

    if not file_path:
        return jsonify({'success': False, 'error': '請提供檔案路徑'}), 400

    try:
        import pythoncom
        import win32com.client

        # 使用 STA threading model，避免 RPC_E_CALL_REJECTED 錯誤
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = True
            wb = excel.Workbooks.Open(file_path)

            # ── COM retry 機制：Excel 開啟後可能忙碌，需重試 ──
            max_retries = 3
            ws = None
            for attempt in range(max_retries):
                try:
                    _time.sleep(2)  # 等待 Excel 載入完成
                    ws = wb.ActiveSheet

                    # xlVAlignCenter = -4108
                    if order_ids:
                        cell_g6 = ws.Range("G6")
                        cell_g6.Value = '\n'.join(order_ids)
                        cell_g6.WrapText = True
                        cell_g6.VerticalAlignment = -4108   # 垂直置中
                        cell_g6.ShrinkToFit = False
                        # 自動調整文字大小以符合儲存格（避免跨到下一排）
                        if len(order_ids) > 1:
                            cell_g6.ShrinkToFit = True

                        # 同時填入 J6 和 K6 預計產量
                        if qty:
                            for cell_name in ["J6", "K6"]:
                                cell = ws.Range(cell_name)
                                cell.Value = qty
                                cell.VerticalAlignment = -4108

                    break  # 成功，跳出重試迴圈
                except pythoncom.com_error as ce:
                    if attempt < max_retries - 1:
                        _time.sleep(2)  # 等待後重試
                    else:
                        raise  # 最後一次仍失敗，拋出錯誤

            return jsonify({'success': True, 'message': f'已開啟並寫入 {os.path.basename(file_path)}'})
        finally:
            pythoncom.CoUninitialize()
    except ImportError:
        # pywin32 未安裝（如 Embeddable Python 環境）→ 直接用 os.startfile 開啟
        try:
            os.startfile(file_path)
            return jsonify({'success': True,
                            'message': f'已開啟 {os.path.basename(file_path)}（此電腦無法自動填入製令號碼，請手動輸入）'})
        except Exception as e2:
            return jsonify({'success': False, 'error': f'無法開啟檔案：{e2}'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


def get_pdm_db():
    """取得 PDMSearch SQLite 連線（自動嘗試多個路徑）"""
    candidates = [
        PDM_DB_PATH,                                                          # config.py 設定值
        os.path.join(_APP_DIR, 'pdm_search.db'),                              # 隨包附帶 (_app/)
        os.path.join(os.path.dirname(_APP_DIR), 'pdm_search.db'),             # 上一層 (製令查詢/)
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'PDMSearch', 'pdm_search.db'),  # 使用者 AppData
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            conn = sqlite3.connect(path)
            conn.row_factory = sqlite3.Row
            return conn
    return None


@app.route('/api/debug/paths')
def debug_paths():
    """診斷用：顯示 PDM 路徑解析結果"""
    candidates = [
        ('config.py 設定值', PDM_DB_PATH),
        ('_APP_DIR/pdm_search.db', os.path.join(_APP_DIR, 'pdm_search.db')),
        ('上一層/pdm_search.db', os.path.join(os.path.dirname(_APP_DIR), 'pdm_search.db')),
        ('LOCALAPPDATA', os.path.join(os.environ.get('LOCALAPPDATA', ''), 'PDMSearch', 'pdm_search.db')),
        ('sys.executable 同層', os.path.join(os.path.dirname(sys.executable), 'pdm_search.db')),
        ('__file__ 目錄', os.path.dirname(os.path.abspath(__file__))),
    ]
    result = {
        '_APP_DIR': _APP_DIR,
        '_BASE': _BASE,
        'PDM_DB_PATH': PDM_DB_PATH,
        'sys.executable': sys.executable,
        '__file__': os.path.abspath(__file__),
        'cwd': os.getcwd(),
        'candidates': []
    }
    for label, path in candidates:
        result['candidates'].append({
            'label': label,
            'path': path,
            'exists': os.path.isfile(path) if not label.endswith('目錄') else os.path.isdir(path),
        })
    # 列出 _APP_DIR 裡有什麼檔案
    try:
        result['_APP_DIR_files'] = os.listdir(_APP_DIR)
    except Exception as e:
        result['_APP_DIR_files'] = str(e)
    return jsonify(result)


@app.route('/api/drawing', methods=['GET'])
def find_drawing():
    """根據品號從 PDMSearch 索引資料庫查找圖面"""
    product_no = request.args.get('product_no', '').strip()
    if not product_no:
        return jsonify({'success': False, 'error': '請提供品號'}), 400

    try:
        conn = get_pdm_db()
        if not conn:
            return jsonify({'success': False, 'error': 'PDM 索引資料庫不存在'}), 500

        cur = conn.cursor()
        cur.execute(
            """SELECT file_path, pin_hao, tu_mian_pin_ming, ji_xing, xing_hao, modified_at
               FROM drawing_index WHERE UPPER(pin_hao) = UPPER(?)""",
            (product_no,)
        )
        rows = cur.fetchall()
        conn.close()

        if rows:
            files = _rows_to_drawing_list(rows)
            return jsonify({'success': True, 'files': files, 'count': len(files)})
        else:
            return jsonify({'success': False, 'error': f'PDM 索引中找不到品號 {product_no} 的圖面', 'files': []})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/drawing/search', methods=['GET'])
def search_drawing():
    """圖面模糊搜尋 API：支援品號或品名模糊搜尋（AND/NOT 語法）"""
    q = request.args.get('q', '').strip()
    if not q:
        return jsonify({'success': False, 'error': '請輸入品號或品名'}), 400

    try:
        conn = get_pdm_db()
        if not conn:
            return jsonify({'success': False, 'error': 'PDM 索引資料庫不存在'}), 500

        cur = conn.cursor()
        # 取出所有的關鍵字（排除 NOT 語法用的）做 SQL LIKE 初篩
        tokens = q.strip().split()
        positive = [t for t in tokens if not t.startswith('-')]
        conditions = []
        params = []
        for t in positive[:3]:  # 最多取前3個正向關鍵字做 SQL 篩選
            conditions.append(
                "(UPPER(pin_hao) LIKE ? OR UPPER(tu_mian_pin_ming) LIKE ? OR UPPER(file_path) LIKE ?)"
            )
            like = f'%{t.upper()}%'
            params.extend([like, like, like])

        where = ' AND '.join(conditions) if conditions else '1=1'
        cur.execute(
            f"""SELECT file_path, pin_hao, tu_mian_pin_ming, ji_xing, xing_hao, modified_at
                FROM drawing_index WHERE {where} LIMIT 200""",
            params
        )
        rows = cur.fetchall()
        conn.close()

        # 再用 match_and_not 做精確的 AND/NOT 過濾
        files = []
        for row in rows:
            combined = f"{row['pin_hao'] or ''} {row['tu_mian_pin_ming'] or ''} {row['file_path'] or ''}"
            if match_and_not(q, combined):
                files.append(_row_to_drawing_dict(row))

        if files:
            return jsonify({'success': True, 'files': files, 'count': len(files)})
        else:
            return jsonify({'success': False, 'error': '查無符合條件的圖面', 'files': []})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


def _row_to_drawing_dict(row):
    """將 SQLite Row 轉為圖面 dict"""
    fp = row['file_path']
    return {
        'file_path': fp,
        'file_name': os.path.basename(fp),
        'dir': os.path.dirname(fp),
        'pin_hao': row['pin_hao'] or '',
        'product_name': row['tu_mian_pin_ming'] or '',
        'ji_xing': row['ji_xing'] or '',
        'xing_hao': row['xing_hao'] or '',
        'modified_at': row['modified_at'] or '',
    }


def _rows_to_drawing_list(rows):
    """將多筆 SQLite Row 轉為圖面 dict list"""
    return [_row_to_drawing_dict(row) for row in rows]


def pdm_get_file(file_path):
    """用 PowerShell 呼叫 PDM COM API 取出檔案（Python COM 的 GetFileCopy 有型別問題，改用 PowerShell）"""
    import subprocess

    # PowerShell 腳本：透過 PDM COM 取出檔案到本機快取
    ps_script = f'''
$ErrorActionPreference = "Stop"
try {{
    $vault = New-Object -ComObject "ConisioLib.EdmVault"
    $vault.LoginAuto("MAXCLAW", 0)
    $path = "{file_path.replace(chr(92), chr(92)+chr(92))}"
    $dir = [System.IO.Path]::GetDirectoryName($path)
    $fname = [System.IO.Path]::GetFileName($path)
    $folder = $vault.GetFolderFromPath($dir)
    if ($null -eq $folder) {{
        Write-Error "PDM 找不到資料夾: $dir"
        exit 1
    }}
    $fileObj = $folder.GetFile($fname)
    if ($null -eq $fileObj) {{
        Write-Error "PDM 找不到檔案: $fname"
        exit 1
    }}
    $fileObj.GetFileCopy(0)
    Write-Output "OK"
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
'''
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
            capture_output=True, text=True, timeout=30
        )

        if result.returncode == 0 and 'OK' in result.stdout:
            return True, 'OK'
        else:
            err = result.stderr.strip()
            if not err:
                err = result.stdout.strip() or '未知錯誤'
            return False, err
    except subprocess.TimeoutExpired:
        return False, 'PDM 取出逾時'
    except Exception as e:
        return False, str(e)


@app.route('/api/drawing/open', methods=['POST'])
def open_drawing():
    """開啟 PDM 圖面檔案"""
    data = request.get_json() or {}
    file_path = data.get('file_path', '').strip()

    if not file_path:
        return jsonify({'success': False, 'error': '請提供圖面路徑'}), 400

    try:
        if os.path.exists(file_path):
            # 檔案已在本機（或可直接存取的網路路徑）
            # 嘗試 PDM 取最新版，但失敗不阻擋開啟
            pdm_get_file(file_path)
            os.startfile(file_path)
            return jsonify({'success': True, 'message': f'已開啟 {os.path.basename(file_path)}'})
        else:
            # 本機沒有此檔案，嘗試透過 PDM COM 取出
            ok, msg = pdm_get_file(file_path)
            if ok:
                os.startfile(file_path)
                return jsonify({'success': True, 'message': f'已開啟 {os.path.basename(file_path)}'})
            else:
                # PDM 取出失敗（可能未安裝 PDM）→ 仍嘗試直接開啟
                try:
                    os.startfile(file_path)
                    return jsonify({'success': True, 'message': f'已開啟 {os.path.basename(file_path)}'})
                except Exception:
                    pdm_hint = '（此電腦未安裝 SolidWorks PDM，無法自動取出圖面）' if 'ConisioLib' in msg or 'EdmVault' in msg or '無法建立' in msg else ''
                    return jsonify({'success': False,
                                    'error': f'無法開啟圖面：{msg}{pdm_hint}'}), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/cache/refresh', methods=['POST'])
def refresh_cache():
    """手動清除快取，下次查詢時會重新抓取 SSRS 資料"""
    with _cache_lock:
        _cache.clear()
    return jsonify({'success': True, 'message': '快取已清除'})


def _preload_cache():
    """背景預載 SSRS 資料到快取（加速首次搜尋）"""
    try:
        with ThreadPoolExecutor(max_workers=2) as pool:
            pool.submit(fetch_all_unfinished)
            pool.submit(fetch_efficiency_data)
    except Exception:
        pass


if __name__ == '__main__':
    # 獨立執行模式（除錯用）— GUI 版由 main.py 啟動
    threading.Thread(target=_preload_cache, daemon=True).start()
    app.run(host=config.FLASK_HOST, port=config.FLASK_PORT, debug=config.FLASK_DEBUG)
