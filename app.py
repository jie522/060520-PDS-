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
import webbrowser
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from flask import Flask, render_template, request, jsonify, redirect
import requests
from requests_ntlm import HttpNtlmAuth
from urllib.parse import quote
try:
    import openpyxl
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

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
    # python311._pth 存在時，嵌入式 Python 不自動加入腳本目錄
    # 需手動插入，讓 import config / templates 等都能正常運作
    if _APP_DIR not in sys.path:
        sys.path.insert(0, _APP_DIR)
    _PARENT = os.path.dirname(_APP_DIR)
    if os.path.exists(os.path.join(_PARENT, 'config.py')) and _PARENT not in sys.path:
        sys.path.insert(0, _PARENT)

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

# ── 技術資料清單資料庫（zume-n.com 圖號↔URL 對照表）──────────────────────
ZUME_DB_PATH = os.path.join(_APP_DIR, 'zume_drawings.db')

def _init_zume_db():
    con = sqlite3.connect(ZUME_DB_PATH)
    con.execute('''
        CREATE TABLE IF NOT EXISTS drawings (
            part_no   TEXT PRIMARY KEY,
            part_name TEXT,
            url       TEXT NOT NULL
        )
    ''')
    con.execute('''
        CREATE TABLE IF NOT EXISTS import_log (
            filename  TEXT PRIMARY KEY,
            imported_at TEXT,
            count     INTEGER
        )
    ''')
    con.commit(); con.close()

def _parse_zume_csv(filepath):
    """解析 zume-n_data_list_*.csv，回傳 [(part_no, part_name, url), ...]"""
    rows = []
    try:
        with open(filepath, encoding='utf-8-sig', newline='') as f:
            reader = csv.reader(f)
            headers = [h.strip() for h in next(reader)]
            idx_no  = next((i for i,h in enumerate(headers) if '圖號' in h), None)
            idx_url = next((i for i,h in enumerate(headers) if h.upper() == 'URL'), None)
            idx_nm  = next((i for i,h in enumerate(headers) if '品名' in h), 1)
            if idx_no is None or idx_url is None:
                return []
            for row in reader:
                if len(row) > max(idx_no, idx_url):
                    no  = row[idx_no].strip()
                    url = row[idx_url].strip()
                    nm  = row[idx_nm].strip() if idx_nm < len(row) else ''
                    if no and url.startswith('http'):
                        rows.append((no, nm, url))
    except Exception:
        pass
    return rows

def _auto_import_zume_csv():
    """啟動時自動掃描 Downloads 資料夾，匯入最新的 zume-n_data_list_*.csv"""
    import glob as _glob, getpass as _gp
    try:
        user = _gp.getuser()
    except Exception:
        user = os.environ.get('USERNAME', 'user')
    downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    pattern = os.path.join(downloads, 'zume-n_data_list_*.csv')
    files = sorted(_glob.glob(pattern), key=os.path.getmtime, reverse=True)
    if not files:
        return
    latest = files[0]
    fname  = os.path.basename(latest)
    con = sqlite3.connect(ZUME_DB_PATH)
    already = con.execute('SELECT count FROM import_log WHERE filename=?', (fname,)).fetchone()
    con.close()
    if already:
        return  # 已匯入過，略過
    rows = _parse_zume_csv(latest)
    if not rows:
        return
    con = sqlite3.connect(ZUME_DB_PATH)
    con.executemany('INSERT OR REPLACE INTO drawings(part_no,part_name,url) VALUES(?,?,?)', rows)
    con.execute('INSERT OR REPLACE INTO import_log(filename,imported_at,count) VALUES(?,datetime("now"),?)',
                (fname, len(rows)))
    con.commit(); con.close()

_init_zume_db()
_auto_import_zume_csv()   # 啟動時自動掃描匯入

app = Flask(__name__,
            template_folder=os.path.join(_BASE, 'templates'),
            static_folder=os.path.join(_BASE, 'static'))
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
                inject = f'<script>document.title="製令查詢 {APP_VERSION}";</script>\n</head>'
                html = html.replace('</head>', inject)
                response.set_data(html)
        except Exception:
            pass
    return response

# ── 全域錯誤處理（確保所有錯誤都回傳 JSON，而非 HTML）────────
@app.errorhandler(400)
def bad_request(e):
    return jsonify({'success': False, 'error': f'請求錯誤：{str(e)}'}), 400

@app.errorhandler(404)
def not_found(e):
    return jsonify({'success': False, 'error': f'路由不存在：{str(e)}'}), 404

@app.errorhandler(500)
def internal_error(e):
    return jsonify({'success': False, 'error': f'伺服器錯誤：{str(e)}'}), 500

@app.errorhandler(Exception)
def unhandled_exception(e):
    return jsonify({'success': False, 'error': f'未預期的錯誤：{str(e)}'}), 500

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


def fetch_efficiency_all_groups():
    """透過 SSRS SOAP Execution Service 取得加工部總效率(全部)報表 CSV"""
    import base64 as _b64
    import re as _re
    session = get_ssrs_session()
    soap_url = f'http://192.168.1.212/ReportServer/ReportExecution2005.asmx'

    def _soap(body, action):
        return session.post(soap_url, data=body.encode('utf-8'),
            headers={'Content-Type': 'text/xml; charset=utf-8',
                     'SOAPAction': f'"http://schemas.microsoft.com/sqlserver/2005/06/30/reporting/reportingservices/{action}"'},
            timeout=30)

    NS = 'xmlns:rs="http://schemas.microsoft.com/sqlserver/2005/06/30/reporting/reportingservices"'
    ENV_OPEN = f'<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" {NS}>'

    # Step 1: LoadReport
    r1 = _soap(f'{ENV_OPEN}<soap:Body><rs:LoadReport><rs:Report>/加工課/加工部總效率</rs:Report></rs:LoadReport></soap:Body></soap:Envelope>', 'LoadReport')
    r1.raise_for_status()
    eid = _re.search(r'<ExecutionID>([^<]+)</ExecutionID>', r1.text)
    if not eid:
        raise RuntimeError('LoadReport: 無法取得 ExecutionID')
    exec_id = eid.group(1)

    # Step 2: SetExecutionParameters ORG=全部
    r2 = _soap(f'{ENV_OPEN}<soap:Header><rs:ExecutionHeader><rs:ExecutionID>{exec_id}</rs:ExecutionID></rs:ExecutionHeader></soap:Header><soap:Body><rs:SetExecutionParameters><rs:Parameters><rs:ParameterValue><rs:Name>ORG</rs:Name><rs:Value>全部</rs:Value></rs:ParameterValue></rs:Parameters><rs:ParameterLanguage>zh-TW</rs:ParameterLanguage></rs:SetExecutionParameters></soap:Body></soap:Envelope>', 'SetExecutionParameters')
    r2.raise_for_status()

    # Step 3: Render as CSV
    r3 = _soap(f'{ENV_OPEN}<soap:Header><rs:ExecutionHeader><rs:ExecutionID>{exec_id}</rs:ExecutionID></rs:ExecutionHeader></soap:Header><soap:Body><rs:Render><rs:Format>CSV</rs:Format><rs:DeviceInfo></rs:DeviceInfo></rs:Render></soap:Body></soap:Envelope>', 'Render')
    r3.raise_for_status()
    b64 = _re.search(r'<Result>([^<]+)</Result>', r3.text)
    if not b64:
        raise RuntimeError('Render: 無法取得 CSV 結果')
    csv_bytes = _b64.b64decode(b64.group(1))
    return csv_bytes.decode('utf-8-sig', errors='replace')


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


# ── ServCloud 設備稼動 ─────────────────────────────────────────────
_servcloud_session = None
_servcloud_session_lock = threading.Lock()
SERVCLOUD_BASE = getattr(config, 'SERVCLOUD_BASE', 'http://192.168.1.69:58080/ServCloud')
SERVCLOUD_USER = getattr(config, 'SERVCLOUD_USER', 'adminstd')
SERVCLOUD_PASS = getattr(config, 'SERVCLOUD_PASS', 'adminstd')

_STATUS_MAP = {
    '0':  ('離線', '#bdbdbd'),
    '11': ('運轉', '#43a047'),
    '12': ('待機', '#fb8c00'),
    '13': ('警報', '#e53935'),
    'B':  ('離線', '#bdbdbd'),
}

def _sc_login(s):
    s.post(f'{SERVCLOUD_BASE}/api/user/login',
           data={'id': SERVCLOUD_USER, 'password': SERVCLOUD_PASS}, timeout=10)

def get_servcloud_session():
    global _servcloud_session
    with _servcloud_session_lock:
        if _servcloud_session is None:
            s = requests.Session()
            _sc_login(s)
            _servcloud_session = s
        return _servcloud_session

def _sc_post(path, payload, retry=True):
    """POST to ServCloud; auto re-login on session expiry (type==2)."""
    session = get_servcloud_session()
    resp = session.post(f'{SERVCLOUD_BASE}{path}', json=payload, timeout=30)
    data = resp.json()
    if data.get('type') == 2 and retry:
        global _servcloud_session
        with _servcloud_session_lock:
            _sc_login(session)
            _servcloud_session = session
        return _sc_post(path, payload, retry=False)
    return data

def _sc_parse(raw):
    """ServCloud data 欄位可能是 JSON 字串或已解析的 list/dict，統一回傳 list。"""
    if isinstance(raw, list):
        return raw
    if isinstance(raw, str):
        return json.loads(raw)
    return []

def servcloud_get_machines():
    data = _sc_post('/api/getdata/db',
                    {'table': 'm_device', 'columns': ['device_id', 'device_name']})
    machines = _sc_parse(data.get('data', []))
    return [m for m in machines if m.get('device_name') and m['device_name'] != m['device_id']]

def servcloud_get_history(machine_ids, date_str):
    """date_str: YYYYMMDD。最多 10 台/次，自動分批。"""
    all_recs = []
    for i in range(0, len(machine_ids), 10):
        batch = machine_ids[i:i+10]
        data = _sc_post('/api/hippo/simple', {
            'space': 'machine_status_history',
            'index': {'machine_id': batch},
            'indexRange': {'key': 'date', 'start': date_str, 'end': date_str},
            'columns': ['machine_id', 'status', 'start_time', 'end_time', 'duration']
        })
        if data.get('type') == 0 and data.get('data'):
            all_recs.extend(_sc_parse(data['data']))
    return all_recs


APP_VERSION = 'V20260520001'

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


@app.route('/production')
def production_page():
    return render_template('production.html', app_version=APP_VERSION)


@app.route('/api/production', methods=['GET'])
def production_data():
    """生產中製令 API：用 SOAP 取得加工部總效率(全部)"""
    try:
        csv_text = fetch_efficiency_all_groups()
        records = parse_csv(csv_text)
        # 回傳原始欄位名稱供前端自動對應
        raw_columns = list(records[0].keys()) if records else []
        return jsonify({'success': True, 'data': records, 'count': len(records), 'columns': raw_columns})
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500


@app.route('/api/production/debug')
def production_debug():
    """除錯：顯示 SSRS CSV 原始欄位名稱"""
    try:
        csv_text = fetch_report_csv(config.REPORT_PATHS['efficiency'], {})
        reader = csv.reader(io.StringIO(csv_text))
        rows = list(reader)
        if not rows:
            return '<h2>CSV 空白</h2>'
        header = rows[0]
        html = '<h2>SSRS CSV 欄位名稱</h2><ol>'
        for i, col in enumerate(header):
            html += f'<li><b>{col}</b> (repr: {repr(col)})</li>'
        html += '</ol>'
        if len(rows) > 1:
            html += '<h3>第一筆資料</h3><table border="1" cellpadding="4"><tr><th>欄位</th><th>值</th></tr>'
            for col, val in zip(header, rows[1]):
                html += f'<tr><td>{col}</td><td>{val}</td></tr>'
            html += '</table>'
        return html
    except Exception as e:
        import traceback
        return f'<pre>Error: {traceback.format_exc()}</pre>', 500


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


# ── PDM 索引重建 ───────────────────────────────────────────────────────────────

import threading as _threading
import subprocess as _subprocess
import re as _re

_reindex_state = {
    'running':    False,
    'phase':      'idle',   # idle | scanning | indexing | done | error
    'scanned':    0,
    'total':      0,
    'indexed':    0,
    'message':    '',
    'error':      '',
    'last_run':   None,
    'last_count': 0,
}
_reindex_lock = _threading.Lock()

# 持久化狀態檔：重啟後仍可讀取最後一次 reindex 結果
_REINDEX_STATE_FILE = os.path.join(_APP_DIR, 'last_reindex.json')


def _save_last_reindex_state(ts: str, cnt: int):
    """reindex 完成後將時間與筆數寫入 JSON（重啟後仍保留）"""
    try:
        with open(_REINDEX_STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump({'last_run': ts, 'last_count': cnt}, f)
    except Exception:
        pass


def _load_last_reindex_state():
    """從 JSON 讀取最後一次 reindex 結果（比 DB indexed_at 更可靠）"""
    try:
        with open(_REINDEX_STATE_FILE, encoding='utf-8') as f:
            d = json.load(f)
            return d.get('last_run'), int(d.get('last_count', 0))
    except Exception:
        return None, 0


def _run_reindex(update_only: bool):
    """背景執行索引重建，解析 stdout 更新進度狀態"""
    global _reindex_state
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'build_pdm_index.py')
    cmd = [sys.executable, script] + (['--update'] if update_only else [])

    with _reindex_lock:
        _reindex_state.update(running=True, phase='scanning', scanned=0,
                              total=0, indexed=0, message='啟動中...', error='')

    try:
        proc = _subprocess.Popen(
            cmd, stdout=_subprocess.PIPE, stderr=_subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace'
        )
        for line in proc.stdout:
            line = line.rstrip()
            if not line:
                continue
            with _reindex_lock:
                _reindex_state['message'] = line

            # 解析掃描進度：「掃描中... 1,200 個檔案」
            m = _re.search(r'掃描中.*?(\d[\d,]*)\s*個', line)
            if m:
                with _reindex_lock:
                    _reindex_state['phase']   = 'scanning'
                    _reindex_state['scanned'] = int(m.group(1).replace(',', ''))
                continue

            # 解析找到總數：「找到 X 個 .SLDDRW」
            m = _re.search(r'找到\s*(\d[\d,]*)\s*個', line)
            if m:
                with _reindex_lock:
                    _reindex_state['total'] = int(m.group(1).replace(',', ''))
                continue

            # 解析寫入進度：「 33%  1000/3000  新增:X」
            m = _re.search(r'(\d+)%.*?(\d[\d,]*)/(\d[\d,]*)', line)
            if m:
                with _reindex_lock:
                    _reindex_state['phase']   = 'indexing'
                    _reindex_state['indexed'] = int(m.group(2).replace(',', ''))
                    _reindex_state['total']   = int(m.group(3).replace(',', ''))
                continue

            # 完成行：「完成！新增:X  更新:X  ...」
            m = _re.search(r'完成.*新增:(\d[\d,]*)', line)
            if m:
                with _reindex_lock:
                    _reindex_state['indexed'] = int(m.group(1).replace(',', ''))

        proc.wait()

        if proc.returncode == 0:
            # 取得資料庫最新筆數
            try:
                conn_tmp = get_pdm_db()
                if conn_tmp:
                    cnt = conn_tmp.execute('SELECT COUNT(*) FROM drawing_index').fetchone()[0]
                    conn_tmp.close()
                else:
                    cnt = _reindex_state['indexed']
            except Exception:
                cnt = _reindex_state['indexed']

            ts_done = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with _reindex_lock:
                _reindex_state.update(
                    running=False, phase='done', last_count=cnt,
                    last_run=ts_done,
                    message=f'更新完成，共 {cnt:,} 筆圖面資料'
                )
            # 持久化到 JSON，重啟後仍可讀取
            _save_last_reindex_state(ts_done, cnt)
        else:
            with _reindex_lock:
                _reindex_state.update(running=False, phase='error',
                                      error='重建失敗，請查看伺服器日誌')
    except Exception as exc:
        with _reindex_lock:
            _reindex_state.update(running=False, phase='error', error=str(exc))


@app.route('/api/drawing/reindex', methods=['POST'])
def drawing_reindex():
    """啟動 PDM 圖面索引重建（背景執行）"""
    with _reindex_lock:
        if _reindex_state['running']:
            return jsonify({'success': False, 'error': '索引重建已在執行中，請稍候'}), 409

    data        = request.get_json() or {}
    update_only = data.get('update_only', True)   # 預設增量更新

    t = _threading.Thread(target=_run_reindex, args=(update_only,), daemon=True)
    t.start()
    return jsonify({'success': True, 'message': '索引重建已啟動'})


def _pdm_last_indexed():
    """從 drawing_index 讀取最後一次 indexed_at（持久，不受重啟影響）"""
    try:
        c = get_pdm_db()
        if not c:
            return None, 0
        # 舊版 DB 可能沒有 indexed_at 欄位，ALTER TABLE 補上（冪等）
        try:
            c.execute("ALTER TABLE drawing_index ADD COLUMN indexed_at TEXT")
        except Exception:
            pass  # 欄位已存在
        row = c.execute(
            "SELECT MAX(indexed_at), COUNT(*) FROM drawing_index"
        ).fetchone()
        c.close()
        return (row[0] or None), (row[1] or 0)
    except Exception:
        return None, 0


@app.route('/api/drawing/reindex/status', methods=['GET'])
def drawing_reindex_status():
    """回傳索引重建進度狀態（含 DB 持久化的 last_run）"""
    with _reindex_lock:
        state = dict(_reindex_state)
    # 若記憶體中 last_run 為空（剛重啟），依序從 JSON → DB 讀取並快取
    if not state.get('last_run'):
        ts, cnt = _load_last_reindex_state()    # 優先：JSON 狀態檔
        if not ts:
            ts, cnt = _pdm_last_indexed()        # 備援：DB indexed_at
        if ts:
            with _reindex_lock:
                _reindex_state['last_run']   = ts
                _reindex_state['last_count'] = cnt
            state['last_run']   = ts
            state['last_count'] = cnt
    return jsonify(state)


@app.route('/api/zume/status')
def zume_status():
    """回傳技術資料清單的目前狀態（件數、最後匯入檔案）"""
    try:
        con = sqlite3.connect(ZUME_DB_PATH)
        total = con.execute('SELECT COUNT(*) FROM drawings').fetchone()[0]
        last  = con.execute('SELECT filename, imported_at FROM import_log ORDER BY imported_at DESC LIMIT 1').fetchone()
        con.close()
        return jsonify({'total': total, 'last_file': last[0] if last else None, 'last_at': last[1] if last else None})
    except Exception as e:
        return jsonify({'total': 0, 'last_file': None, 'last_at': None, 'error': str(e)})


@app.route('/api/zume/scan', methods=['POST'])
def zume_scan():
    """重新掃描 Downloads 資料夾並匯入最新 CSV（強制重掃）"""
    import glob as _glob
    downloads = os.path.join(os.path.expanduser('~'), 'Downloads')
    pattern   = os.path.join(downloads, 'zume-n_data_list_*.csv')
    files     = sorted(_glob.glob(pattern), key=os.path.getmtime, reverse=True)
    if not files:
        return jsonify({'success': False, 'error': f'Downloads 資料夾找不到 zume-n_data_list_*.csv\n路徑：{downloads}'}), 404
    latest = files[0]
    fname  = os.path.basename(latest)
    rows   = _parse_zume_csv(latest)
    if not rows:
        return jsonify({'success': False, 'error': f'檔案解析失敗或無資料：{fname}'}), 400
    con = sqlite3.connect(ZUME_DB_PATH)
    con.executemany('INSERT OR REPLACE INTO drawings(part_no,part_name,url) VALUES(?,?,?)', rows)
    con.execute('INSERT OR REPLACE INTO import_log(filename,imported_at,count) VALUES(?,datetime("now"),?)',
                (fname, len(rows)))
    con.commit()
    total = con.execute('SELECT COUNT(*) FROM drawings').fetchone()[0]
    con.close()
    return jsonify({'success': True, 'imported': len(rows), 'total': total, 'file': fname})


@app.route('/api/zume/import', methods=['POST'])
def zume_import():
    """手動上傳 CSV/XLSX 匯入（保留作為備用方式）"""
    f = request.files.get('file')
    if not f:
        return jsonify({'success': False, 'error': '未上傳檔案'}), 400
    fname_lower = f.filename.lower()
    rows = []
    try:
        if fname_lower.endswith('.csv'):
            content = f.read().decode('utf-8-sig')
            reader  = csv.reader(io.StringIO(content))
            headers = [h.strip() for h in next(reader)]
            idx_no  = next((i for i,h in enumerate(headers) if '圖號' in h), None)
            idx_url = next((i for i,h in enumerate(headers) if h.upper() == 'URL'), None)
            idx_nm  = next((i for i,h in enumerate(headers) if '品名' in h), 1)
            if idx_no is None or idx_url is None:
                return jsonify({'success': False, 'error': f'找不到「圖號」或「URL」欄'}), 400
            for row in reader:
                if len(row) > max(idx_no, idx_url):
                    no  = row[idx_no].strip(); url = row[idx_url].strip()
                    nm  = row[idx_nm].strip() if idx_nm < len(row) else ''
                    if no and url.startswith('http'):
                        rows.append((no, nm, url))
        elif fname_lower.endswith(('.xlsx', '.xlsm')):
            if not _OPENPYXL_OK:
                return jsonify({'success': False, 'error': '請改用 CSV 格式'}), 400
            wb = openpyxl.load_workbook(io.BytesIO(f.read()), read_only=True)
            ws = wb.active
            headers = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
            idx_no  = next((i for i,h in enumerate(headers) if '圖號' in h), None)
            idx_url = next((i for i,h in enumerate(headers) if h.upper() == 'URL'), None)
            idx_nm  = next((i for i,h in enumerate(headers) if '品名' in h), 1)
            if idx_no is None or idx_url is None:
                return jsonify({'success': False, 'error': '找不到「圖號」或「URL」欄'}), 400
            for row in ws.iter_rows(min_row=2, values_only=True):
                no  = str(row[idx_no] or '').strip()
                url = str(row[idx_url] or '').strip()
                nm  = str(row[idx_nm]  or '').strip() if idx_nm < len(row) else ''
                if no and url.startswith('http'):
                    rows.append((no, nm, url))
            wb.close()
        else:
            return jsonify({'success': False, 'error': '僅支援 .csv 或 .xlsx'}), 400
    except Exception as e:
        return jsonify({'success': False, 'error': f'解析失敗：{e}'}), 500
    if not rows:
        return jsonify({'success': False, 'error': '檔案無有效資料'}), 400
    fname = f.filename
    con = sqlite3.connect(ZUME_DB_PATH)
    con.executemany('INSERT OR REPLACE INTO drawings(part_no,part_name,url) VALUES(?,?,?)', rows)
    con.execute('INSERT OR REPLACE INTO import_log(filename,imported_at,count) VALUES(?,datetime("now"),?)',
                (fname, len(rows)))
    con.commit()
    total = con.execute('SELECT COUNT(*) FROM drawings').fetchone()[0]
    con.close()
    return jsonify({'success': True, 'imported': len(rows), 'total': total})


@app.route('/api/zume/lookup', methods=['POST'])
def zume_lookup():
    """查詢品號對應的 zume-n.com URL；找不到則回傳搜尋 URL"""
    try:
        data     = request.get_json() or {}
        item_nos = data.get('item_nos', [])
        if not item_nos:
            return jsonify({'success': False, 'error': '請提供品號'}), 400
        con = sqlite3.connect(ZUME_DB_PATH)
        result = []
        for no in item_nos[:20]:
            no = no.strip()
            row = con.execute('SELECT part_name,url FROM drawings WHERE part_no=?', (no,)).fetchone()
            if row:
                result.append({'no': no, 'name': row[0], 'url': row[1], 'found': True})
            else:
                fallback_url = 'https://zume-n.com/freeword_search?query=' + urllib.parse.quote(no, safe='') + '&searchType=drawing'
                result.append({'no': no, 'name': '', 'url': fallback_url, 'found': False})
        con.close()
        return jsonify({'success': True, 'items': result})
    except Exception as e:
        return jsonify({'success': False, 'error': f'查詢失敗：{str(e)}'}), 500


@app.route('/api/zume/open', methods=['POST'])
def zume_open():
    """以系統預設瀏覽器開啟 zume-n.com（優先用清單的直接 URL）"""
    try:
        data     = request.get_json() or {}
        item_nos = data.get('item_nos', [])
        if not item_nos:
            return jsonify({'success': False, 'error': '請提供品號'}), 400
        con = sqlite3.connect(ZUME_DB_PATH)
        opened = []; not_found = []
        for no in item_nos[:10]:
            no = no.strip()
            row = con.execute('SELECT url FROM drawings WHERE part_no=?', (no,)).fetchone()
            url = row[0] if row else ('https://zume-n.com/freeword_search?query=' + urllib.parse.quote(no, safe='') + '&searchType=drawing')
            webbrowser.open(url, new=2)
            if row:
                opened.append(no)
            else:
                not_found.append(no)
        con.close()
        return jsonify({'success': True, 'opened': opened, 'not_found': not_found})
    except Exception as e:
        return jsonify({'success': False, 'error': f'開啟失敗：{str(e)}'}), 500


@app.route('/equipment')
def equipment_page():
    return render_template('equipment.html', app_version=APP_VERSION)


@app.route('/api/equipment/machines')
def equipment_machines():
    try:
        machines = servcloud_get_machines()
        return jsonify({'success': True, 'data': machines})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/equipment/history')
def equipment_history():
    date_str    = request.args.get('date', '')           # YYYY-MM-DD
    start_hour  = int(request.args.get('start_hour', 7))
    end_hour    = int(request.args.get('end_hour', 20))
    machine_ids = request.args.getlist('machines')

    if not date_str or not machine_ids:
        return jsonify({'success': False, 'error': '請提供日期和機台'}), 400

    date_api = date_str.replace('-', '')

    try:
        history = servcloud_get_history(machine_ids, date_api)
        all_machines = servcloud_get_machines()
        name_map = {m['device_id']: m['device_name'] for m in all_machines}

        range_start_s = start_hour * 3600   # 秒
        range_end_s   = end_hour   * 3600
        total_min     = (range_end_s - range_start_s) // 60

        # Group segments by machine (秒精度)
        buckets = {mid: [] for mid in machine_ids}
        for rec in history:
            mid = rec.get('machine_id', '')
            if mid not in buckets:
                continue
            st = rec.get('start_time', '')
            et = rec.get('end_time',   '')
            if len(st) < 14 or len(et) < 14:
                continue
            s_sec = int(st[8:10])*3600 + int(st[10:12])*60 + int(st[12:14])
            e_sec = int(et[8:10])*3600 + int(et[10:12])*60 + int(et[12:14])
            # Handle overnight
            if e_sec < s_sec:
                e_sec += 24 * 3600
            s_clip = max(s_sec, range_start_s)
            e_clip = min(e_sec, range_end_s)
            if s_clip >= e_clip:
                continue
            status = rec.get('status', '0')
            status_name, color = _STATUS_MAP.get(status, ('未知', '#9e9e9e'))
            dur_sec = e_clip - s_clip
            buckets[mid].append({
                'status':      status,
                'status_name': status_name,
                'color':       color,
                'start_min':   (s_clip - range_start_s) / 60,   # float，用於圖表定位
                'end_min':     (e_clip - range_start_s) / 60,
                'start_str':   f"{s_clip//3600%24:02d}:{s_clip%3600//60:02d}:{s_clip%60:02d}",
                'end_str':     f"{e_clip//3600%24:02d}:{e_clip%3600//60:02d}:{e_clip%60:02d}",
                'dur_sec':     dur_sec,
            })

        result = [
            {'id': mid, 'name': name_map.get(mid, mid), 'segments': buckets[mid]}
            for mid in machine_ids
        ]
        return jsonify({
            'success': True, 'machines': result,
            'date': date_str, 'start_hour': start_hour,
            'end_hour': end_hour, 'total_minutes': total_min
        })
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500


# ── BOM 查詢系統 ───────────────────────────────────────────────────
# 資料庫：192.168.1.140 / YC01（Computech ERP）
# 品號主檔：INVMB  材料BOM：BOMMD（MCPI11模組）
_ERP_SQL_SERVER   = getattr(config, 'ERP_SQL_SERVER',   '192.168.1.140')
_ERP_SQL_DATABASE = getattr(config, 'ERP_SQL_DATABASE', 'YC01')
_ERP_SQL_USERNAME = getattr(config, 'ERP_SQL_USERNAME', 'sa')
_ERP_SQL_PASSWORD = getattr(config, 'ERP_SQL_PASSWORD', 'dsc55877948')

# 屬性代碼對照表（INVMB.MB025）
_ATTR_MAP = {'M': 'M:自製件', 'S': 'S:訂外加工', 'P': 'P:採購件',
             'R': 'R:委外', 'F': 'F:虛擬件'}


def get_erp_conn():
    """建立 ERP SQL Server 連線（pyodbc）"""
    try:
        import pyodbc
    except ImportError:
        raise RuntimeError('未安裝 pyodbc，請執行: pip install pyodbc')

    cs = (f'DRIVER={{ODBC Driver 17 for SQL Server}};'
          f'SERVER={_ERP_SQL_SERVER};DATABASE={_ERP_SQL_DATABASE};'
          f'UID={_ERP_SQL_USERNAME};PWD={_ERP_SQL_PASSWORD};'
          f'TrustServerCertificate=yes;Connection Timeout=8;')
    return pyodbc.connect(cs, timeout=8)


@app.route('/routing')
def routing_page():
    return render_template('routing.html', app_version=APP_VERSION)


@app.route('/api/routing/search')
def routing_search():
    """搜尋有途程的品號（INVMB + BOMME，多關鍵字空格分隔，品號/品名同時模糊搜尋）
       line 參數：指定線別篩選（MF006），空字串代表全部"""
    q    = request.args.get('q', '').strip()
    line = request.args.get('line', '').strip()   # '000-1','000-2','000-3' 或 ''=全部
    if not q:
        return jsonify({'success': False, 'error': '請輸入品號或品名'}), 400
    try:
        conn = get_erp_conn()
        cur  = conn.cursor()
        # 多關鍵字：以空格分割，每個關鍵字都要符合（AND 邏輯）
        keywords = [kw for kw in q.split() if kw]
        where_parts = []
        params = []
        for kw in keywords:
            like = f'%{kw}%'
            where_parts.append("(RTRIM(b.MB001) LIKE ? OR b.MB002 LIKE ?)")
            params.extend([like, like])
        keyword_clause = ' AND '.join(where_parts) if where_parts else '1=1'

        # 線別篩選：只顯示「最新版次」途程中包含該線別至少一步的品號
        if line:
            line_clause = """AND EXISTS (
                SELECT 1 FROM BOMMF f2
                WHERE RTRIM(f2.MF001) = RTRIM(b.MB001)
                  AND f2.MF002 = (
                      SELECT MAX(e2.ME002) FROM BOMME e2
                      WHERE RTRIM(e2.ME001) = RTRIM(b.MB001)
                  )
                  AND RTRIM(f2.MF006) = ?
            )"""
            params.append(line)
        else:
            line_clause = ''

        cur.execute(f"""
            SELECT TOP 200
                RTRIM(b.MB001)            AS item_no,
                RTRIM(ISNULL(b.MB002,'')) AS item_name,
                RTRIM(ISNULL(b.MB003,'')) AS spec,
                ISNULL(b.MB004,'')        AS unit,
                ISNULL(b.MB025,'')        AS attr_code
            FROM INVMB b
            WHERE EXISTS (
                SELECT 1 FROM BOMME e WHERE RTRIM(e.ME001) = RTRIM(b.MB001)
            )
            AND {keyword_clause}
            {line_clause}
            ORDER BY b.MB001
        """, params)
        rows = cur.fetchall()
        conn.close()
        results = [{
            'item_no':   r[0],
            'item_name': r[1],
            'spec':      r[2],
            'unit':      r[3],
            'attr':      _ATTR_MAP.get(str(r[4]).strip(), str(r[4]).strip()),
        } for r in rows]
        return jsonify({'success': True, 'data': results, 'count': len(results)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/bom')
def bom_page():
    # 舊連結相容：直接導向產品途程查詢頁（BOM子分頁）
    return redirect('/routing?tab=bom')


@app.route('/api/bom/search')
def bom_search():
    """搜尋有材料BOM的品號（INVMB + BOMMD，品號/品名模糊搜尋，可篩選品號開頭）"""
    q      = request.args.get('q', '').strip()
    prefix = request.args.get('prefix', '').strip()   # '5', '6', '56' 或空
    if not q:
        return jsonify({'success': False, 'error': '請輸入品號或品名'}), 400
    try:
        conn = get_erp_conn()
        cur  = conn.cursor()
        like = f'%{q}%'

        # 品號開頭篩選條件
        if prefix == '5':
            prefix_cond = "AND RTRIM(b.MB001) LIKE '5%'"
        elif prefix == '6':
            prefix_cond = "AND RTRIM(b.MB001) LIKE '6%'"
        elif prefix in ('56', '65'):
            prefix_cond = "AND (RTRIM(b.MB001) LIKE '5%' OR RTRIM(b.MB001) LIKE '6%')"
        else:
            prefix_cond = ''

        cur.execute(f"""
            SELECT TOP 200
                RTRIM(b.MB001)          AS item_no,
                RTRIM(ISNULL(b.MB002,'')) AS item_name,
                RTRIM(ISNULL(b.MB003,'')) AS spec,
                ISNULL(b.MB004,'')      AS unit,
                ISNULL(b.MB025,'')      AS attr_code
            FROM INVMB b
            WHERE EXISTS (
                SELECT 1 FROM BOMMD d WHERE RTRIM(d.MD001) = RTRIM(b.MB001)
            )
            {prefix_cond}
            AND (RTRIM(b.MB001) LIKE ? OR b.MB002 LIKE ?)
            ORDER BY b.MB001
        """, (like, like))
        rows = cur.fetchall()
        conn.close()
        results = [{
            'item_no':   r[0],
            'item_name': r[1],
            'spec':      r[2],
            'unit':      r[3],
            'attr':      _ATTR_MAP.get(str(r[4]).strip(), str(r[4]).strip()),
        } for r in rows]
        return jsonify({'success': True, 'data': results, 'count': len(results)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/bom/detail')
def bom_detail():
    """取得指定品號的材料BOM明細（BOMMD JOIN INVMB，對應 MCPI11 BOM用量資料）"""
    item_no = request.args.get('item_no', '').strip()
    if not item_no:
        return jsonify({'success': False, 'error': '請提供品號'}), 400
    try:
        conn = get_erp_conn()
        cur  = conn.cursor()

        # 取品號主檔（母件資訊）
        cur.execute("""
            SELECT RTRIM(MB001), RTRIM(ISNULL(MB002,'')),
                   RTRIM(ISNULL(MB003,'')), ISNULL(MB004,''),
                   ISNULL(MB005,''), ISNULL(MB025,'')
            FROM INVMB WHERE RTRIM(MB001) = ?
        """, (item_no,))
        hrow = cur.fetchone()
        header = {}
        if hrow:
            header = {
                'item_no':   hrow[0],
                'item_name': hrow[1],
                'spec':      hrow[2],
                'unit':      hrow[3],
                'sub_unit':  hrow[4],
                'attr':      _ATTR_MAP.get(str(hrow[5]).strip(), str(hrow[5]).strip()),
            }

        # 取材料BOM明細（BOMMD JOIN INVMB 取子件品名/規格/屬性）
        cur.execute("""
            SELECT
                d.MD002                           AS seq,
                RTRIM(d.MD003)                    AS child_no,
                RTRIM(ISNULL(b.MB002,''))         AS child_name,
                RTRIM(ISNULL(b.MB003,''))         AS spec,
                ISNULL(d.MD004,'')                AS unit,
                ISNULL(d.MD005,'')                AS sub_unit,
                ISNULL(b.MB025,'')                AS attr_code,
                CAST(ISNULL(d.MD006,0) AS VARCHAR(20)) AS qty
            FROM BOMMD d
            LEFT JOIN INVMB b ON RTRIM(b.MB001) = RTRIM(d.MD003)
            WHERE RTRIM(d.MD001) = ?
            ORDER BY d.MD002
        """, (item_no,))
        rows = cur.fetchall()
        conn.close()

        detail = []
        for r in rows:
            attr_code = str(r[6] or '').strip()
            detail.append({
                'seq':        str(r[0] or '').strip(),
                'child_no':   str(r[1] or '').strip(),
                'child_name': str(r[2] or '').strip(),
                'spec':       str(r[3] or '').strip(),
                'unit':       str(r[4] or '').strip(),
                'sub_unit':   str(r[5] or '').strip(),
                'attr':       _ATTR_MAP.get(attr_code, attr_code),
                'qty':        str(r[7] or '').strip(),
            })

        return jsonify({'success': True, 'header': header, 'detail': detail})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/bom/routing')
def bom_routing():
    """取得指定品號的產品途程（BOMME + BOMMF，最新版次，對應 BOMMI07）"""
    item_no = request.args.get('item_no', '').strip()
    if not item_no:
        return jsonify({'success': False, 'error': '請提供品號'}), 400
    try:
        conn = get_erp_conn()
        cur  = conn.cursor()

        # 品號主檔
        cur.execute("""
            SELECT RTRIM(MB001), RTRIM(ISNULL(MB002,'')),
                   RTRIM(ISNULL(MB003,'')), ISNULL(MB004,''), ISNULL(MB025,'')
            FROM INVMB WHERE RTRIM(MB001) = ?
        """, (item_no,))
        hrow = cur.fetchone()
        header = {}
        if hrow:
            header = {
                'item_no':   hrow[0],
                'item_name': hrow[1],
                'spec':      hrow[2],
                'unit':      hrow[3],
                'attr':      _ATTR_MAP.get(str(hrow[4]).strip(), str(hrow[4]).strip()),
            }

        # 最新途程版次
        cur.execute("""
            SELECT MAX(ME002) FROM BOMME WHERE RTRIM(ME001) = ?
        """, (item_no,))
        vrow = cur.fetchone()
        latest_ver = (vrow[0] or '').strip() if vrow else ''
        if header:
            header['version'] = latest_ver

        # 途程明細（BOMMF JOIN CMSMW 取製程名稱/機台代號/資源群組代號）
        routing = []
        if latest_ver:
            _NATURE = {'1': '1:廠內', '2': '2:訂外', '3': '3:外包'}
            cur.execute("""
                SELECT
                    f.MF003,
                    RTRIM(ISNULL(f.MF004,'')),
                    RTRIM(ISNULL(w.MW002,''))   AS proc_name,
                    f.MF005,
                    RTRIM(ISNULL(f.MF006,'')),
                    RTRIM(ISNULL(f.MF007,'')),
                    RTRIM(ISNULL(f.MF008,'')),
                    RTRIM(ISNULL(f.MF040,'')),
                    RTRIM(ISNULL(f.MF034,'')),
                    CAST(ISNULL(f.MF010,0) AS DECIMAL(10,3)),
                    RTRIM(ISNULL(x.MX003,''))   AS machine_name
                FROM BOMMF f
                LEFT JOIN CMSMW  w ON RTRIM(w.MW001) = RTRIM(f.MF004)
                LEFT JOIN CMSMX  x ON RTRIM(x.MX001) = RTRIM(f.MF040)
                WHERE RTRIM(f.MF001) = ? AND f.MF002 = ?
                ORDER BY f.MF003
            """, (item_no, latest_ver))
            rows = cur.fetchall()
            for r in rows:
                nature_code = str(r[3] or '').strip()
                vmh_raw = r[9]
                if vmh_raw is None or float(vmh_raw) == 0:
                    vmh = ''
                else:
                    vmh = '{:g}'.format(float(vmh_raw))
                routing.append({
                    'seq':          str(r[0] or '').strip(),
                    'proc_code':    str(r[1] or '').strip(),
                    'proc_name':    str(r[2] or '').strip(),
                    'nature':       _NATURE.get(nature_code, nature_code),
                    'vendor_no':    str(r[4] or '').strip(),
                    'vendor_name':  str(r[5] or '').strip(),
                    'description':  str(r[6] or '').strip(),
                    'machine_code': str(r[7] or '').strip(),
                    'res_group':    str(r[8] or '').strip(),
                    'var_man_hrs':  vmh,
                    'machine_name': str(r[10] or '').strip(),
                })

        conn.close()
        has_routing = bool(latest_ver and routing)
        return jsonify({
            'success': True,
            'has_routing': has_routing,
            'header': header,
            'routing': routing,
        })

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
