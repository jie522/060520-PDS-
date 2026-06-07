#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDM 圖面索引重建腳本
透過 PDM COM API 掃描 MAXCLAW Vault，重建 pdm_search.db 索引資料庫。

用法：
  python build_pdm_index.py              # 完整重建
  python build_pdm_index.py --update     # 增量更新（只新增 Vault 中不存在於 DB 的檔案）
  python build_pdm_index.py --deploy     # 重建後自動複製到 dist_embed 目錄
  python build_pdm_index.py --update --deploy

說明：
  - 品號 (pin_hao)  : 取自檔名（去除 .SLDDRW）
  - 機型 (ji_xing)  : 從資料夾路徑推導（如 FTD\\FTD430B\\ -> FTD430B）
  - 發圖日期        : 從 PDM SQL DB 直查，失敗改用本機快取檔案日期
  - 圖面品名        : 從 PDM SQL DB 直查
"""

import os
import sys
import re
import shutil
import sqlite3
import datetime
import winreg
import win32com.client

# ══════════════════════════════════════════════════════════
#  設定
# ══════════════════════════════════════════════════════════

VAULT_NAME  = 'MAXCLAW'
TARGET_EXT  = '.SLDDRW'
TARGET_VARS = ['品名', '發圖日期']

DB_DIR  = os.path.join(os.environ.get('LOCALAPPDATA', r'C:\Users\Public'), 'PDMSearch')
DB_PATH = os.path.join(DB_DIR, 'pdm_search.db')

# 機型資料夾名稱正規表達式（如 FTD430B, FTB245, MSB300）
MACHINE_RE = re.compile(r'^[A-Za-z]{2,6}\d{2,6}[A-Za-z]{0,3}$')

# ══════════════════════════════════════════════════════════
#  PDM COM 工具函式
# ══════════════════════════════════════════════════════════

def connect_vault():
    """連線 PDM Vault。
    優先使用 gencache.EnsureDispatch（早期繫結），讓 win32com 載入型別庫後
    所有回傳的 COM 物件都能正確 QI 到高版本介面（IEdmFile5 等），
    使 GetVar / GetVarStr 等方法可用。
    """
    try:
        from win32com.client import gencache
        v = gencache.EnsureDispatch('ConisioLib.EdmVault')
    except Exception:
        v = win32com.client.Dispatch('ConisioLib.EdmVault')
    v.LoginAuto(VAULT_NAME, 0)
    return v


def derive_ji_xing(folder_local_path):
    """
    從資料夾路徑推導機型。
    例：C:\\MAXCLAW\\00-研發部\\01-正式圖面\\01-產品圖\\FTD\\FTD430B  -> FTD430B
    """
    parts = folder_local_path.replace('/', '\\').split('\\')
    for part in reversed(parts):
        if MACHINE_RE.match(part):
            return part
    return parts[-1] if parts else ''


def file_mdate(local_path):
    """取得本機檔案修改日期（格式 YYYY/MM/DD），失敗回傳空字串"""
    try:
        mt = os.path.getmtime(local_path)
        return datetime.datetime.fromtimestamp(mt).strftime('%Y/%m/%d')
    except Exception:
        return ''


# ══════════════════════════════════════════════════════════
#  PDM SQL Server 直連（ADO via win32com）
# ══════════════════════════════════════════════════════════

def _find_sql_creds_from_registry(vault_name):
    """從 Windows Registry 找 PDM SQL Server 名稱及認證資訊
    回傳 dict: {server, user, password, database}
    """
    paths = [
        rf'SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise\Databases\{vault_name}',
        rf'SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise\Vaults\{vault_name}',
        rf'SOFTWARE\SolidWorks\PDMWorks Enterprise\Databases\{vault_name}',
        rf'SOFTWARE\Wow6432Node\SolidWorks\Applications\PDMWorks Enterprise\Databases\{vault_name}',
    ]
    creds = {}
    for rpath in paths:
        for hive in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            try:
                key = winreg.OpenKey(hive, rpath)
                # 嘗試讀取所有值（診斷用）
                try:
                    i = 0
                    while True:
                        name, data, _ = winreg.EnumValue(key, i)
                        creds[name] = data
                        i += 1
                except OSError:
                    pass
            except Exception:
                pass
    return creds


def load_vars_from_sql(vault_name=VAULT_NAME, log=None):
    """
    直接透過 ADO 連線 PDM SQL Server，讀取所有 SLDDRW 的品名與發圖日期。
    回傳 dict: {filename_lower: {'品名': str, '發圖日期': str}}
    """
    def _log(msg):
        print(msg)
        if log:
            log.write(msg + '\n')

    # 從 Registry 讀取所有 PDM 資料庫設定值
    reg_creds = _find_sql_creds_from_registry(vault_name)
    _log(f'  [SQL] Registry 所有值: {reg_creds}')

    hostname   = os.environ.get('COMPUTERNAME', 'localhost')
    reg_server = reg_creds.get('Server') or reg_creds.get('SQLServer') or reg_creds.get('DbServer')
    reg_user   = reg_creds.get('User') or reg_creds.get('Username') or reg_creds.get('Login')
    reg_pass   = reg_creds.get('Password') or reg_creds.get('Pwd')
    reg_db     = reg_creds.get('Database') or reg_creds.get('DbName') or vault_name

    candidates = []
    if reg_server:
        candidates.append(reg_server)
    candidates += [
        f'{hostname}\\SOLIDWORKSPDM',
        '(local)\\SOLIDWORKSPDM',
        f'{hostname}\\SQLEXPRESS',
        '(local)\\SQLEXPRESS',
        '(local)',
        hostname,
    ]
    seen = set()
    candidates = [c for c in candidates if not (c in seen or seen.add(c))]

    _log(f'  [SQL] 嘗試伺服器: {candidates}')
    if reg_user:
        _log(f'  [SQL] 找到 Registry 帳號: {reg_user}')

    var_list_sql = ', '.join(f"N'{v}'" for v in TARGET_VARS)

    QUERIES = [
        f"""SELECT d.Filename, v.VariableName, vv.ValueText
FROM VariableValue vv
JOIN Variable v ON v.VariableID = vv.VariableID
JOIN Document d ON d.DocumentID = vv.DocumentID
WHERE v.VariableName IN ({var_list_sql})
  AND d.Filename LIKE N'%.SLDDRW'""",
        f"""SELECT d.Filename, v.Name AS VariableName, vv.Value AS ValueText
FROM VariableValue vv
JOIN Variable v ON v.ID = vv.VariableID
JOIN Document d ON d.ID = vv.DocumentID
WHERE v.Name IN ({var_list_sql})
  AND d.Filename LIKE N'%.SLDDRW'""",
    ]

    def _try_server(server, user, password):
        cn = win32com.client.Dispatch('ADODB.Connection')
        if user:
            conn_str = (f'Provider=SQLOLEDB;Data Source={server};'
                        f'Initial Catalog={reg_db};'
                        f'User ID={user};Password={password or ""};'
                        f'Connect Timeout=5')
        else:
            conn_str = (f'Provider=SQLOLEDB;Data Source={server};'
                        f'Initial Catalog={reg_db};Integrated Security=SSPI;'
                        f'Connect Timeout=5')
        cn.Open(conn_str)
        return cn

    for server in candidates:
        # 先試 SQL auth（如果有帳號），再試 Windows auth
        auth_combos = []
        if reg_user:
            auth_combos.append((reg_user, reg_pass))
        auth_combos.append((None, None))  # Windows auth

        for user, pwd in auth_combos:
            try:
                cn = _try_server(server, user, pwd)
                _log(f'  [SQL] 連線成功: {server} (user={user or "Windows auth"})')

                # 列出資料表（診斷）
                try:
                    rs_t = win32com.client.Dispatch('ADODB.Recordset')
                    rs_t.Open(
                        "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
                        "WHERE TABLE_TYPE='BASE TABLE' ORDER BY TABLE_NAME",
                        cn, 0, 1)
                    tables = []
                    while not rs_t.EOF:
                        tables.append(str(rs_t.Fields.Item(0).Value))
                        rs_t.MoveNext()
                    rs_t.Close()
                    _log(f'  [SQL] 資料表: {tables[:30]}')
                except Exception as et:
                    _log(f'  [SQL] 列表失敗: {et}')

                for q in QUERIES:
                    try:
                        rs = win32com.client.Dispatch('ADODB.Recordset')
                        rs.Open(q.strip(), cn, 0, 1)
                        result = {}
                        while not rs.EOF:
                            fname   = str(rs.Fields.Item('Filename').Value   or '').strip().lower()
                            varname = str(rs.Fields.Item('VariableName').Value or '').strip()
                            value   = str(rs.Fields.Item('ValueText').Value  or '').strip()
                            if fname and varname and value:
                                result.setdefault(fname, {})[varname] = value
                            rs.MoveNext()
                        rs.Close()
                        cn.Close()
                        _log(f'  [SQL] 查詢成功，取得 {len(result)} 筆檔案資料')
                        return result
                    except Exception as eq:
                        _log(f'  [SQL] 查詢失敗: {eq}')

                cn.Close()
                break
            except Exception as ec:
                _log(f'  [SQL] 連線失敗 {server} (user={user}): {ec}')

    _log('  [SQL] 所有伺服器均連線失敗，改用 COM GetFileFromPath 嘗試讀取')
    return {}


def load_vars_via_com(vault, results, log=None):
    """
    透過 IEdmEnumeratorVariable10 的 GetVarFromDb / StoreValuesFromDatabase
    讀取品名、發圖日期。這些方法會從 PDM Archive Server 取值，不需本機快取。
    回傳 dict: {filename_lower: {'品名': str, '發圖日期': str}}
    """
    def _log(msg):
        print(msg)
        if log:
            log.write(msg + '\n')

    _log('\n=== COM EnumeratorVariable10 診斷 ===')

    # ── 步驟 1: 測試前 10 筆 ──
    # 策略 A: CastTo IEdmEnumeratorVariable10 + GetVarFromDb(varName, cfgName)
    # 策略 B: StoreValuesFromDatabase(folderID) 拉值到快取 → GetVar(varName, cfgName)
    # 策略 C: GetVar2(varName, cfgName, folderID)
    _log('\n  --- 策略測試（前10筆）---')
    best_strategy = None
    test_results = {}

    for folder_path, fname in results[:10]:
        try:
            folder = vault.GetFolderFromPath(folder_path)
            file5 = folder.GetFile(fname)
            ev_early = file5.GetEnumeratorVariable()
            folder_id = folder.ID
            row = f'    {fname} folderID={folder_id}:'

            # --- 策略 A: CastTo IEdmEnumeratorVariable10 + GetVarFromDb ---
            if best_strategy is None or best_strategy == 'A':
                try:
                    ev10 = win32com.client.CastTo(ev_early, 'IEdmEnumeratorVariable10')
                    for vn in TARGET_VARS:
                        r = ev10.GetVarFromDb(vn, '@')
                        row += f' A:{vn}={repr(r)[:50]}'
                        if r and r[0] is True:
                            test_results.setdefault('A', []).append((fname, vn, r[1]))
                except Exception as e:
                    row += f' A:ERR={e}'

            # --- 策略 B: StoreValuesFromDatabase → GetVar ---
            if best_strategy is None or best_strategy == 'B':
                try:
                    ev10b = win32com.client.CastTo(ev_early, 'IEdmEnumeratorVariable10')
                    ev10b.StoreValuesFromDatabase(folder_id, False, 0)
                    for vn in TARGET_VARS:
                        r = ev10b.GetVar(vn, '@')
                        row += f' B:{vn}={repr(r)[:50]}'
                        if r and r[0] is True:
                            test_results.setdefault('B', []).append((fname, vn, r[1]))
                except Exception as e:
                    row += f' B:ERR={e}'

            # --- 策略 C: GetVar2 with folder ID ---
            if best_strategy is None or best_strategy == 'C':
                try:
                    ev10c = win32com.client.CastTo(ev_early, 'IEdmEnumeratorVariable10')
                    for vn in TARGET_VARS:
                        r = ev10c.GetVar2(vn, '@', folder_id)
                        row += f' C:{vn}={repr(r)[:50]}'
                        if r and r[0] is True:
                            test_results.setdefault('C', []).append((fname, vn, r[1]))
                except Exception as e:
                    row += f' C:ERR={e}'

            # --- 策略 D: GetVarAsText with folder ID ---
            if best_strategy is None or best_strategy == 'D':
                try:
                    ev10d = win32com.client.CastTo(ev_early, 'IEdmEnumeratorVariable10')
                    for vn in TARGET_VARS:
                        r = ev10d.GetVarAsText(vn, '@', folder_id)
                        row += f' D:{vn}={repr(r)[:50]}'
                        if r and r[0] is True:
                            test_results.setdefault('D', []).append((fname, vn, r[1]))
                except Exception as e:
                    row += f' D:ERR={e}'

            _log(row)
        except Exception as e:
            _log(f'    {fname}: 整體失敗: {e}')

    # 選出最佳策略
    for strat in ['A', 'B', 'C', 'D']:
        hits = test_results.get(strat, [])
        _log(f'  策略 {strat} 成功筆數: {len(hits)}')
        if hits:
            for fn, vn, val in hits[:3]:
                _log(f'    {fn}: {vn}={val}')

    for strat in ['A', 'B', 'C', 'D']:
        if test_results.get(strat):
            best_strategy = strat
            break

    if best_strategy is None:
        _log('  所有策略均無法讀取變數值')
        return {}

    _log(f'\n  選用策略: {best_strategy}')

    # ── 步驟 2: 批量讀取所有檔案 ──
    _log(f'\n  批量讀取 {len(results)} 筆 ...')
    result = {}
    folder_cache = {}
    err_count = 0

    for idx, (folder_path, fname) in enumerate(results, 1):
        try:
            if folder_path not in folder_cache:
                f_obj = vault.GetFolderFromPath(folder_path)
                folder_cache[folder_path] = (f_obj, f_obj.ID)
            folder, folder_id = folder_cache[folder_path]
            file5 = folder.GetFile(fname)
            ev_early = file5.GetEnumeratorVariable()
            ev10 = win32com.client.CastTo(ev_early, 'IEdmEnumeratorVariable10')

            if best_strategy == 'B':
                ev10.StoreValuesFromDatabase(folder_id, False, 0)

            fkey = fname.lower()
            for vn in TARGET_VARS:
                try:
                    if best_strategy == 'A':
                        r = ev10.GetVarFromDb(vn, '@')
                    elif best_strategy == 'B':
                        r = ev10.GetVar(vn, '@')
                    elif best_strategy == 'C':
                        r = ev10.GetVar2(vn, '@', folder_id)
                    elif best_strategy == 'D':
                        r = ev10.GetVarAsText(vn, '@', folder_id)

                    if r and r[0] is True:
                        val = r[1]
                        if val is not None:
                            s = str(val).strip()
                            if s:
                                result.setdefault(fkey, {})[vn] = s
                except Exception:
                    pass
        except Exception:
            err_count += 1

        if idx % 500 == 0:
            _log(f'    {idx}/{len(results)} ... (有值: {len(result)}, 錯誤: {err_count})')
            print(f'  {idx:,}/{len(results):,}')

    _log(f'\n  批量讀取完成: {len(result)} 筆有值, {err_count} 筆錯誤')
    return result


# ══════════════════════════════════════════════════════════
#  Vault 遞迴掃描
# ══════════════════════════════════════════════════════════

def scan_folder(folder, results, counters):
    """遞迴掃描 PDM 資料夾，收集 SLDDRW 檔案
    路徑在掃描階段就轉成字串，避免後續 COM 狀態污染導致 LocalPath 失效
    """
    ext = TARGET_EXT.upper()
    folder_path = str(folder.LocalPath)

    pos = folder.GetFirstFilePosition()
    while not pos.IsNull:
        f = folder.GetNextFile(pos)
        if f.Name.upper().endswith(ext):
            results.append((folder_path, str(f.Name)))
            counters[0] += 1
            if counters[0] % 200 == 0:
                print(f'  掃描中... {counters[0]:,} 個檔案')

    spos = folder.GetFirstSubFolderPosition()
    while not spos.IsNull:
        sub = folder.GetNextSubFolder(spos)
        scan_folder(sub, results, counters)


# ══════════════════════════════════════════════════════════
#  資料庫操作
# ══════════════════════════════════════════════════════════

SCHEMA = """
CREATE TABLE IF NOT EXISTS drawing_index (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path        TEXT NOT NULL,
    pin_hao          TEXT,
    tu_mian_pin_ming TEXT,
    ji_xing          TEXT,
    xing_hao         TEXT,
    modified_at      TEXT,
    indexed_at       TEXT DEFAULT (datetime('now')),
    thumbnail        BLOB
);
CREATE INDEX IF NOT EXISTS idx_pin_hao  ON drawing_index (pin_hao  COLLATE NOCASE);
CREATE INDEX IF NOT EXISTS idx_ji_xing  ON drawing_index (ji_xing);
CREATE INDEX IF NOT EXISTS idx_modified ON drawing_index (modified_at);
"""


def init_db(conn, full_rebuild):
    cur = conn.cursor()
    cur.executescript(SCHEMA)
    if full_rebuild:
        cur.execute('DELETE FROM drawing_index')
    conn.commit()


def get_existing_paths(conn):
    cur = conn.cursor()
    cur.execute('SELECT file_path FROM drawing_index')
    return {row[0] for row in cur.fetchall()}


def get_empty_paths(conn):
    cur = conn.cursor()
    cur.execute(
        "SELECT file_path FROM drawing_index "
        "WHERE tu_mian_pin_ming='' OR tu_mian_pin_ming IS NULL "
        "   OR modified_at='' OR modified_at IS NULL")
    return {row[0] for row in cur.fetchall()}


# ══════════════════════════════════════════════════════════
#  主流程
# ══════════════════════════════════════════════════════════

def rebuild(update_only=False, deploy=False):
    os.makedirs(DB_DIR, exist_ok=True)

    # 步驟 1：連線 Vault
    print('[1/4] 連線 PDM Vault ...')
    vault = connect_vault()
    print(f'  連線成功 | 根目錄：{vault.RootFolder.LocalPath}')

    # 步驟 2：遞迴掃描（只存路徑字串，不保留 COM 物件）
    print('[2/4] 掃描 Vault 中所有 SLDDRW 檔案 ...')
    results  = []
    counters = [0]
    scan_folder(vault.RootFolder, results, counters)
    total = len(results)
    print(f'  找到 {total:,} 個 {TARGET_EXT} 檔案')

    if total == 0:
        print('  未找到任何檔案，結束。')
        return 0

    # 步驟 3：初始化 DB
    print('[3/4] 初始化資料庫 ...')
    if os.path.exists(DB_PATH) and not update_only:
        ts  = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        bak = DB_PATH.replace('.db', f'_backup_{ts}.db')
        shutil.copy2(DB_PATH, bak)
        print(f'  備份 -> {os.path.basename(bak)}')

    conn = sqlite3.connect(DB_PATH)
    init_db(conn, full_rebuild=not update_only)
    cur  = conn.cursor()

    existing_paths = get_existing_paths(conn) if update_only else set()
    empty_paths    = get_empty_paths(conn)   if update_only else set()

    # ── 嘗試從 PDM SQL 資料庫直接讀取所有變數值 ──
    log_path = os.path.join(DB_DIR, 'pdm_debug.txt')
    print(f'  嘗試連線 PDM SQL Server ...')
    with open(log_path, 'w', encoding='utf-8') as lf:
        lf.write('=== PDM SQL 直連診斷 ===\n')
        sql_vars = load_vars_from_sql(log=lf)
        lf.write(f'\n取得 SQL 變數筆數: {len(sql_vars)}\n')
        sample = list(sql_vars.items())[:5]
        for fn, vals in sample:
            lf.write(f'  {fn}: {vals}\n')

        if not sql_vars:
            # SQL 失敗 → 改用 COM VariableMgr + GetVarFromDb
            sql_vars = load_vars_via_com(vault, results, log=lf)

    hit = sum(1 for v in sql_vars.values() if '品名' in v)
    print(f'  變數讀取完成：{len(sql_vars)} 筆，其中有品名：{hit} 筆')

    # 步驟 4：寫入索引
    print('[4/4] 寫入索引 ...')
    inserted = updated = skipped = errors = 0
    BATCH = 300

    for idx, (folder_path, fname) in enumerate(results, 1):
        file_path = os.path.join(folder_path, fname)

        if update_only and file_path in existing_paths and file_path not in empty_paths:
            skipped += 1
            continue

        try:
            pin_hao  = os.path.splitext(fname)[0]
            ji_xing  = derive_ji_xing(folder_path)
            xing_hao = ji_xing

            # 從 SQL 查詢結果取變數值（key 為小寫檔名）
            fname_key = fname.lower()
            sql_row   = sql_vars.get(fname_key, {})
            tu_mian_pin_ming = sql_row.get('品名', '').strip()
            modified_at      = sql_row.get('發圖日期', '').strip()

            # fallback：本機快取檔案修改日期
            if not modified_at:
                modified_at = file_mdate(file_path)

            now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            if update_only and file_path in existing_paths:
                cur.execute(
                    'UPDATE drawing_index SET pin_hao=?, tu_mian_pin_ming=?, '
                    'ji_xing=?, xing_hao=?, modified_at=?, indexed_at=? '
                    'WHERE file_path=?',
                    (pin_hao, tu_mian_pin_ming, ji_xing, xing_hao,
                     modified_at, now_str, file_path))
                updated += 1
            else:
                cur.execute(
                    'INSERT INTO drawing_index '
                    '(file_path, pin_hao, tu_mian_pin_ming, ji_xing, xing_hao, modified_at, indexed_at) '
                    'VALUES (?,?,?,?,?,?,?)',
                    (file_path, pin_hao, tu_mian_pin_ming, ji_xing, xing_hao,
                     modified_at, now_str))
                inserted += 1

        except Exception as e:
            errors += 1
            if errors <= 5:
                print(f'  [ERR] {fname}: {e}')

        if idx % BATCH == 0:
            conn.commit()
            pct = idx * 100 // total
            print(f'  {pct:3d}%  {idx:,}/{total:,}  新增:{inserted:,} 更新:{updated:,} 跳過:{skipped:,}')

    conn.commit()
    conn.close()

    print(f'\n  完成！新增:{inserted:,}  更新:{updated:,}  跳過:{skipped:,}  錯誤:{errors:,}')
    print(f'  DB -> {DB_PATH}')

    if deploy:
        _deploy(DB_PATH)

    return inserted + updated


def _deploy(src):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir   = os.path.join(script_dir, 'dist_embed', '\u88fd\u4ee4\u67e5\u8a62')
    if not os.path.isdir(dist_dir):
        print(f'  [WARN] dist_embed 目錄不存在，略過部署：{dist_dir}')
        return
    dst = os.path.join(dist_dir, 'pdm_search.db')
    shutil.copy2(src, dst)
    size_mb = os.path.getsize(src) / 1024 / 1024
    print(f'  部署 pdm_search.db ({size_mb:.1f} MB) -> {dist_dir}')


# ══════════════════════════════════════════════════════════
#  入口
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    args = sys.argv[1:]
    update_only = '--update' in args
    deploy      = '--deploy' in args

    start = datetime.datetime.now()
    print('=' * 55)
    print('  PDM 圖面索引重建工具 v1.2')
    print(f'  模式: {"增量更新" if update_only else "完整重建"}{"  +部署" if deploy else ""}')
    print(f'  時間: {start:%Y-%m-%d %H:%M:%S}')
    print('=' * 55)

    try:
        count   = rebuild(update_only=update_only, deploy=deploy)
        elapsed = int((datetime.datetime.now() - start).total_seconds())
        m, s    = divmod(elapsed, 60)
        print(f'\n  耗時: {m}分{s}秒')
        print('=' * 55)
    except KeyboardInterrupt:
        print('\n  使用者中斷')
        sys.exit(1)
    except Exception as e:
        print(f'\n  [ERROR] {e}')
        import traceback; traceback.print_exc()
        sys.exit(1)
