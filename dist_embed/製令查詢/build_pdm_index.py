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
  - 發圖日期        : 嘗試 PDM GetVar，失敗改用本機快取檔案日期
  - 圖面品名        : 嘗試 PDM GetVar（若未本機快取通常讀不到，留空）
"""

import os
import sys
import re
import shutil
import sqlite3
import datetime
import win32com.client

# ══════════════════════════════════════════════════════════
#  設定
# ══════════════════════════════════════════════════════════

VAULT_NAME = 'MAXCLAW'
TARGET_EXT = '.SLDDRW'

DB_DIR  = os.path.join(os.environ.get('LOCALAPPDATA', r'C:\Users\Public'), 'PDMSearch')
DB_PATH = os.path.join(DB_DIR, 'pdm_search.db')

# 機型資料夾名稱正規表達式（如 FTD430B, FTB245, MSB300）
MACHINE_RE = re.compile(r'^[A-Za-z]{2,6}\d{2,6}[A-Za-z]{0,3}$')

# ══════════════════════════════════════════════════════════
#  PDM COM 工具函式
# ══════════════════════════════════════════════════════════

def connect_vault():
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
    # fallback：最後一層資料夾
    return parts[-1] if parts else ''


def try_get_var(ev, file_obj, var_name, version, var_ids=None):
    """嘗試讀取 PDM 自訂變數
    方法1: ev.GetVarFromDb（直讀資料庫，不需本機快取）
    方法2: ev.GetVar via enumerator（本機快取備援）
    """
    # 方法 1：ev.GetVarFromDb — 直讀 PDM SQL 資料庫（需已探索到 var_id）
    if var_ids:
        var_id = var_ids.get(var_name)
        if var_id:
            try:
                r = ev.GetVarFromDb(var_id, version)
                if r is not None:
                    vals = r if isinstance(r, tuple) else (r,)
                    for val in vals:
                        s = str(val).strip() if val is not None else ''
                        if s and s not in ('None', 'False', '0'):
                            return s
            except Exception:
                pass

    # 方法 2：GetVar（只對本機已快取的檔案有效）
    # 嘗試多種 config 名稱，涵蓋不同 PDM 設定
    configs = ['', '@', 'Default']
    for v in ([version] if version and version > 0 else []) + [0]:
        for cfg in configs:
            try:
                r = ev.GetVar(var_name, v, cfg)
                if isinstance(r, tuple) and len(r) >= 2 and r[1]:
                    s = str(r[1]).strip()
                    if s:
                        return s
            except Exception:
                pass
    return ''


def file_mdate(local_path):
    """取得本機檔案修改日期（格式 YYYY/MM/DD），失敗回傳空字串"""
    try:
        mt = os.path.getmtime(local_path)
        return datetime.datetime.fromtimestamp(mt).strftime('%Y/%m/%d')
    except Exception:
        return ''



# ══════════════════════════════════════════════════════════
#  Vault 遞迴掃描
# ══════════════════════════════════════════════════════════

def scan_folder(folder, results, counters):
    """遞迴掃描 PDM 資料夾，收集 SLDDRW 檔案
    results 存 (folder_path_str, file_name_str, file_obj)
    路徑在掃描階段就轉成字串，避免後續 COM 狀態污染導致 LocalPath 失效
    """
    ext = TARGET_EXT.upper()
    # 掃描階段立即讀取 LocalPath，存成 Python 字串
    folder_path = str(folder.LocalPath)

    # --- 列舉此資料夾的檔案 ---
    pos = folder.GetFirstFilePosition()
    while not pos.IsNull:
        f = folder.GetNextFile(pos)
        if f.Name.upper().endswith(ext):
            results.append((folder_path, str(f.Name), f))
            counters[0] += 1
            if counters[0] % 200 == 0:
                print(f'  掃描中... {counters[0]:,} 個檔案')

    # --- 遞迴子資料夾 ---
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
    """取得品名或日期為空的既有路徑（增量模式下需補填）"""
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

    # 步驟 2：遞迴掃描
    print('[2/4] 掃描 Vault 中所有 SLDDRW 檔案 ...')
    results  = []
    counters = [0]
    scan_folder(vault.RootFolder, results, counters)
    total = len(results)
    print(f'  找到 {total:,} 個 {TARGET_EXT} 檔案')

    if total == 0:
        print('  未找到任何檔案，結束。')
        return 0

    # 步驟 3：備份 + 初始化 DB
    print('[3/4] 初始化資料庫 ...')
    if os.path.exists(DB_PATH) and not update_only:
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        bak = DB_PATH.replace('.db', f'_backup_{ts}.db')
        shutil.copy2(DB_PATH, bak)
        print(f'  備份 -> {os.path.basename(bak)}')

    conn = sqlite3.connect(DB_PATH)
    init_db(conn, full_rebuild=not update_only)
    cur  = conn.cursor()

    existing_paths = get_existing_paths(conn) if update_only else set()
    empty_paths    = get_empty_paths(conn)   if update_only else set()

    # ── 掃描完成後，透過快取檔案反向探索變數 ID ──
    _var_ids = {}
    TARGET_VARS = ['品名', '發圖日期']
    _log_path = os.path.join(DB_DIR, 'pdm_debug.txt')

    with open(_log_path, 'w', encoding='utf-8') as _lf:
        # 0. 診斷：印前 5 筆路徑及 os.path.exists 結果
        _lf.write('=== 路徑診斷 ===\n')
        for folder_path, fname, file_obj in results[:5]:
            fp = os.path.join(folder_path, fname)
            _lf.write(f'  exists={os.path.exists(fp)}  repr={repr(fp)}\n')

        # 0b. 直接測前 20 筆的 ev.GetVar（各種 config）及 ev.GetVarFromDb（ID 1~10）
        _lf.write('\n=== ev.GetVar / GetVarFromDb 測試（前20筆）===\n')
        for folder_path, fname, file_obj in results[:20]:
            ev  = file_obj.GetEnumeratorVariable()
            ver = file_obj.CurrentVersion
            row = f'  {fname}: ver={ver}'
            for cfg in ['', '@', 'Default']:
                try:
                    r = ev.GetVar('品名', ver, cfg)
                    row += f' GetVar(cfg={repr(cfg)})={repr(r)}'
                except Exception as e:
                    row += f' GetVar(cfg={repr(cfg)})=ERR:{e}'
            # 也測 GetVarFromDb id=1~10 看有無資料
            ids_found = {}
            for vid in range(1, 11):
                try:
                    r = ev.GetVarFromDb(vid, ver)
                    if r is not None and r not in ((False, None), (False, '')):
                        vals = r if isinstance(r, tuple) else (r,)
                        for val in vals:
                            if val is not None:
                                s = str(val).strip()
                                if s and s not in ('None', 'False'):
                                    ids_found[vid] = s
                except Exception:
                    pass
            if ids_found:
                row += f' DbIds={ids_found}'
            _lf.write(row + '\n')

        # 0c. 第一個檔案：掃描 GetVarFromDb ID 1~500，列出所有有值的 ID（不需比對預期值）
        _lf.write('\n=== 第一筆檔案 GetVarFromDb 全掃（ID 1~500）===\n')
        if results:
            fp0, fn0, fo0 = results[0]
            ev0  = fo0.GetEnumeratorVariable()
            ver0 = fo0.CurrentVersion
            _lf.write(f'  檔案: {fn0}  ver={ver0}\n')
            for vid in range(1, 501):
                try:
                    r = ev0.GetVarFromDb(vid, ver0)
                    if r is not None and r not in ((False, None), (False, '')):
                        vals = r if isinstance(r, tuple) else (r,)
                        for val in vals:
                            if val is not None:
                                s = str(val).strip()
                                if s and s not in ('None', 'False'):
                                    _lf.write(f'  id={vid}: {repr(s)}\n')
                except Exception:
                    pass

        # 步驟 A：找第一個 ev.GetVar 有品名值的檔案（不用 os.path.exists，嘗試多種 config）
        known_values = {}
        known_file_obj = {}
        for folder_path, fname, file_obj in results:
            ev  = file_obj.GetEnumeratorVariable()
            ver = file_obj.CurrentVersion
            for vn in TARGET_VARS:
                if vn not in known_values:
                    for cfg in ['', '@', 'Default']:
                        try:
                            r = ev.GetVar(vn, ver, cfg)
                            if isinstance(r, tuple) and len(r) >= 2 and r[1]:
                                val = str(r[1]).strip()
                                if val:
                                    known_values[vn] = (val, ev, ver)
                                    known_file_obj[vn] = (file_obj, fname)
                                    break
                        except Exception:
                            pass
            if len(known_values) == len(TARGET_VARS):
                break

        _lf.write(f'\n=== 步驟A 已知值 ===\n')
        _lf.write(f'已知值: { {k: v[0] for k, v in known_values.items()} }\n')
        for k, (fo, fn) in known_file_obj.items():
            _lf.write(f'  {k} 來源檔案: {fn}\n')

        # 步驟 B：brute-force ev.GetVarFromDb(id, ver) 找出 id→name 對應
        for var_name, (expected_val, ev, ver) in known_values.items():
            _lf.write(f'\n搜尋 "{var_name}" (已知值="{expected_val}") ID 1~500:\n')
            for vid in range(1, 501):
                try:
                    r = ev.GetVarFromDb(vid, ver)
                    if r is not None and r not in ((False, None), (False, '')):
                        vals = r if isinstance(r, tuple) else (r,)
                        for val in vals:
                            if val is not None:
                                s = str(val).strip()
                                if s and s not in ('None', 'False'):
                                    _lf.write(f'  id={vid}: {repr(s)}\n')
                                    if s.strip() == expected_val.strip():
                                        _var_ids[var_name] = vid
                                        _lf.write(f'  *** MATCH: {var_name} = id {vid} ***\n')
                except Exception:
                    pass
            if var_name not in _var_ids:
                _lf.write(f'  (未找到 {var_name} 對應 ID)\n')

        _lf.write(f'\n最終 var_ids: {_var_ids}\n')

    print(f'  [VAR-ID] 探索結果：{_var_ids}，詳見 {_log_path}')
    print(f'  取得 {len(_var_ids)} 個 PDM 變數 ID')

    # 步驟 4：寫入索引
    print('[4/4] 寫入索引 ...')
    inserted = updated = skipped = errors = 0
    BATCH = 300

    for idx, (folder_path, fname, file_obj) in enumerate(results, 1):
        file_path = os.path.join(folder_path, fname)

        # 增量模式：已存在且欄位完整則跳過；欄位有空值仍繼續嘗試補填
        if update_only and file_path in existing_paths and file_path not in empty_paths:
            skipped += 1
            continue

        try:
            pin_hao  = os.path.splitext(fname)[0]
            ji_xing  = derive_ji_xing(folder_path)
            xing_hao = ji_xing

            # 嘗試讀取 PDM 變數
            ev  = file_obj.GetEnumeratorVariable()
            ver = file_obj.CurrentVersion
            tu_mian_pin_ming = try_get_var(ev, file_obj, '品名',   ver, _var_ids)
            modified_at      = try_get_var(ev, file_obj, '發圖日期', ver, _var_ids)

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

    # 部署
    if deploy:
        _deploy(DB_PATH)

    return inserted + updated


def _deploy(src):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir   = os.path.join(script_dir, 'dist_embed', '\u88fd\u4ee4\u67e5\u8a62')  # 製令查詢
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
    print('  PDM 圖面索引重建工具 v1.1')
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
