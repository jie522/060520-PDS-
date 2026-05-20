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


def try_get_var(ev, file_obj, var_name, version):
    """嘗試讀取 PDM 自訂變數"""
    try:
        r = ev.GetVar(var_name, version, '')
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


def pdm_version_date(file_obj):
    """從 PDM IEdmVersion5 取版次建立日期（不依賴本機快取）"""
    try:
        ver_num = file_obj.CurrentVersion
        if ver_num and ver_num > 0:
            ver_obj = file_obj.GetVersion(ver_num, '')
            if ver_obj:
                d = getattr(ver_obj, 'DateCreated', None)
                if d:
                    if hasattr(d, 'strftime'):          # Python datetime
                        return d.strftime('%Y/%m/%d')
                    else:                                # COM date float
                        base = datetime.datetime(1899, 12, 30)
                        dt = base + datetime.timedelta(days=float(d))
                        return dt.strftime('%Y/%m/%d')
    except Exception:
        pass
    return ''


# ══════════════════════════════════════════════════════════
#  Vault 遞迴掃描
# ══════════════════════════════════════════════════════════

def scan_folder(folder, results, counters):
    """遞迴掃描 PDM 資料夾，收集 SLDDRW 檔案"""
    ext = TARGET_EXT.upper()

    # --- 列舉此資料夾的檔案 ---
    pos = folder.GetFirstFilePosition()
    while not pos.IsNull:
        f = folder.GetNextFile(pos)
        if f.Name.upper().endswith(ext):
            results.append((folder, f))
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

    # 步驟 4：寫入索引
    print('[4/4] 寫入索引 ...')
    inserted = updated = skipped = errors = 0
    BATCH = 300

    for idx, (folder, file_obj) in enumerate(results, 1):
        folder_path = folder.LocalPath
        fname       = file_obj.Name
        file_path   = os.path.join(folder_path, fname)

        # 增量模式：已存在且欄位完整則跳過；欄位有空值仍繼續嘗試補填
        if update_only and file_path in existing_paths and file_path not in empty_paths:
            skipped += 1
            continue

        try:
            pin_hao  = os.path.splitext(fname)[0]
            ji_xing  = derive_ji_xing(folder_path)
            xing_hao = ji_xing

            # 嘗試讀取 PDM 變數
            ev = file_obj.GetEnumeratorVariable()
            ver = file_obj.CurrentVersion
            tu_mian_pin_ming = try_get_var(ev, file_obj, '品名', ver)
            modified_at      = try_get_var(ev, file_obj, '發圖日期', ver)

            # fallback 1：本機快取檔案修改日期
            if not modified_at:
                modified_at = file_mdate(file_path)
            # fallback 2：PDM 版次建立日期（不需要本機快取）
            if not modified_at:
                modified_at = pdm_version_date(file_obj)

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
