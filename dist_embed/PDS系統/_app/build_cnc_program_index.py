#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CNC 程式索引重建工具

掃描 config.CNC_PROGRAM_ROOT_PATH（網路資料夾）底下所有 CNC 程式檔案，
寫入 cnc_program_index.db（與 app.py 同目錄，隨 dist_embed 一起部署）。

資料夾慣例（不完全一致，索引時用寬鬆規則解析）：
  P.程式/
    FTB系列/[FTB210A夾管座]/250310.txt              ← 無機台子資料夾
    FTB系列/[FTB400A側孔]/【綜銑01(協鴻)】/240227.txt ← 有機台子資料夾
    第6台程式/O0001.txt                              ← 無系列/品號分類
  跳過：【空白範本】（只是空白機台資料夾範本）、「已刪除」資料夾（軟刪除）

用法：
  python build_cnc_program_index.py            # 重建索引
  python build_cnc_program_index.py --deploy   # 重建後同步至 dist_embed
"""

import os
import re
import sys
import sqlite3
import shutil
import datetime
import concurrent.futures

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(_APP_DIR, 'cnc_program_index.db')

# 只索引這些副檔名：.txt 為 CNC 程式本體，圖片為加工/夾治具參考照片
INDEXABLE_EXT = {'.txt', '.jpg', '.jpeg', '.png', '.bmp'}
SKIP_FOLDER_NAMES = {'【空白範本】', '已刪除'}

# 掃描結果先寫進「建置中」的暫存表，掃完確認筆數合理才原子交換成正式表，
# 避免網路資料夾暫時無法列出內容時，把正式索引清空覆蓋成 0 筆（2026-07 真實故障）。
BUILD_TABLE = 'cnc_program_index_build'
LIVE_TABLE = 'cnc_program_index'

SCHEMA = """
DROP TABLE IF EXISTS {table};
CREATE TABLE {table} (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    top_folder  TEXT,
    model       TEXT,
    machine     TEXT,
    filename    TEXT NOT NULL,
    remark      TEXT,
    relpath     TEXT NOT NULL UNIQUE,
    ext         TEXT,
    size        INTEGER,
    mtime       TEXT,
    indexed_at  TEXT DEFAULT (datetime('now'))
);
"""

# 掃描結果比舊索引少於這個比例就視為異常（網路異常/半途中斷等），中止並保留舊資料
MIN_KEEP_RATIO = 0.5

_MODEL_RE   = re.compile(r'^\[(.+)\]$')
_MACHINE_RE = re.compile(r'^【(.+)】$')
_REMARK_RE  = re.compile(r'^\d{6,8}[_\-](.+)$')


def parse_path(relpath):
    """從相對路徑解析 top_folder / model / machine"""
    segments = relpath.split('/')[:-1]  # 不含檔名
    top_folder = segments[0] if segments else ''
    model = ''
    machine = ''
    for seg in segments:
        m = _MODEL_RE.match(seg)
        if m and not model:
            model = m.group(1)
    for seg in segments[1:]:
        m = _MACHINE_RE.match(seg)
        if m and not machine:
            machine = m.group(1)
    return top_folder, model, machine


def parse_remark(filename):
    stem = os.path.splitext(filename)[0]
    m = _REMARK_RE.match(stem)
    if not m:
        return ''
    rest = m.group(1)
    # 去掉結尾的 [機台代號] 之類標記，只留文字說明
    rest = re.sub(r'_?\[[^\]]*\]$', '', rest).strip('_- ')
    return rest


def _scandir_safe(path):
    with os.scandir(path) as it:
        return list(it)


def _walk_entries_parallel(root, max_workers=16):
    """平行掃描網路資料夾樹，回傳所有檔案 DirEntry（用 DirEntry.stat() 重用掃描時
    就拿到的大小/時間，不再對每個檔案多打一次 os.stat()）。

    實測發現真正的瓶頸不是每檔案的 stat，而是「列出一個資料夾」這個動作本身在網路
    磁碟機（SMB）上的來回延遲——這批資料夾結構深（機台程式/底下大量時間戳子資料夾），
    序列掃描等於把這個延遲乘以資料夾數量。改用執行緒併發列出多個資料夾（I/O bound，
    系統呼叫期間會釋放 GIL），把原本疊加的延遲變成平行，是縮短總時間的關鍵。"""
    file_entries = []
    scanned_dirs = 0
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_scandir_safe, root): root}
        while futures:
            done, _ = concurrent.futures.wait(futures, return_when=concurrent.futures.FIRST_COMPLETED)
            for fut in done:
                dirpath = futures.pop(fut)
                try:
                    entries = fut.result()
                except OSError as e:
                    print(f'  [ERR] 無法列出 {dirpath}: {e}')
                    continue
                scanned_dirs += 1
                if scanned_dirs % 50 == 0:
                    try:
                        print(f'  掃描中... 已列出 {scanned_dirs} 個資料夾，找到 {len(file_entries)} 個檔案 ...')
                    except UnicodeEncodeError:
                        pass
                for entry in entries:
                    try:
                        is_dir = entry.is_dir(follow_symlinks=False)
                    except OSError:
                        continue
                    if is_dir:
                        if entry.name in SKIP_FOLDER_NAMES:
                            continue
                        futures[pool.submit(_scandir_safe, entry.path)] = entry.path
                    else:
                        file_entries.append(entry)
    return file_entries


def init_db(conn, table):
    conn.cursor().executescript(SCHEMA.format(table=table))
    conn.commit()


def _table_exists(conn, table):
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
    ).fetchone()
    return row is not None


def _row_count(conn, table):
    if not _table_exists(conn, table):
        return 0
    return conn.execute(f'SELECT COUNT(*) FROM {table}').fetchone()[0]


def _swap_in_build_table(conn):
    """把建置完成的暫存表原子性地換成正式表（先建好索引，再交換，交換全程在同一 transaction）"""
    cur = conn.cursor()
    cur.execute(f'DROP TABLE IF EXISTS {LIVE_TABLE}')
    cur.execute(f'ALTER TABLE {BUILD_TABLE} RENAME TO {LIVE_TABLE}')
    cur.execute(f'CREATE INDEX IF NOT EXISTS idx_cnc_model ON {LIVE_TABLE}(model)')
    cur.execute(f'CREATE INDEX IF NOT EXISTS idx_cnc_top   ON {LIVE_TABLE}(top_folder)')
    conn.commit()


def _deploy(src):
    dist_app = os.path.join(_APP_DIR, 'dist_embed', 'PDS系統', '_app')
    if not os.path.isdir(dist_app):
        print(f'  [WARN] dist_embed 目錄不存在，略過部署：{dist_app}')
        return
    dst = os.path.join(dist_app, os.path.basename(src))
    shutil.copy2(src, dst)
    print(f'  部署 {os.path.basename(src)} -> {dist_app}')


def rebuild(deploy=False):
    root = config.CNC_PROGRAM_ROOT_PATH
    print(f'掃描資料夾：{root}')
    if not os.path.isdir(root):
        print(f'  [ERROR] 資料夾不存在或無法存取：{root}')
        return -1

    conn = sqlite3.connect(DB_PATH)
    prev_count = _row_count(conn, LIVE_TABLE)
    init_db(conn, BUILD_TABLE)  # 寫進暫存表，正式表在確認掃描結果合理前完全不動
    cur = conn.cursor()

    inserted = 0
    errors = 0

    file_entries = _walk_entries_parallel(root)
    print(f'  掃描完成，共找到 {len(file_entries)} 個檔案，開始寫入索引 ...')

    for entry in file_entries:
        fn = entry.name
        if fn.startswith('~$'):
            continue
        ext = os.path.splitext(fn)[1].lower()
        if ext not in INDEXABLE_EXT:
            continue
        full_path = entry.path
        relpath = os.path.relpath(full_path, root).replace(os.sep, '/')
        try:
            st = entry.stat()
        except OSError as e:
            errors += 1
            print(f'  [ERR] {relpath}: {e}')
            continue

        top_folder, model, machine = parse_path(relpath)
        remark = parse_remark(fn)
        mtime = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

        try:
            cur.execute(
                f'INSERT OR REPLACE INTO {BUILD_TABLE} '
                '(top_folder, model, machine, filename, remark, relpath, ext, size, mtime) '
                'VALUES (?,?,?,?,?,?,?,?,?)',
                (top_folder, model, machine, fn, remark, relpath, ext, st.st_size, mtime)
            )
            inserted += 1
        except Exception as e:
            errors += 1
            print(f'  [ERR] {relpath}: {e}')

        if inserted % 200 == 0:
            conn.commit()
            try:
                print(f'  進度：已索引 {inserted} 筆 ...')
            except UnicodeEncodeError:
                pass

    conn.commit()

    # 安全檢查：掃描結果比舊索引少太多，視為異常（網路中斷/半途失敗等），
    # 中止並保留舊的正式表，不做交換也不部署，避免用 0 筆或殘缺資料覆蓋掉還堪用的舊索引。
    if prev_count > 0 and inserted < prev_count * MIN_KEEP_RATIO:
        conn.close()
        print(f'\n[ABORT] 掃描結果異常：本次僅 {inserted} 筆，舊索引有 {prev_count} 筆'
              f'（低於 {MIN_KEEP_RATIO:.0%} 門檻），懷疑網路資料夾暫時無法完整列出。')
        print('  已保留舊的正式索引，不覆蓋、不部署。請確認網路資料夾可正常存取後再重試。')
        return -1

    _swap_in_build_table(conn)
    conn.close()

    print(f'\n完成！索引 {inserted} 筆，錯誤 {errors} 筆（舊索引 {prev_count} 筆）')
    print(f'DB -> {DB_PATH}')

    if deploy:
        _deploy(DB_PATH)

    return inserted


if __name__ == '__main__':
    args = sys.argv[1:]
    deploy = '--deploy' in args

    start = datetime.datetime.now()
    print('=' * 55)
    print('  CNC 程式索引重建工具')
    print(f'  時間: {start:%Y-%m-%d %H:%M:%S}')
    print('=' * 55)

    result = rebuild(deploy=deploy)

    elapsed = (datetime.datetime.now() - start).total_seconds()
    print(f'\n總耗時: {elapsed:.1f} 秒')

    if result < 0:
        sys.exit(1)  # 讓呼叫端（app.py 的 /api/cnc_program/rebuild）能判斷這次是失敗，不是「成功但 0 筆」
