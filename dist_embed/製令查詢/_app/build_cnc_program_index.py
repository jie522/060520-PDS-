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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(_APP_DIR, 'cnc_program_index.db')

# 只索引這些副檔名：.txt 為 CNC 程式本體，圖片為加工/夾治具參考照片
INDEXABLE_EXT = {'.txt', '.jpg', '.jpeg', '.png', '.bmp'}
SKIP_FOLDER_NAMES = {'【空白範本】', '已刪除'}

SCHEMA = """
DROP TABLE IF EXISTS cnc_program_index;
CREATE TABLE cnc_program_index (
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
CREATE INDEX idx_cnc_model ON cnc_program_index(model);
CREATE INDEX idx_cnc_top   ON cnc_program_index(top_folder);
"""

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


def init_db(conn):
    conn.cursor().executescript(SCHEMA)
    conn.commit()


def _deploy(src):
    dist_app = os.path.join(_APP_DIR, 'dist_embed', '製令查詢', '_app')
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
        return 0

    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    cur = conn.cursor()

    inserted = 0
    errors = 0

    for dirpath, dirnames, filenames in os.walk(root):
        # 就地過濾不要遞迴進去的子資料夾（空白範本、已刪除）
        dirnames[:] = [d for d in dirnames if d not in SKIP_FOLDER_NAMES]

        for fn in filenames:
            if fn.startswith('~$'):
                continue
            ext = os.path.splitext(fn)[1].lower()
            if ext not in INDEXABLE_EXT:
                continue
            full_path = os.path.join(dirpath, fn)
            relpath = os.path.relpath(full_path, root).replace(os.sep, '/')
            try:
                st = os.stat(full_path)
            except OSError as e:
                errors += 1
                print(f'  [ERR] {relpath}: {e}')
                continue

            top_folder, model, machine = parse_path(relpath)
            remark = parse_remark(fn)
            mtime = datetime.datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

            try:
                cur.execute(
                    'INSERT OR REPLACE INTO cnc_program_index '
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
    conn.close()

    print(f'\n完成！索引 {inserted} 筆，錯誤 {errors} 筆')
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

    rebuild(deploy=deploy)

    elapsed = (datetime.datetime.now() - start).total_seconds()
    print(f'\n總耗時: {elapsed:.1f} 秒')
