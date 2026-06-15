#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
治檢具清單索引重建工具

掃描 PDM Vault『06-生技課\\01-模檢治具(2026後)』底下的子資料夾（如 PT2606004），
讀取每個子資料夾的資料卡欄位（產品機型、品名、圖面機型、經辦、提出人員、YC_狀態、取出者），
篩選出提出人員為指定名單者，寫入 pdm_search.db 的 jig_index 表。

用法：
  python build_jig_index.py            # 重建治檢具索引
  python build_jig_index.py --deploy   # 重建後同步至 dist_embed
"""

import os
import sys
import sqlite3
import datetime
import win32com.client

import config
from build_pdm_index import connect_vault, DB_DIR, DB_PATH, _deploy

TARGET_VARS = ['產品機型', '品名', '圖面機型', '經辦', '提出人員', 'YC_狀態', '取出者']

SCHEMA = """
CREATE TABLE IF NOT EXISTS jig_index (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_name   TEXT NOT NULL,
    folder_path   TEXT NOT NULL UNIQUE,
    product_model TEXT,
    item_name     TEXT,
    drawing_model TEXT,
    handler       TEXT,
    submitter     TEXT,
    status        TEXT,
    taken_by      TEXT,
    indexed_at    TEXT DEFAULT (datetime('now'))
);
"""


def get_subfolders(vault, root_path):
    """取得指定路徑下所有直接子資料夾物件"""
    root = vault.GetFolderFromPath(root_path)
    subs = []
    pos = root.GetFirstSubFolderPosition()
    while not pos.IsNull:
        subs.append(root.GetNextSubFolder(pos))
    return subs


def read_folder_vars(folder, debug=None):
    """讀取單一資料夾的資料卡變數，回傳 {變數名: 值}。debug 為 list 時會附加診斷字串"""
    def _dbg(msg):
        if debug is not None:
            debug.append(msg)

    values = {}
    try:
        ev_early = folder.GetEnumeratorVariable()
    except Exception as e:
        _dbg(f'GetEnumeratorVariable 失敗: {e}')
        return values

    try:
        ev10 = win32com.client.CastTo(ev_early, 'IEdmEnumeratorVariable10')
    except Exception as e:
        _dbg(f'CastTo IEdmEnumeratorVariable10 失敗: {e}')
        return values

    folder_id = folder.ID

    # 策略 A：GetVarFromDb
    for vn in TARGET_VARS:
        try:
            r = ev10.GetVarFromDb(vn, '@')
            _dbg(f'A:{vn}={r!r}')
            if r and r[0] is True and r[1]:
                values[vn] = str(r[1]).strip()
        except Exception as e:
            _dbg(f'A:{vn} ERR={e}')

    if values:
        return values

    # 策略 B：StoreValuesFromDatabase + GetVar
    try:
        ev10.StoreValuesFromDatabase(folder_id, False, 0)
        for vn in TARGET_VARS:
            try:
                r = ev10.GetVar(vn, '@')
                _dbg(f'B:{vn}={r!r}')
                if r and r[0] is True and r[1]:
                    values[vn] = str(r[1]).strip()
            except Exception as e:
                _dbg(f'B:{vn} ERR={e}')
    except Exception as e:
        _dbg(f'B:StoreValuesFromDatabase ERR={e}')

    if values:
        return values

    # 策略 C：GetVar2 with folder ID
    for vn in TARGET_VARS:
        try:
            r = ev10.GetVar2(vn, '@', folder_id)
            _dbg(f'C:{vn}={r!r}')
            if r and r[0] is True and r[1]:
                values[vn] = str(r[1]).strip()
        except Exception as e:
            _dbg(f'C:{vn} ERR={e}')

    if values:
        return values

    # 策略 D：GetVarAsText with folder ID
    for vn in TARGET_VARS:
        try:
            r = ev10.GetVarAsText(vn, '@', folder_id)
            _dbg(f'D:{vn}={r!r}')
            if r and r[0] is True and r[1]:
                values[vn] = str(r[1]).strip()
        except Exception as e:
            _dbg(f'D:{vn} ERR={e}')

    return values


def rebuild(deploy=False):
    os.makedirs(DB_DIR, exist_ok=True)

    print('連線 PDM Vault ...')
    vault = connect_vault()

    print(f'掃描資料夾：{config.JIG_VAULT_PATH}')
    subfolders = get_subfolders(vault, config.JIG_VAULT_PATH)
    print(f'  找到 {len(subfolders)} 個子資料夾')

    # ── 診斷：印出前 3 個資料夾的讀取結果 ──
    print('\n--- 診斷（前 3 個資料夾）---')
    for sub in subfolders[:3]:
        debug = []
        vals = read_folder_vars(sub, debug=debug)
        print(f'  {sub.Name}:')
        for line in debug:
            print(f'    {line}')
        print(f'    => {vals}')

    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA)
    conn.execute('DELETE FROM jig_index')

    print('\n--- 開始讀取 ---')
    matched = 0
    for sub in subfolders:
        name = str(sub.Name)
        path = str(sub.LocalPath)
        vals = read_folder_vars(sub)
        submitter = vals.get('提出人員', '')

        if submitter not in config.JIG_SUBMITTERS:
            continue

        conn.execute(
            'INSERT INTO jig_index '
            '(folder_name, folder_path, product_model, item_name, drawing_model, '
            ' handler, submitter, status, taken_by) VALUES (?,?,?,?,?,?,?,?,?)',
            (
                name, path,
                vals.get('產品機型', ''),
                vals.get('品名', ''),
                vals.get('圖面機型', ''),
                vals.get('經辦', ''),
                submitter,
                vals.get('YC_狀態', ''),
                vals.get('取出者', ''),
            )
        )
        matched += 1
        print(f'  [{matched}] {name}  提出人員={submitter}  品名={vals.get("品名","")}')

    conn.commit()
    conn.close()

    print(f'\n完成！共 {len(subfolders)} 個資料夾，符合條件 {matched} 筆')
    print(f'DB -> {DB_PATH}')

    if deploy:
        _deploy(DB_PATH)


if __name__ == '__main__':
    args = sys.argv[1:]
    deploy = '--deploy' in args

    start = datetime.datetime.now()
    print('=' * 55)
    print('  治檢具清單索引重建工具')
    print(f'  時間: {start:%Y-%m-%d %H:%M:%S}')
    print('=' * 55)

    rebuild(deploy=deploy)

    elapsed = (datetime.datetime.now() - start).total_seconds()
    print(f'\n總耗時: {elapsed:.1f} 秒')
