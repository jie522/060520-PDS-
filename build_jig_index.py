#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
治檢具清單索引重建工具

掃描 PDM Vault『06-生技課\\01-模檢治具(2026後)』底下的子資料夾（如 PT2603003），
下載各子資料夾內的「機器模檢治具申請單」xlsm 檔案，
讀取其 Excel Defined Names（YC_機型、YC_品名、YC_經辦、YC_提出人員）
及檔案的 PDM 工作流程狀態（CurrentState），
篩選出提出人員為指定名單者，寫入 pdm_search.db 的 jig_index 表。

用法：
  python build_jig_index.py            # 重建治檢具索引
  python build_jig_index.py --deploy   # 重建後同步至 dist_embed
"""

import os
import sys
import sqlite3
import datetime

import openpyxl

import config
from build_pdm_index import connect_vault, DB_DIR, DB_PATH, _deploy

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


def get_defined_value(wb, name):
    """讀取 Excel Defined Name 對應的儲存格值，不存在則回傳 None"""
    if name not in wb.defined_names:
        return None
    dn = wb.defined_names[name]
    for sheet_name, coord in dn.destinations:
        try:
            return wb[sheet_name][coord].value
        except Exception:
            return None
    return None


def read_application_form(folder, f):
    """嘗試將 f 視為「申請單」xlsm 並讀取所需欄位，失敗或非申請單回傳 None"""
    if not f.Name.lower().endswith('.xlsm'):
        return None

    try:
        f.GetFileCopy(0, '', folder.ID, 1)
        local_path = f.GetLocalPath(folder.ID)
        if not os.path.exists(local_path):
            return None

        wb = openpyxl.load_workbook(local_path, data_only=True, keep_vba=False)

        submitter = get_defined_value(wb, 'YC_提出人員')
        if not submitter or not str(submitter).strip():
            return None

        product_model = get_defined_value(wb, 'YC_機型')
        item_name      = get_defined_value(wb, 'YC_品名')
        handler        = get_defined_value(wb, 'YC_經辦')

        status = None
        try:
            status = f.CurrentState.Name
        except Exception:
            pass

        return {
            'product_model': (str(product_model).strip() if product_model else ''),
            'item_name':     (str(item_name).strip() if item_name else ''),
            'handler':       (str(handler).strip() if handler else ''),
            'submitter':     str(submitter).strip(),
            'status':        (str(status).strip() if status else ''),
        }
    except Exception:
        return None


def init_db(conn):
    cur = conn.cursor()
    cur.executescript(SCHEMA)
    cur.execute('DELETE FROM jig_index')
    conn.commit()


def rebuild(deploy=False):
    os.makedirs(DB_DIR, exist_ok=True)

    print('連線 PDM Vault ...')
    vault = connect_vault()

    print(f'掃描資料夾：{config.JIG_VAULT_PATH}')
    subfolders = get_subfolders(vault, config.JIG_VAULT_PATH)
    total = len(subfolders)
    print(f'  找到 {total} 個子資料夾')

    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    cur = conn.cursor()

    matched = skipped = errors = 0

    for idx, folder in enumerate(subfolders, 1):
        try:
            record = None
            pos = folder.GetFirstFilePosition()
            while not pos.IsNull:
                f = folder.GetNextFile(pos)
                record = read_application_form(folder, f)
                if record:
                    break

            if not record:
                skipped += 1
                continue

            if record['submitter'] not in config.JIG_SUBMITTERS:
                skipped += 1
                continue

            folder_path = os.path.join(config.JIG_VAULT_PATH, folder.Name)
            cur.execute(
                'INSERT OR REPLACE INTO jig_index '
                '(folder_name, folder_path, product_model, item_name, handler, submitter, status) '
                'VALUES (?,?,?,?,?,?,?)',
                (folder.Name, folder_path, record['product_model'], record['item_name'],
                 record['handler'], record['submitter'], record['status']))
            matched += 1
            print(f'  [{idx}/{total}] {folder.Name}: 機型={record["product_model"]} '
                  f'品名={record["item_name"]} 經辦={record["handler"]} '
                  f'提出人員={record["submitter"]} 狀態={record["status"]}')

        except Exception as e:
            errors += 1
            print(f'  [{idx}/{total}] {folder.Name}: ERR={e}')

        if idx % 10 == 0:
            conn.commit()
            print(f'  進度 {idx}/{total}  符合:{matched}  跳過:{skipped}  錯誤:{errors}')

    conn.commit()
    conn.close()

    print(f'\n完成！符合條件並寫入：{matched} 筆，跳過：{skipped} 筆，錯誤：{errors} 筆')
    print(f'DB -> {DB_PATH}')

    if deploy:
        _deploy(DB_PATH)

    return matched


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
