#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
治檢具清單索引重建工具

掃描 PDM Vault『06-生技課\\01-模檢治具(2026後)』底下的子資料夾（如 PT2603003），
下載各子資料夾內的「機器模檢治具申請單」xlsm 檔案，
讀取其 Excel Defined Names（YC_機型、YC_品名、YC_經辦、YC_提出人員、YC_申請日期、單位等）
及檔案的 PDM 工作流程狀態（CurrentState），
寫入 pdm_search.db 的 jig_index 表（不限提出人員，全部索引）。

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
DROP TABLE IF EXISTS jig_index;
CREATE TABLE jig_index (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_name   TEXT NOT NULL,
    folder_path   TEXT NOT NULL UNIQUE,
    product_model TEXT,
    item_name     TEXT,
    handler       TEXT,
    submitter     TEXT,
    status        TEXT,
    apply_date    TEXT,
    unit          TEXT,
    indexed_at    TEXT DEFAULT (datetime('now'))
);
"""

_first_folder_done = False


def get_subfolders(vault, root_path):
    """取得指定路徑下所有直接子資料夾物件"""
    root = vault.GetFolderFromPath(root_path)
    subs = []
    pos = root.GetFirstSubFolderPosition()
    while not pos.IsNull:
        subs.append(root.GetNextSubFolder(pos))
    return subs


def get_defined_value(wb, *names):
    """依序嘗試多個 Defined Name，回傳第一個非空的儲存格值"""
    for name in names:
        if name not in wb.defined_names:
            continue
        dn = wb.defined_names[name]
        for sheet_name, coord in dn.destinations:
            try:
                v = wb[sheet_name][coord].value
                if v is not None and str(v).strip():
                    return v
            except Exception:
                continue
    return None


def read_application_form(folder, f):
    """嘗試將 f 視為「申請單」xlsm 並讀取所需欄位，失敗或非申請單回傳 None"""
    global _first_folder_done

    if not f.Name.lower().endswith('.xlsm'):
        return None

    try:
        f.GetFileCopy(0, '', folder.ID, 1)
        local_path = f.GetLocalPath(folder.ID)
        if not os.path.exists(local_path):
            return None

        wb = openpyxl.load_workbook(local_path, data_only=True, keep_vba=False)

        # 第一個成功讀取的 xlsm 印出所有 Defined Names（診斷用）
        if not _first_folder_done:
            _first_folder_done = True
            print(f'  [診斷] {folder.Name}/{f.Name} 的所有 Defined Names:')
            for dn_name in sorted(wb.defined_names):
                dn = wb.defined_names[dn_name]
                for sn, coord in dn.destinations:
                    try:
                        val = wb[sn][coord].value
                        print(f'    {dn_name} = {val!r}')
                    except Exception:
                        print(f'    {dn_name} = (讀取失敗)')

        submitter = get_defined_value(wb, 'YC_提出人員')
        if not submitter or not str(submitter).strip():
            return None

        product_model = get_defined_value(wb, 'YC_機型')
        item_name      = get_defined_value(wb, 'YC_品名')
        handler        = get_defined_value(wb, 'YC_經辦')
        apply_date     = get_defined_value(wb, 'YC_日期', 'YC_申請日期', '申請日期')
        unit           = get_defined_value(wb, '單位', 'YC_保管單位', '保管單位')

        status = None
        try:
            status = f.CurrentState.Name
        except Exception:
            pass

        # 申請日期格式標準化（datetime/date → 字串 YYYY-MM-DD）
        if apply_date is not None:
            if hasattr(apply_date, 'strftime'):
                apply_date = apply_date.strftime('%Y-%m-%d')
            else:
                apply_date = str(apply_date).strip()

        return {
            'product_model': (str(product_model).strip() if product_model else ''),
            'item_name':     (str(item_name).strip() if item_name else ''),
            'handler':       (str(handler).strip() if handler else ''),
            'submitter':     str(submitter).strip(),
            'status':        (str(status).strip() if status else ''),
            'apply_date':    apply_date or '',
            'unit':          (str(unit).strip() if unit else ''),
        }
    except Exception as e:
        return None


def init_db(conn):
    cur = conn.cursor()
    cur.executescript(SCHEMA)
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

    inserted = skipped = errors = 0

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

            folder_path = os.path.join(config.JIG_VAULT_PATH, folder.Name)
            cur.execute(
                'INSERT OR REPLACE INTO jig_index '
                '(folder_name, folder_path, product_model, item_name, '
                ' handler, submitter, status, apply_date, unit) '
                'VALUES (?,?,?,?,?,?,?,?,?)',
                (folder.Name, folder_path, record['product_model'], record['item_name'],
                 record['handler'], record['submitter'], record['status'],
                 record['apply_date'], record['unit']))
            inserted += 1
            print(f'  [{idx}/{total}] {folder.Name}: '
                  f'機型={record["product_model"]} 品名={record["item_name"]} '
                  f'申請人={record["submitter"]} 日期={record["apply_date"]} '
                  f'保管={record["unit"]} 狀態={record["status"]}')

        except Exception as e:
            errors += 1
            print(f'  [{idx}/{total}] {folder.Name}: ERR={e}')

        if idx % 10 == 0:
            conn.commit()
            print(f'  進度 {idx}/{total}  寫入:{inserted}  跳過:{skipped}  錯誤:{errors}')

    conn.commit()
    conn.close()

    print(f'\n完成！寫入：{inserted} 筆，跳過：{skipped} 筆，錯誤：{errors} 筆')
    print(f'DB -> {DB_PATH}')

    if deploy:
        _deploy(DB_PATH)

    return inserted


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
