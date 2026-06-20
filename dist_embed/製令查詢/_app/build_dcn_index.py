#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
設計變更通知單索引重建工具

掃描 PDM Vault『00-研發部\\02-文件資料\\02-設計變更通知單』底下的子資料夾，
下載 xlsm 申請表，優先從 docProps/custom.xml 讀取 PDM Card 變數值：
  - YC_日期       → issue_date
  - YC_機型       → product_model
  - YC_提出人員   → submitter
  - YC_經辦 / YC_經辦01 → handler
  - 工作流程狀態 CurrentState.Name → status

寫入 pdm_search.db 的 dcn_index 表。

用法：
  python build_dcn_index.py            # 重建索引
  python build_dcn_index.py --deploy   # 重建後同步至 dist_embed
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import sqlite3
import datetime
import zipfile
import xml.etree.ElementTree as ET

import config
from build_pdm_index import connect_vault, DB_DIR, DB_PATH, _deploy

SCHEMA = """
DROP TABLE IF EXISTS dcn_index;
CREATE TABLE dcn_index (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_name   TEXT NOT NULL,
    folder_path   TEXT NOT NULL UNIQUE,
    issue_date    TEXT,
    product_model TEXT,
    submitter     TEXT,
    handler       TEXT,
    reviewer      TEXT,
    approver      TEXT,
    reason        TEXT,
    description   TEXT,
    status        TEXT,
    indexed_at    TEXT DEFAULT (datetime('now'))
);
"""

_REASON_KEYS = [
    'PP_R_004_客戶要求_tasky',
    'PP_R_004_設計錯誤_tasky',
    'PP_R_004_物料需求_tasky',
    'PP_R_004_製造問題_tasky',
    'PP_R_004_功能改善_tasky',
    'PP_R_004_變更供應商_tasky',
    'PP_R_004_其他_tasky',
]

_NS_CUSTOM = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
_NS_VT     = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
_SKIP      = {'--', '-', '', 'None'}

_first_done = False


def get_subfolders(vault, root_path):
    root = vault.GetFolderFromPath(root_path)
    subs = []
    pos = root.GetFirstSubFolderPosition()
    while not pos.IsNull:
        subs.append(root.GetNextSubFolder(pos))
    return subs


def read_custom_props(local_path):
    """讀取 docProps/custom.xml 裡的 PDM Card 變數（name→value dict）"""
    try:
        with zipfile.ZipFile(local_path, 'r') as zf:
            if 'docProps/custom.xml' not in zf.namelist():
                return {}
            raw = zf.read('docProps/custom.xml')
            root = ET.fromstring(raw)
            props = {}
            for prop in root.findall(f'{{{_NS_CUSTOM}}}property'):
                name = prop.get('name', '')
                if not name:
                    continue
                elem = prop.find(f'{{{_NS_VT}}}lpwstr')
                if elem is not None and elem.text:
                    v = elem.text.strip()
                    if v and v not in _SKIP:
                        props[name] = v
            return props
    except Exception:
        return {}


def fmt_date(v):
    if not v:
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip().replace('/', '-')
    return s[:10] if len(s) >= 10 else s


def read_xlsm(folder, f):
    """讀取 xlsm 申請表，回傳欄位 dict 或 None（非 xlsm / 讀取失敗則跳過）"""
    global _first_done

    if not f.Name.lower().endswith('.xlsm'):
        return None

    try:
        f.GetFileCopy(0, '', folder.ID, 1)
        local_path = f.GetLocalPath(folder.ID)
        if not os.path.exists(local_path):
            return None

        # 優先讀 docProps/custom.xml（PDM Card 同步的位置）
        props = read_custom_props(local_path)

        if not _first_done:
            _first_done = True
            try:
                print(f'  [診斷] {folder.Name}/{f.Name} custom.xml props:')
                for k, v in sorted(props.items()):
                    print(f'    {k} = {v!r}')
            except UnicodeEncodeError:
                pass

        issue_date    = fmt_date(
            props.get('YC_日期') or props.get('YC_申請日期') or props.get('YC_發行日期', ''))
        product_model = (
            props.get('YC_機型') or props.get('YC_機種', ''))
        submitter     = (
            props.get('YC_提出人員') or props.get('YC_申請人') or props.get('YC_提出人', ''))
        handler       = (
            props.get('YC_經辦') or props.get('YC_經辦01') or
            props.get('YC_承辦人') or props.get('YC_設計者', ''))
        reviewer      = (props.get('YC_審核', '') or '')    # 研發主管
        approver      = (props.get('YC_核決', '') or '')    # 申請主管

        # 申請原因（篩選有打勾的）
        checked = []
        for k in _REASON_KEYS:
            v = props.get(k, '')
            if v.startswith('■'):
                checked.append(v.lstrip('■').strip())
        reason = '、'.join(checked)

        # 申請變更內容說明
        description = (
            props.get('PP_R_004_變更規格敘述_tasky') or
            props.get('YC_申請變更內容說明') or
            props.get('YC_變更說明') or ''
        )

        # 狀態從 PDM 工作流程取
        status = ''
        try:
            status = str(f.CurrentState.Name).strip()
        except Exception:
            pass

        return {
            'issue_date':    issue_date,
            'product_model': product_model,
            'submitter':     submitter,
            'handler':       handler,
            'reviewer':      reviewer,
            'approver':      approver,
            'reason':        reason,
            'description':   description,
            'status':        status,
        }
    except Exception:
        return None


def init_db(conn):
    conn.cursor().executescript(SCHEMA)
    conn.commit()


def rebuild(deploy=False):
    os.makedirs(DB_DIR, exist_ok=True)

    print('連線 PDM Vault ...')
    vault = connect_vault()

    print(f'掃描資料夾：{config.DCN_VAULT_PATH}')
    subfolders = get_subfolders(vault, config.DCN_VAULT_PATH)
    total = len(subfolders)
    print(f'  找到 {total} 個子資料夾')

    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    cur = conn.cursor()

    inserted = skipped = errors = 0

    for idx, folder in enumerate(subfolders, 1):
        try:
            record = None

            # 找 xlsm 檔（取第一個成功讀取的）
            pos = folder.GetFirstFilePosition()
            while not pos.IsNull:
                f = folder.GetNextFile(pos)
                record = read_xlsm(folder, f)
                if record:
                    break

            if not record:
                # 無法讀取 xlsm，仍記錄資料夾名稱與狀態
                status = ''
                try:
                    pos2 = folder.GetFirstFilePosition()
                    if not pos2.IsNull:
                        ff = folder.GetNextFile(pos2)
                        status = str(ff.CurrentState.Name).strip()
                except Exception:
                    pass
                record = {
                    'issue_date': '', 'product_model': '',
                    'submitter': '', 'handler': '', 'reviewer': '',
                    'approver': '', 'reason': '', 'description': '',
                    'status': status,
                }
                skipped += 1
            else:
                inserted += 1

            folder_path = os.path.join(config.DCN_VAULT_PATH, folder.Name)
            cur.execute(
                'INSERT OR REPLACE INTO dcn_index '
                '(folder_name, folder_path, issue_date, product_model, '
                ' submitter, handler, reviewer, approver, reason, description, status) '
                'VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                (folder.Name, folder_path,
                 record['issue_date'], record['product_model'],
                 record['submitter'], record['handler'],
                 record['reviewer'], record['approver'],
                 record['reason'], record['description'], record['status'])
            )

            if record['issue_date'] or record['product_model']:
                try:
                    print(f'  [{idx}/{total}] {folder.Name}: '
                          f'日期={record["issue_date"]} 機型={record["product_model"]} '
                          f'提出={record["submitter"]} 經辦={record["handler"]} '
                          f'狀態={record["status"]}')
                except UnicodeEncodeError:
                    print(f'  [{idx}/{total}] {folder.Name}: 日期={record["issue_date"]}')

        except Exception as e:
            errors += 1
            try:
                print(f'  [{idx}/{total}] {folder.Name}: ERR={e}')
            except UnicodeEncodeError:
                print(f'  [{idx}/{total}]: ERR')

        if idx % 20 == 0:
            conn.commit()
            try:
                print(f'  進度 {idx}/{total}  寫入:{inserted}  跳過:{skipped}  錯誤:{errors}')
            except UnicodeEncodeError:
                pass

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
    print('  設計變更通知單索引重建工具')
    print(f'  時間: {start:%Y-%m-%d %H:%M:%S}')
    print('=' * 55)

    rebuild(deploy=deploy)

    elapsed = (datetime.datetime.now() - start).total_seconds()
    print(f'\n總耗時: {elapsed:.1f} 秒')
