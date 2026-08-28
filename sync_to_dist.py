"""
快速同步源碼到 dist_embed（不重新下載 Python）
用法: python sync_to_dist.py
"""
import os
import shutil

SRC  = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(SRC, 'dist_embed', 'PDS系統')
APP  = os.path.join(DIST, '_app')

def sync():
    print('同步源碼到 dist_embed ...')

    # 1. 複製 Python 檔案（原始碼／建置工具，隨時同步安全）
    for f in ['app.py', 'main.py', 'config.py', 'build_jig_index.py', 'build_pdm_index.py',
              'build_dcn_index.py', 'build_cnc_program_index.py',
              'build_equipment_index.py', 'build_oil_index.py']:
        src = os.path.join(SRC, f)
        dst = os.path.join(APP, f)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f'  [OK] {f}')

    # 1b. cnc_program_index.db / calendar.db 是**執行期資料庫**，桌面應用執行中會直接
    # 寫入（CNC 程式上傳、行事曆工作日設定…），不能跟原始碼一樣無條件覆蓋——
    # 2026-07-30 實測發生過類似狀況：使用者在桌面應用寫入的資料庫，被這支腳本拿開發機
    # 沒有那筆資料的舊版蓋過去，記錄整筆消失。cnc_program_index.db 改成只能透過
    # `python build_cnc_program_index.py --deploy` 明確指令推送，不是這支腳本的預設行為。
    # calendar.db 同理，而且它**沒有任何 build 工具能重建**，純粹是使用者一天一天點出來的
    # 資料，被覆蓋就真的救不回來。
    # equipment.db 則是 2026-08-07 起改放網芳共用（config.EQUIPMENT_DB_PATH），
    # 開發機跟桌面應用讀寫同一份網芳檔案，已經沒有「本機兩份互相覆蓋」的問題，
    # 不需要也不應該被這支腳本碰。
    print('  [SKIP] cnc_program_index.db / calendar.db（執行期資料庫，cnc_program_index.db '
          '用 build_cnc_program_index.py --deploy 明確推送，calendar.db 不隨原始碼同步覆蓋；'
          'equipment.db 已改放網芳共用，不在本機，不受此腳本影響）')

    # 2. 複製 config.py 到根目錄（使用者可編輯）
    shutil.copy2(os.path.join(SRC, 'config.py'), DIST)
    print(f'  [OK] config.py → 根目錄')

    # 3. 同步 templates
    tpl_src = os.path.join(SRC, 'templates')
    tpl_dst = os.path.join(APP, 'templates')
    if os.path.exists(tpl_dst):
        shutil.rmtree(tpl_dst)
    shutil.copytree(tpl_src, tpl_dst)
    n = len(os.listdir(tpl_dst))
    print(f'  [OK] templates/ ({n} 個檔案)')

    # 3b. 同步 static（含 pds-ibm.css 與 form_templates/）
    static_src = os.path.join(SRC, 'static')
    static_dst = os.path.join(APP, 'static')
    if os.path.exists(static_src):
        if os.path.exists(static_dst):
            shutil.rmtree(static_dst)
        shutil.copytree(static_src, static_dst)
        print(f'  [OK] static/')

    # 4. 清除快取
    for cache_dir in [
        os.path.join(APP, '__pycache__'),
        os.path.join(DIST, '__pycache__'),
        os.path.join(DIST, '_edge_data'),
    ]:
        if os.path.exists(cache_dir):
            shutil.rmtree(cache_dir)
            print(f'  [DEL] 清除 {os.path.basename(cache_dir)}/')

    print('\n同步完成！')

if __name__ == '__main__':
    sync()
