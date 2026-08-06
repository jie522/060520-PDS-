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
              'build_equipment_index.py']:
        src = os.path.join(SRC, f)
        dst = os.path.join(APP, f)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f'  [OK] {f}')

    # 1b. cnc_program_index.db / equipment.db 是**執行期資料庫**，桌面應用執行中會直接
    # 寫入（照片上傳、設備編輯、CNC 程式上傳…），不能跟原始碼一樣無條件覆蓋——
    # 2026-07-30 實測發生過：使用者在桌面應用貼上照片存進 dist 那份 equipment.db，
    # 之後這支腳本又把開發機這份（沒有那筆照片）蓋過去，照片檔案還在網芳但資料庫
    # 記錄整筆消失。這兩個資料庫改成只能透過各自的 --deploy 明確指令推送
    # （python build_equipment_index.py --deploy / python build_cnc_program_index.py --deploy），
    # 不再是這支腳本的預設行為。
    # calendar.db（管理頁行事曆的工作日設定）同理，而且它**沒有任何 build 工具會產生**，
    # 純粹是使用者一天一天點出來的資料，被覆蓋就真的救不回來。這支腳本本來就只複製
    # 上面那份明確清單，不會碰 *.db；日後要加檔案時千萬不要把這三個資料庫加進去。
    print('  [SKIP] cnc_program_index.db / equipment.db / calendar.db（執行期資料庫，'
          '用 build_*_index.py --deploy 明確推送，不隨原始碼同步覆蓋）')

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
