"""
製令查詢系統（GUI 版）— Embeddable 打包腳本
使用 Python 官方 Embeddable Package + pywebview，
打包成獨立可攜帶的桌面應用程式。
執行: python build_embed.py
"""
import os, sys, urllib.request, zipfile, shutil, subprocess, textwrap

# ═══════════════════════════════════════════════
#  設定
# ═══════════════════════════════════════════════
PY_VER    = '3.11.9'          # 穩定版，DLL 相容性最佳
ARCH      = 'amd64'
PY_ZIP    = f'python-{PY_VER}-embed-{ARCH}.zip'
PY_URL    = f'https://www.python.org/ftp/python/{PY_VER}/{PY_ZIP}'
GET_PIP   = 'https://bootstrap.pypa.io/get-pip.py'
DIST      = os.path.join('dist_embed', '製令查詢')
APP_SRC   = os.path.dirname(os.path.abspath(__file__))

PACKAGES  = ['flask', 'requests', 'requests-ntlm', 'pywin32']

def step(msg):
    print(f'\n{"─"*50}\n  {msg}\n{"─"*50}')

# ═══════════════════════════════════════════════
#  清除舊資料夾
# ═══════════════════════════════════════════════
step('1/7  清除舊打包資料夾')
if os.path.exists(DIST):
    shutil.rmtree(DIST)
os.makedirs(DIST, exist_ok=True)
py_dir = os.path.join(DIST, '_python')
os.makedirs(py_dir, exist_ok=True)

# ═══════════════════════════════════════════════
#  下載 Python Embeddable
# ═══════════════════════════════════════════════
step(f'2/7  下載 Python {PY_VER} Embeddable Package')
zip_local = os.path.join(DIST, PY_ZIP)
if not os.path.exists(zip_local):
    print(f'  下載中: {PY_URL}')
    def _progress(block, block_size, total):
        done = block * block_size
        pct  = min(done * 100 // total, 100) if total > 0 else 0
        print(f'\r  {pct:3d}%  {done//1024:,} KB', end='', flush=True)
    urllib.request.urlretrieve(PY_URL, zip_local, _progress)
    print()
else:
    print(f'  已存在，略過下載')

print(f'  解壓縮到 _python/ ...')
with zipfile.ZipFile(zip_local, 'r') as z:
    z.extractall(py_dir)
os.remove(zip_local)

# ═══════════════════════════════════════════════
#  啟用 pip（修改 _pth 檔案）
# ═══════════════════════════════════════════════
step('3/7  啟用 pip')
pth_files = [f for f in os.listdir(py_dir) if f.endswith('._pth')]
if pth_files:
    pth = os.path.join(py_dir, pth_files[0])
    content = open(pth, encoding='utf-8').read()
    content = content.replace('#import site', 'import site')
    # 明確加入 site-packages 路徑（相對於 python.exe）
    if '.\\Lib\\site-packages' not in content:
        content = content.rstrip('\n') + '\n.\\Lib\\site-packages\n'
    open(pth, 'w', encoding='utf-8').write(content)
    print(f'  已修改 {pth_files[0]}')

# 下載 get-pip.py
pip_py = os.path.abspath(os.path.join(py_dir, 'get-pip.py'))
print(f'  下載 get-pip.py ...')
urllib.request.urlretrieve(GET_PIP, pip_py)

python_exe = os.path.abspath(os.path.join(py_dir, 'python.exe'))
subprocess.run([python_exe, pip_py, '--quiet'], check=True)
os.remove(pip_py)
print('  pip 安裝完成')

# ═══════════════════════════════════════════════
#  安裝套件
# ═══════════════════════════════════════════════
step(f'4/7  安裝套件: {", ".join(PACKAGES)}')
pip_exe = os.path.abspath(os.path.join(py_dir, 'Scripts', 'pip.exe'))
subprocess.run(
    [pip_exe, 'install', '--quiet', '--no-warn-script-location'] + PACKAGES,
    check=True
)
print('  安裝完成')

# pywin32 特殊處理：DLL 需從 pywin32_system32 複製到 _python\ 才能載入
print('  複製 pywin32 DLL ...')
_pw32_dll_src = os.path.abspath(os.path.join(py_dir, 'Lib', 'site-packages', 'pywin32_system32'))
if os.path.isdir(_pw32_dll_src):
    for _dll in os.listdir(_pw32_dll_src):
        if _dll.endswith('.dll'):
            _src_dll = os.path.join(_pw32_dll_src, _dll)
            _dst_dll = os.path.join(py_dir, _dll)
            shutil.copy2(_src_dll, _dst_dll)
            print(f'    {_dll}')
else:
    print('  （找不到 pywin32_system32，略過）')

# ═══════════════════════════════════════════════
#  複製應用程式檔案
# ═══════════════════════════════════════════════
step('5/7  複製應用程式')
app_dst = os.path.join(DIST, '_app')
os.makedirs(app_dst, exist_ok=True)

copy_files = ['app.py', 'main.py', 'config.py']
for f in copy_files:
    src = os.path.join(APP_SRC, f)
    if os.path.exists(src):
        shutil.copy2(src, app_dst)
        print(f'  複製 {f}')

# 複製 PDMSearch 索引資料庫（讓各電腦共用同一份索引）
_pdm_db_candidates = [
    r'C:\Users\user\AppData\Local\PDMSearch\pdm_search.db',
    os.path.join(APP_SRC, 'pdm_search.db'),
]
for _pdm_src in _pdm_db_candidates:
    if os.path.exists(_pdm_src):
        shutil.copy2(_pdm_src, os.path.join(app_dst, 'pdm_search.db'))
        size_mb = os.path.getsize(_pdm_src) / 1024 / 1024
        print(f'  複製 pdm_search.db  ({size_mb:.1f} MB) ← 各電腦共用')
        break
else:
    print('  ⚠ 找不到 pdm_search.db，圖面查詢在目標電腦可能無法使用')

# 複製 templates 資料夾
templates_src = os.path.join(APP_SRC, 'templates')
templates_dst = os.path.join(app_dst, 'templates')
if os.path.exists(templates_src):
    shutil.copytree(templates_src, templates_dst)
    n = len(os.listdir(templates_dst))
    print(f'  複製 templates/ ({n} 個檔案)')

# config.py 也放到根目錄（讓使用者直接編輯）
shutil.copy2(os.path.join(APP_SRC, 'config.py'), DIST)

# ═══════════════════════════════════════════════
#  建立啟動器
# ═══════════════════════════════════════════════
step('6/7  建立啟動器')

# VBScript 啟動器（靜默執行，不顯示 CMD 視窗）
vbs = os.path.join(DIST, '製令查詢.vbs')
open(vbs, 'w', encoding='utf-8').write(textwrap.dedent(r"""
    ' 製令查詢系統啟動器（GUI 版）
    Dim shell : Set shell = CreateObject("WScript.Shell")
    Dim fso   : Set fso   = CreateObject("Scripting.FileSystemObject")
    Dim base  : base = fso.GetParentFolderName(WScript.ScriptFullName)
    Dim pyexe : pyexe = base & "\_python\pythonw.exe"
    Dim script: script = base & "\_app\main.py"
    shell.Run Chr(34) & pyexe & Chr(34) & " " & Chr(34) & script & Chr(34), 0, False
""").strip())
print('  建立 製令查詢.vbs')

# 批次捷徑（方便測試，顯示輸出）
bat = os.path.join(DIST, '除錯模式執行.bat')
open(bat, 'w', encoding='utf-8').write(
    '@echo off\nchcp 65001>nul\necho 製令查詢系統（除錯模式）...\n'
    r'"%~dp0_python\python.exe" "%~dp0_app\main.py"' + '\npause\n'
)
print('  建立 除錯模式執行.bat')

# 使用說明
readme = os.path.join(DIST, '使用說明.txt')
open(readme, 'w', encoding='utf-8').write(textwrap.dedent(f"""
    製令查詢系統（GUI 版）使用說明
    ==============================

    【啟動方式】
      雙擊「製令查詢.vbs」即可啟動
      系統會開啟獨立的應用程式視窗（不需要瀏覽器）

    【設定修改】
      編輯同目錄的「config.py」可修改 SSRS/PDM 設定
      修改後重新啟動 .vbs 生效

    【如果無法啟動】
      雙擊「除錯模式執行.bat」查看錯誤訊息
      確認電腦已安裝 Microsoft Edge WebView2 Runtime
      （Windows 10 2004 以上版本通常已內建）

    【PDM 圖面查詢】
      確認 config.py 中的 PDM_DB_PATH 路徑正確

    【Python 版本】
      內建 Python {PY_VER}（Embeddable，無需安裝）
      無需安裝 Python / VC++ Runtime / 任何套件

    【注意事項】
      需連接公司網路才能查詢 SSRS 報表
      每台電腦獨立執行，互不影響
""").strip())

# ═══════════════════════════════════════════════
#  統計結果
# ═══════════════════════════════════════════════
step('7/7  完成')
total = sum(
    os.path.getsize(os.path.join(dp, f))
    for dp, _, files in os.walk(DIST)
    for f in files
) / 1024 / 1024

print(f'  輸出目錄: {os.path.abspath(DIST)}')
print(f'  總大小:   {total:.1f} MB')
print()
print('  將整個「製令查詢」資料夾複製到其他電腦即可使用')
print('  雙擊「製令查詢.vbs」啟動')
