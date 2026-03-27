# ════════════════════════════════════════
#  製令查詢系統 — 設定檔（GUI 版）
#  修改此檔案後重新啟動程式即可生效
# ════════════════════════════════════════

# SSRS 連線設定
SSRS_BASE_URL = 'http://192.168.1.212/ReportServer'
SSRS_USERNAME = 'max'
SSRS_PASSWORD = 'max'

# 報表路徑
REPORT_PATHS = {
    'unfinished': '/資料查詢/全公司-未完工製令情況',
    'all': '/資料查詢/全公司-製令情況',
    'daily': '/加工課-生產日報表',
    'efficiency': '/加工課/加工部總效率',
}

# PDM 圖面索引資料庫路徑（自動偵測當前使用者，無需手動修改）
import os as _os
PDM_DB_PATH = _os.path.join(_os.environ.get('LOCALAPPDATA', r'C:\Users\Public'), 'PDMSearch', 'pdm_search.db')

# Flask 設定（GUI 版綁定 127.0.0.1，不對外開放）
FLASK_HOST = '127.0.0.1'
FLASK_PORT = 5088
FLASK_DEBUG = False   # EXE 模式請保持 False

# 視窗設定
WINDOW_WIDTH = 1400
WINDOW_HEIGHT = 900
WINDOW_TITLE = '製令查詢系統'
