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
WINDOW_TITLE = '詠基-加工部查詢系統PDIS'

# ── ERP SQL Server 連線設定（BOM查詢用）──────────────────────────
# 請填入正確的 SQL Server 連線資訊，儲存後重新啟動程式即可生效
ERP_SQL_SERVER   = '192.168.1.212'      # SQL Server 位址
ERP_SQL_DATABASE = 'MCPDB'              # ERP 資料庫名稱（請確認）
ERP_SQL_USERNAME = 'max'                # SQL Server 帳號（空字串=使用 Windows 驗證）
ERP_SQL_PASSWORD = 'max'               # SQL Server 密碼

# BOM 查詢資料表名稱（Computech ERP 預設）
BOM_HEADER_TABLE = 'BOMMH'   # BOM 母件主檔
BOM_DETAIL_TABLE = 'BOMMD'   # BOM 子件明細
