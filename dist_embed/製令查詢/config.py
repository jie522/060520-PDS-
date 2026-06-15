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

# ── 申請單功能設定 ────────────────────────────────────────────────
# 申請單存放路徑（NAS）
APPLICATION_STORE_PATH = r'\\192.168.1.99\加工部-資料夾\部門別日常管理-製一\@內部業務聯絡單'
# Word 範本資料夾（None = 自動使用 app.py 旁邊的 static/form_templates/，無需修改）
FORM_TEMPLATES_DIR = None

# ── 管理頁面：Google 試算表（請假單）設定 ────────────────────────
LEAVE_SHEET_ID   = '1hH2bbU3jNtQ_ddG8D9cr5qJc1qolf30gT9YjqhpORno'
LEAVE_SHEET_NAME = 'M3.請假單'

# ── 管理頁面：Google 試算表（員工登錄系統，工號對應姓名）設定 ────
EMPLOYEE_SHEET_ID  = '12LuQyBhebmzr0WyTL4Uk36i1osVU9d2J17YCxSL6A5k'
EMPLOYEE_SHEET_GID = '884596320'

# ── 管理頁面：Google 試算表（加班統計）設定 ────
OVERTIME_SHEET_ID  = '1hH2bbU3jNtQ_ddG8D9cr5qJc1qolf30gT9YjqhpORno'
OVERTIME_SHEET_NAME = 'M2.加班人員'

# ── 管理頁面：治檢具清單（PDM 資料夾）設定 ────────────────────────
JIG_VAULT_PATH = r'C:\MAXCLAW\06-生技課\01-模檢治具(2026後)'
JIG_SUBMITTERS = ['楊良捷', '陳銘仁', '林昭佑']

# ── ERP SQL Server 連線設定（BOM查詢用）──────────────────────────
# 連線至 192.168.1.140 的 Computech ERP YC01 資料庫
ERP_SQL_SERVER   = '192.168.1.140'      # SFT/ERP SQL Server 位址
ERP_SQL_DATABASE = 'YC01'               # YC01 公司 ERP 資料庫
ERP_SQL_USERNAME = 'sa'                 # SQL Server 帳號
ERP_SQL_PASSWORD = 'dsc55877948'        # SQL Server 密碼
