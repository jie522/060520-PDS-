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

# ── 手機掃碼回報（設備保養 P3，見 docs/equipment-maintenance.md）──────
# 預設 False：程式只綁 127.0.0.1，跟以前完全一樣，手機連不進來。
# 改成 True 之後才會改綁 0.0.0.0，讓同一個內網的手機用 http://<這台電腦IP>:5088/m 進來。
#
# 打開之前請先確認三件事：
#   1. 這台電腦要常開（關機手機就進不來），建議固定一台當服務端，不要用個人電腦
#   2. Windows 防火牆要放行 5088 埠的「私人網路」連入
#   3. 系統**沒有登入機制**——同一個內網掃到 QR 的人都能回報保養、看設備資料
MOBILE_ACCESS = True
# 產生 QR 標籤時要寫進去的網址開頭（例如 'http://192.168.1.50:5088'）。
# 留空就用列印當下瀏覽器的網址——但桌面應用的網址是 http://127.0.0.1:5088，
# 那樣印出來的 QR 手機掃了會連到手機自己，什麼都不會出現，所以這裡一定要填服務端的內網 IP。
MOBILE_BASE_URL = 'http://192.168.1.163:5088'
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

# ── 管理頁面：Google 試算表（報廢統計）設定 ─────────────────────
SCRAP_SHEET_ID  = '1v8mKx0m7RyUSjzOeBJMy576hqvGKyerxyNQo4pfVw04'
SCRAP_SHEET_GID = '0'

# ── 管理頁面：Google 試算表（生產日報表）設定 ────────────────────
# 注意：此分頁無「機台名稱」「備註」欄位
PROD_REPORT_SHEET_ID   = '1U_fY-DuNyTPa54PrSy2q9c_JGWqHa48R1fUmA6XBzU4'
PROD_REPORT_SHEET_NAME = 'P5.3生產日報表data_ref'  # 明細與圖表共用同一分頁，用名稱比 GID 穩定

# ── 管理頁面：Google 試算表（ABC 分類）設定 ──────────────────────
# 以品號+製程代號為 key，Q欄 = A/B/C 分類
CATEGORY_SHEET_ID  = '1gBdHsnEpTb75PUR2LTm3ZSWgVBHSAjDFgtEyzh8dQKs'
CATEGORY_SHEET_GID = '3658929'

# ── 管理頁面：Google 試算表（員工登錄系統，工號對應姓名）設定 ────
EMPLOYEE_SHEET_ID  = '12LuQyBhebmzr0WyTL4Uk36i1osVU9d2J17YCxSL6A5k'
EMPLOYEE_SHEET_GID = '884596320'

# ── 管理頁面：Google 試算表（加班統計）設定 ────
OVERTIME_SHEET_ID  = '1hH2bbU3jNtQ_ddG8D9cr5qJc1qolf30gT9YjqhpORno'
OVERTIME_SHEET_NAME = 'M2.加班人員'

# ── 管理頁面：Google 試算表（出勤統計）設定 ────
ATTENDANCE_SHEET_ID  = '1hH2bbU3jNtQ_ddG8D9cr5qJc1qolf30gT9YjqhpORno'
ATTENDANCE_SHEET_GID = '1760796831'

# ── 管理頁面：Google 試算表（採購登入表）設定 ────
PURCHASE_SHEET_ID  = '1HVSiu4GdO8R8ABkjM5oObLr-e7ZEkTwY5em9Vf9QoDA'
PURCHASE_SHEET_NAME = '主表'

# ── 批成本計算（共用區 Excel 範本）設定 ────────────────────────
BATCH_COST_FILE_PATH = r'\\192.168.1.99\共用區\品保加工共用平台\加工課自主巡檢表(空白)\【批成本計算】.xlsx'
BATCH_COST_TOOL_SHEET    = '刀具資料'    # 單價/刃數對照
BATCH_COST_TOOLMAP_SHEET = '刀表'        # 品號→T1~T39 刀具配置
BATCH_COST_RECORD_SHEET  = '批成本計算'  # 結果輸出（不存在時自動建立）
BATCH_COST_LIFESPAN_SHEET = '製令與刀具壽命'  # 每把刀具使用次數明細（不存在時自動建立）
BATCH_COST_EQUIPMENT_SHEET = '設備耗用設定'   # 每台設備電流/電壓/每日耗油費用（不存在時自動建立）
BATCH_COST_GLOBAL_SHEET   = '批成本全域設定'  # 每度電價等全域參數，單列（不存在時自動建立）

# ── 管理頁面：治檢具清單（PDM 資料夾）設定 ────────────────────────
JIG_VAULT_PATH = r'C:\MAXCLAW\06-生技課\01-模檢治具(2026後)'
JIG_SUBMITTERS = ['楊良捷', '陳銘仁', '林昭佑']
# 新增申請單用的空白 Excel 範本（vault 內路徑）
JIG_APPLY_TEMPLATE_XLSM = r'C:\MAXCLAW\06-生技課\00-空白範本\PP-M-041機器模檢治具申請單(2026後).xlsm'

# ── 管理頁面：設計變更通知單（PDM 資料夾）設定 ──────────────────────
DCN_VAULT_PATH = r'C:\MAXCLAW\00-研發部\02-文件資料\02-設計變更通知單'
# 新增申請單用的結構範本：PDM Admin 登記的「新增精靈」範本檔位置查不到（無讀取 API），
# 改用一筆既有的舊申請單當結構範本（只讀取複製，不會覆寫這筆原始記錄，
# 複製後所有申請專屬欄位都會被新申請的值蓋掉，見 docs/dcn-index.md）
DCN_APPLY_TEMPLATE_SOURCE = r'C:\MAXCLAW\00-研發部\02-文件資料\02-設計變更通知單\RR2607041\RR2607041設計變更申請單.xlsm'

# ── ZUMEN 線上圖面管理 GUI（Node.js 工具，選配）────────────────────
# 本機有 260704-zumen 工具（https://github.com/jie522/260704-zumen）資料夾時，
# 管理頁 ZUMEN 子頁可一鍵啟動並開啟。沒有的電腦按鈕會顯示不可用，不影響其他功能。
# 內網無外網電腦部署：整個 260704-zumen 資料夾（含已 npm install 的 node_modules/
# 跟免安裝 Node.js 可攜版 _node/，開發機上已備好）直接複製過去即可，不需要在目標
# 電腦上跑 npm install 或另外安裝 Node.js；只要把下面路徑改成該電腦上的實際位置。
ZUMEN_GUI_DIR  = r'D:\claude code\260704 ZUMEN圖面管理'
ZUMEN_GUI_PORT = 3456

# ── CNC 程式管理（網路資料夾）設定 ────────────────────────────────
CNC_PROGRAM_ROOT_PATH = r'\\192.168.1.99\加工部-資料夾\【技術資料】\P.程式'
CNC_PROGRAM_TEMPLATE_PATH = CNC_PROGRAM_ROOT_PATH + r'\【空白範本】'

# ── 設備管理（設備主檔／編碼／照片，網路資料夾）設定 ──────────────
# 詳見 docs/equipment-master.md。設備資料正本是 equipment.db，下面的 Excel 只是匯入來源。
EQUIPMENT_ROOT_PATH   = r'\\192.168.1.99\加工部-資料夾\【技術資料】\D.設備'
EQUIPMENT_CODING_XLSX = EQUIPMENT_ROOT_PATH + r'\設施設備編碼原則.xlsx'
EQUIPMENT_CODING_SHEET = 'PDM新格式編碼'   # 只讀這個分頁，其他分頁（APS設備對照等）不使用
EQUIPMENT_PHOTO_ROOT  = EQUIPMENT_ROOT_PATH + r'\設備照片'
EQUIPMENT_TECH_ROOT   = EQUIPMENT_ROOT_PATH + r'\設備技術資料'
# 異動歷程匯出用的獨立檔案（刻意不寫進 EQUIPMENT_CODING_XLSX——那份的 PDM編碼欄是公式，
# openpyxl 開檔存檔會把公式快取值洗掉，2026-07 實測踩過，見 docs/equipment-master.md）
EQUIPMENT_HISTORY_EXPORT_XLSX = EQUIPMENT_ROOT_PATH + r'\設備歷程.xlsx'
# 設備主檔資料庫（2026-08-07 從本機搬到網芳，讓多人共用同一份設備資料）：
# SQLite 直接放 Windows 網芳給多人同時讀寫，鎖檔機制不像本機硬碟可靠，這是刻意接受的
# 風險換來的多人共用——app.py 開連線時有加 busy timeout 降低「database is locked」，
# 但仍應避免多人「同時」大量寫入（例如同時上傳多張照片、同時編輯同一台設備）。
EQUIPMENT_DB_PATH = EQUIPMENT_ROOT_PATH + r'\equipment.db'
# ── 設備保養基準書（見 docs/equipment-maintenance.md）─────────────
# 基準書資料存在 equipment.db 的 mt_* 表；下面是附件與匯出檔案的存放位置。
EQUIPMENT_MAINT_ROOT  = EQUIPMENT_ROOT_PATH + r'\保養基準書'          # 基準書附件（PDF/圖片），資料夾不存在時自動建立
EQUIPMENT_MAINT_PHOTO_ROOT = EQUIPMENT_ROOT_PATH + r'\保養回報照片'    # 現場回報時拍的照片（依設備編碼分資料夾）
EQUIPMENT_MAINT_EXPORT_XLSX = EQUIPMENT_ROOT_PATH + r'\設備保養基準書.xlsx'   # 匯出用獨立檔案
# 舊的維護一覽表（一次性匯入來源，匯入後就不再使用這個檔案）
EQUIPMENT_MAINT_LEGACY_XLSX = EQUIPMENT_TECH_ROOT + r'\設備維護一覽表.xlsx'

# ── 油品管理（油品主檔／MSDS／更換記錄，網路資料夾）設定 ──────────
# 詳見 docs/oil-management.md。油品資料正本是 oil.db（跟 equipment.db 一樣放網芳共用），
# MSDS/簡介 PDF 實體檔案留在網芳，系統只建索引與連結。
OIL_ROOT_PATH   = r'\\192.168.1.99\加工部-資料夾\【技術資料】\O.油品'
OIL_MSDS_ROOT   = OIL_ROOT_PATH + r'\MSDS'          # MSDS，子資料夾「報廢」= 停用油品
OIL_DOC_ROOT    = OIL_ROOT_PATH + r'\油品簡介'       # 型錄/使用注意事項等文件
OIL_DB_PATH     = OIL_ROOT_PATH + r'\oil.db'        # 油品主檔資料庫（網芳共用，多人同時使用）
# 更換記錄的一次性匯入來源（匯入後正本在 oil.db，不再回頭改這個檔案）
OIL_CHANGE_LEGACY_XLSX = OIL_ROOT_PATH + r'\油品更換記錄表.xlsx'
# 匯出用獨立檔案（每次重建，不覆寫上面那份舊表）
OIL_EXPORT_XLSX = OIL_ROOT_PATH + r'\油品主檔.xlsx'

# ── ERP SQL Server 連線設定（BOM查詢用）──────────────────────────
# 連線至 192.168.1.140 的 Computech ERP YC01 資料庫
ERP_SQL_SERVER   = '192.168.1.140'      # SFT/ERP SQL Server 位址
ERP_SQL_DATABASE = 'YC01'               # YC01 公司 ERP 資料庫
ERP_SQL_USERNAME = 'sa'                 # SQL Server 帳號
ERP_SQL_PASSWORD = 'dsc55877948'        # SQL Server 密碼
