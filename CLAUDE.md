# 加工部查詢系統 PDS（製令查詢）

詠基公司加工課用的製令查詢 Web 應用。Flask 後端 + Bootstrap 5 前端，以 Edge/Chrome --app 模式呈現為桌面視窗。

## 技術架構

- **後端**：Flask（Python 3.11）
- **前端**：HTML5 + Bootstrap 5.3 + Vanilla JS
- **資料來源**：SSRS（ERP 報表）、PDM Vault（SolidWorks）、ServCloud API（設備）、ERP SQL Server（Computech YC01）、Google 試算表（管理頁：請假/出勤/加班/報廢/生產日報/分類對照）
- **本地 DB**：SQLite（pdm_search.db 圖面索引含 dcn_index 表、zume_drawings.db 技術規格）
- **部署**：Python Embeddable Package，免安裝，double-click VBScript 啟動

## 關鍵設定（config.py）

```python
SSRS_BASE_URL = 'http://192.168.1.212/ReportServer'
ERP_SQL_SERVER = '192.168.1.140'   # Computech YC01 資料庫
ERP_SQL_DATABASE = 'YC01'
FLASK_PORT = 5088
PDM Vault 名稱：MAXCLAW
```

## 常用指令

```bash
# 開發模式啟動
python app.py

# 打包成可攜式發行版（會下載 Python Embeddable）
python build_embed.py

# 快速同步修改到 dist_embed（不重下載 Python）
python sync_to_dist.py

# 重建 PDM 圖面索引（完整）
python build_pdm_index.py

# 增量更新 + 同步到 dist_embed
python build_pdm_index.py --update --deploy

# 重建治檢具清單索引（+ 同步）
python build_jig_index.py --deploy

# 重建設計變更通知單（DCN）索引（+ 同步）
python build_dcn_index.py --deploy
```

## 重要檔案

| 檔案 | 說明 |
|------|------|
| `app.py` | Flask 主程式，所有 API 端點與商業邏輯 |
| `main.py` | GUI 啟動器，開 Edge/Chrome --app 模式視窗 |
| `config.py` | 外部設定（IP、port、路徑） |
| `build_pdm_index.py` | PDM 圖面索引重建工具（v1.2，SQL Server 直連） |
| `build_jig_index.py` | 治檢具清單索引重建工具 |
| `build_dcn_index.py` | 設計變更通知單（DCN）索引重建工具 |
| `build_embed.py` | 打包成可攜式版本 |
| `sync_to_dist.py` | 快速同步修改到 dist_embed |
| `templates/index.html` | 主畫面（製令搜尋） |
| `templates/drawing.html` | 圖面查詢 |
| `templates/bom.html` | BOM 查詢 |
| `templates/routing.html` | 途程查詢 |
| `templates/equipment.html` | 設備稼動查詢 |
| `templates/production.html` | 生產日報 |
| `templates/print_report.html` | 列印工單 |
| `templates/management.html` | 管理頁（請假/出勤/加班/申請單/治檢具/DCN/報廢/生產日報P5.3/生產報工P2） |

## dist_embed 結構

```
dist_embed/製令查詢/
├── _python/          # Python 3.11.9 Embeddable（免安裝）
├── _app/             # 程式碼（app.py, main.py, config.py, templates/）
│   ├── pdm_search.db # PDM 圖面索引
│   └── zume_drawings.db
├── 製令查詢.vbs      # 主啟動捷徑（雙擊執行）
├── 除錯模式執行.bat
└── config.py         # 使用者可編輯的設定
```

## PDM 索引（build_pdm_index.py）

- 連線 SolidWorks MAXCLAW Vault（COM API）
- 品名/發圖日期優先從 PDM SQL Server 直連讀取（ADO via win32com）
- SQL 失敗時備援使用 COM IEdmEnumeratorVariable10 多策略讀取
- 掃描階段只存路徑字串，不保留 COM 物件（避免狀態污染）
- 索引存於 `%LOCALAPPDATA%\PDMSearch\pdm_search.db`

## 管理頁面（management.html）

頂層導覽列「管理」分頁，左側選單分三群組（各以 📁 圖示 + 標題）：

- **人員勤務**：請假統計、出勤統計、加班統計、申請單
- **生產管理**：生產日報表P5.3、生產報工統計P2、報廢統計
- **技術文件**：治檢具清單、設計變更通知單

各子頁資料來源（`config.py` 內 Google 試算表 ID/GID/分頁名）：

| 子頁 | 來源 | 備註 |
|------|------|------|
| 請假/加班 | `LEAVE_SHEET_ID`（M3.請假單 / M2.加班人員） | |
| 出勤統計 | `ATTENDANCE_SHEET_ID` | |
| 報廢統計 | `SCRAP_SHEET_ID` | |
| 生產日報表P5.3 | `PROD_REPORT_SHEET_ID` + `PROD_REPORT_SHEET_GID`（P5.3生產日報表data_ref，gid 567727173） | 明細與圖表同源；此分頁無機台名稱/備註 |
| 生產報工統計P2 | `CATEGORY_SHEET_ID`（K1_P2.ref，出站數量） | 與生產日報表是**不同**資料來源，數字本就不同 |
| 員工工號對照 | `EMPLOYEE_SHEET_ID` | 工號 → 姓名 |
| ABC 分類對照 | `CATEGORY_SHEET_ID` | key = 品號+製程代號 → A/B/C/L |

- Google 試算表以「知道連結的人可檢視」共用，後端用 gviz CSV 端點讀取（`fetch_google_sheet_csv`），無需金鑰。
- 生產日報/報工的分類：① 品號+製程對照表 → ② 機台名稱/代號/製程關鍵字備援（CNC→A、鑽→B、沖→C、車→L）。

## 設計變更通知單索引（build_dcn_index.py）

- 掃描 PDM Vault『00-研發部\02-文件資料\02-設計變更通知單』子資料夾，下載 xlsm 申請表。
- **PDM Card 變數值讀取自 Office 檔的 `docProps/custom.xml`**（Office Custom Properties），不是 worksheet cell——RR 新格式（2021+）的 cell 是空的，唯一可靠來源是 custom.xml。
- 讀取欄位：機型、提出人員、經辦、審核/核決、申請原因（PP_R_004_*_tasky 勾選項）、申請變更內容說明、工作流程狀態。
- 寫入 `pdm_search.db` 的 `dcn_index` 表。前端可點列開卡片彈窗、篩選提出人員、顯示已作廢/已結案。

## 搜尋語法

製令搜尋支援 AND/NOT：
- 空格 = AND（例：`FTB 夾管座`）
- `-前綴` = NOT（例：`FTB -舊版`）

## TTL 快取

SSRS 查詢有 120 秒 TTL 快取（`cache_get/cache_set`），避免頻繁呼叫。

## 版本管理

版本號格式：`VYYYYMMDD`（年月日）
- 定義位置：`app.py` 的 `APP_VERSION` 常數
- **每次完成更新後，務必修正版本號為當日期日期**（例：2026-06-15 修正後應改為 `V20260615`）
- 修正後執行 `python sync_to_dist.py` 同步至 dist_embed

## 導覽列設計規範

所有頁面的導覽列 (`<div class="header">`) 應遵循以下規範，確保視覺一致性。設計參考 ZUMEN 導覽列風格。

### 顏色
- **背景色**：純藍色 `#1565C0`（參考 ZUMEN）
- **標題文字色**：白色 `#FFFFFF`
- **Tab 文字色**（未激活）：半透明白色 `rgba(255,255,255,.85)`
- **Tab 背景色**（未激活）：半透明白色 `rgba(255,255,255,.15)`
- **Tab 文字色 & 背景**（激活/hover）：白色背景 + 深藍文字 `#1565C0`

### 字體 & 字號
- **全頁面字體**：`'Noto Sans TC', 'Source Han Sans TC', sans-serif`
- **標題**：字號 `1.1rem`、加粗 `700`、不換行 `nowrap`
- **Tab**：字號 `1rem`（比預設大 2 號）、加粗 `600`

### 間距與尺寸
- **Header padding**：上下 `0.4rem`、左右 `1rem`（以 `padding: 0.4rem 1rem 0` 為基準）
- **標題 padding-bottom**：`0.45rem`
- **標題右外邊距**：`0.5rem`
- **Tab padding**：上下 `0.3rem`、左右 `1rem`
- **Tab 間距**：`gap: 2px`
- **Tab border-radius**：`0`（直角，無圓角，參考 ZUMEN 風格）

### 動畫 & 其他效果
- **過渡效果**：`transition: background 0.15s`（hover/active 狀態）
- **陰影**：`box-shadow: 0 2px 6px rgba(0,0,0,.3)`
- **位置**：`position: sticky; top: 0; z-index: 200`（黏性定位，頁面置頂）

## 字體規範

全站統一採用 **Source Han Sans TC（思源黑體）**，以 Google Fonts 的 **Noto Sans TC** 名稱引入（同一字型，不同授權方命名）。此為強制規定，禁止改用宋體或其他字型。

### 引入方式
在 HTML `<head>` 中新增：
```html
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap" rel="stylesheet">
```

### CSS 設定
```css
body, * {
  font-family: 'Noto Sans TC', 'Source Han Sans TC', sans-serif;
}
```

### 字重對應
- `400`：正常（Regular）
- `500`：中等（Medium）
- `600`：加粗（SemiBold）
- `700`：粗體（Bold）

### 注意事項
- 禁止使用 `Noto Serif TC`、`Microsoft JhengHei`、`Segoe UI` 作為主字型
- 通用備援字型使用 `sans-serif`（非 `serif`）
- Adobe 名稱為「Source Han Sans TC」，可作為第二備援

## 語言規則

與使用者溝通（說明、進度、總結、錯誤訊息）一律使用繁體中文，不可使用英文或簡體中文。程式碼中的變數/函式/檔名等技術識別字維持英文，但註解一律繁體中文。
