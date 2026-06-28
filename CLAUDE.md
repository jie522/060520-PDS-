# 加工部查詢系統 PDS（製令查詢）

詠基公司加工課用的製令查詢 Web 應用。Flask 後端 + Bootstrap 5 前端，以 Edge/Chrome --app 模式呈現為桌面視窗。

## 技術架構

- **後端**：Flask（Python 3.11）
- **前端**：HTML5 + Bootstrap 5.3 + Vanilla JS
- **資料來源**：SSRS（ERP 報表）、PDM Vault（SolidWorks）、ServCloud API（設備）、ERP SQL Server（Computech YC01）、Google 試算表（管理頁：請假/出勤/加班/報廢/生產日報/分類對照）
- **本地 DB**：SQLite（pdm_search.db 圖面索引含 dcn_index 表、zume_drawings.db ZUMEN 圖面對照表）
- **部署**：Python Embeddable Package，免安裝，double-click VBScript 啟動

**第三方 JS/CSS 函式庫一律下載到本地 `static/vendor/`，禁止用外部 CDN 連結**（如 `cdn.jsdelivr.net`）。原因：部署環境是內網電腦，有些電腦沒有外網連線，CDN 載入失敗時 Bootstrap 的 `.modal{display:none}` 規則不生效，會讓彈窗內容（含寫死的「搜尋中...」loading 文字）直接裸露顯示在頁面上卡住，且 JS 失敗導致按鈕也沒反應——這是實際發生過的真實故障（2026-06 修正）。目前 `static/vendor/` 已有 `bootstrap.min.css`、`bootstrap.bundle.min.js`、`JsBarcode.all.min.js`；新增任何前端函式庫都要照同樣方式下載進來，模板裡只能用 `/static/vendor/xxx` 引用，不可寫 `https://cdn...`。

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
| `templates/batch_cost.html` | 批成本計算（依製令計算刀具+加工成本，含明細/統計/刀表與刀具資料維護） |
| `templates/management.html` | 管理頁（請假/出勤/加班/申請單/治檢具/DCN/ZUMEN圖面/報廢/生產日報P5.3/生產報工P2） |

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

## 延伸文件（docs/）

以下細節文件**不會自動載入**，只有在實際碰到對應功能時才需要用 Read 工具去查：

| 文件 | 何時要讀 |
|------|------|
| `docs/pdm-index.md` | 修改 PDM 圖面索引（build_pdm_index.py）相關邏輯 |
| `docs/batch-cost.md` | 修改批成本計算（batch_cost.html / `/api/batch_cost/*`） |
| `docs/management-page.md` | 修改管理頁面（management.html）任何子頁 |
| `docs/category-colors.md` | 新增或修改「依分類」呈現的圖表配色 |
| `docs/table-design.md` | 新增明細表格，需要套用一致樣式 |
| `docs/badge-filter-convention.md` | 新增或修改表格內彩色徽章的篩選互動 |
| `docs/dcn-index.md` | 修改設計變更通知單（build_dcn_index.py）相關邏輯 |
| `docs/zumen.md` | 修改 ZUMEN 圖面子頁或 `/api/zume/*` |
| `docs/nav-design.md` | 新增頂層分頁或修改導覽列樣式 |
| `docs/dev-workflow.md` | 要 push GitHub，或用 preview 工具驗證前端改動時 |

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

## 字體規範

全站統一採用 **Source Han Sans TC（思源黑體）**，以 Google Fonts 的 **Noto Sans TC** 名稱引入（同一字型，不同授權方命名）。此為強制規定，禁止改用宋體或其他字型。

```html
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap" rel="stylesheet">
```
```css
body, * { font-family: 'Noto Sans TC', 'Source Han Sans TC', sans-serif; }
```

- 字重對應：`400` 正常／`500` 中等／`600` 加粗／`700` 粗體
- 禁止使用 `Noto Serif TC`、`Microsoft JhengHei`、`Segoe UI` 作為主字型；通用備援字型用 `sans-serif`（非 `serif`）；Adobe 名稱「Source Han Sans TC」可作第二備援
- **emoji 例外**：font-family 尾端可加 `'Segoe UI Emoji','Apple Color Emoji','Noto Color Emoji'` 作為 emoji 彩色顯示備援（只在遇到 emoji 字元時生效）。management.html 已套用，避免選單圖示在無彩色 emoji 字型的機器上顯示成怪方框。

## 語言規則

與使用者溝通（說明、進度、總結、錯誤訊息）一律使用繁體中文，不可使用英文或簡體中文。程式碼中的變數/函式/檔名等技術識別字維持英文，但註解一律繁體中文。
