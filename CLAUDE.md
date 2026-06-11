# 加工部查詢系統 PDS（製令查詢）

詠基公司加工課用的製令查詢 Web 應用。Flask 後端 + Bootstrap 5 前端，以 Edge/Chrome --app 模式呈現為桌面視窗。

## 技術架構

- **後端**：Flask（Python 3.11）
- **前端**：HTML5 + Bootstrap 5.3 + Vanilla JS
- **資料來源**：SSRS（ERP 報表）、PDM Vault（SolidWorks）、ServCloud API（設備）、ERP SQL Server（Computech YC01）
- **本地 DB**：SQLite（pdm_search.db 圖面索引、zume_drawings.db 技術規格）
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
```

## 重要檔案

| 檔案 | 說明 |
|------|------|
| `app.py` | Flask 主程式，所有 API 端點與商業邏輯 |
| `main.py` | GUI 啟動器，開 Edge/Chrome --app 模式視窗 |
| `config.py` | 外部設定（IP、port、路徑） |
| `build_pdm_index.py` | PDM 圖面索引重建工具（v1.2，SQL Server 直連） |
| `build_embed.py` | 打包成可攜式版本 |
| `sync_to_dist.py` | 快速同步修改到 dist_embed |
| `templates/index.html` | 主畫面（製令搜尋） |
| `templates/drawing.html` | 圖面查詢 |
| `templates/bom.html` | BOM 查詢 |
| `templates/routing.html` | 途程查詢 |
| `templates/equipment.html` | 設備稼動查詢 |
| `templates/production.html` | 生產日報 |
| `templates/print_report.html` | 列印工單 |

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

## 搜尋語法

製令搜尋支援 AND/NOT：
- 空格 = AND（例：`FTB 夾管座`）
- `-前綴` = NOT（例：`FTB -舊版`）

## TTL 快取

SSRS 查詢有 120 秒 TTL 快取（`cache_get/cache_set`），避免頻繁呼叫。

## 語言規則

與使用者溝通（說明、進度、總結、錯誤訊息）一律使用繁體中文，不可使用英文或簡體中文。程式碼中的變數/函式/檔名等技術識別字維持英文，但註解一律繁體中文。
