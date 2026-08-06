# 加工部查詢系統 PDS（製令查詢）

詠基公司加工課用的製令查詢 Web 應用。Flask（Python 3.11）+ Bootstrap 5 + Vanilla JS，
以 Edge/Chrome --app 模式呈現為桌面視窗。部署為 Python Embeddable 免安裝包（dist_embed/）。

**每個新 session 先讀 `docs/agent/diagnosis.md`（3 分鐘，避開本環境三大坑）。**

## 資料來源與關鍵設定（config.py）

- SSRS 報表 `http://192.168.1.212/ReportServer`｜ERP SQL Server `192.168.1.140` / DB `YC01`
- PDM Vault：MAXCLAW（SolidWorks PDM 2021，COM API）｜ServCloud API（設備）｜Google 試算表（管理頁各統計）
- 本地 SQLite：pdm_search.db（圖面/治檢具/DCN 索引）、zume_drawings.db
- FLASK_PORT = 5088

## 常用指令

```bash
python app.py                        # 開發模式啟動
python sync_to_dist.py               # 快速同步修改到 dist_embed（檔案被鎖時改用 Copy-Item 逐檔複製）
python build_embed.py                # 重新打包可攜式發行版
python build_pdm_index.py --update --deploy   # PDM 圖面索引增量更新
python build_jig_index.py --deploy   # 治檢具索引重建
python build_dcn_index.py --deploy   # DCN 索引重建
python build_cnc_program_index.py --deploy    # CNC 程式索引重建
python build_equipment_index.py --deploy      # 設備主檔匯入（Excel→equipment.db）
```

## 重要檔案

| 檔案 | 說明 |
|------|------|
| `app.py` | Flask 主程式，所有 API 端點與商業邏輯 |
| `main.py` / `config.py` | GUI 啟動器 / 外部設定（IP、port、路徑） |
| `templates/index.html` | 製令搜尋主畫面（搜尋語法：空格=AND、`-`前綴=NOT） |
| `templates/management.html` | 管理頁（多個子頁，改前必讀 `docs/management-page.md`） |
| `templates/batch_cost.html` | 批成本計算（改前必讀 `docs/batch-cost.md`） |
| `templates/cnc_program.html` | CNC 程式管理 |
| 其他 templates | drawing / bom / routing / equipment / production / print_report |
| `build_*.py` | 各索引重建工具 |

## dist_embed 同步鐵則（最常出錯，背下來）

```
dist_embed/PDS系統/
├── _python/    # Python 3.11.9 Embeddable（含 site-packages：pywin32、openpyxl、Pillow…）
├── _app/       # 程式碼副本 ← 改完的檔案要同步到這裡
├── config.py   # ★使用者可編輯設定，執行時載入的是這份 ← config.py 有改必須「兩份」都同步
└── PDS系統.vbs
```
1. 改 .py 後先 `python -m py_compile` 再同步。
2. 「user-mapped section open」= dist 版程式執行中鎖檔，先處理程序再同步。
3. 每天第一次完成更新時：`app.py` 的 `APP_VERSION` 改為當日 `VYYYYMMDD`（同日多次更新不用重改），同步後請使用者重啟測試。
4. **`equipment.db`／`cnc_program_index.db`／`calendar.db` 不在 `sync_to_dist.py` 的同步清單裡**——
   這三個是執行期資料庫，桌面應用執行中會直接寫入（照片上傳、設備編輯、CNC 程式上傳、
   行事曆工作日設定、維修停機時數）。曾經跟原始碼放一起無條件覆蓋，結果使用者在正式
   桌面應用上傳的照片被開發機的舊資料庫蓋掉（2026-07-30 真實故障，檔案還在網芳但資料庫
   記錄消失）。要推送前兩個，一律用各自明確的指令（`build_equipment_index.py --deploy` /
   `build_cnc_program_index.py --deploy`），不要手動 `shutil.copy2` 或改回加進 `sync_to_dist.py`。
   **`calendar.db` 更沒有任何 build 工具能重建**（純使用者手動點選的資料），蓋掉就救不回來。

## 延伸文件（不自動載入，碰到對應功能時用 Read 查）

| 文件 | 何時要讀 |
|------|------|
| `docs/agent/diagnosis.md` | ★每個新 session 開頭 |
| `docs/agent/dispatch.md` | 要派 subagent／選模型時 |
| `docs/agent/judgment.md` | 拿不準「該不該問／算不算完成／要不要換方向」時 |
| `docs/agent/prompts.md` | 寫 subagent 交辦 prompt 時 |
| `docs/agent/maintenance.md` | 要改制度檔／寫入踩雷教訓時 |
| `docs/pdm-api-cookbook.md` | ★動任何 PDM 讀寫之前（COM API 實測手冊） |
| `docs/pdm-jig-application-sop.md` | 要比照治檢具申請單的做法做其他 PDM 申請單自動化（換帳號/換申請單類型）時 |
| `docs/pdm-index.md` | 改 PDM 圖面索引（build_pdm_index.py） |
| `docs/batch-cost.md` | 改批成本計算 |
| `docs/management-page.md` | 改管理頁任何子頁 |
| `docs/category-colors.md` / `docs/chart-style.md` / `docs/table-design.md` / `docs/badge-filter-convention.md` | 分類圖表配色／趨勢折線圖樣式／新表格樣式／徽章篩選互動 |
| `docs/dcn-index.md` / `docs/cnc-program.md` / `docs/sfcr06.md` / `docs/zumen.md` | 對應功能 |
| `docs/equipment-master.md` | 改設備管理（設備主檔／編碼／照片／規格／妥善率） |
| `docs/calendar.md` | 改管理頁行事曆，或任何需要「工作天數」的統計 |
| `docs/nav-design.md` | 新增頂層分頁或改導覽列 |
| `docs/dev-workflow.md` | push GitHub、preview 工具驗證前端 |

## 硬規則（無例外）

- **禁止外部 CDN**：第三方 JS/CSS 一律下載到 `static/vendor/` 本地引用（內網電腦沒外網，
  CDN 失敗會讓 Bootstrap modal 裸露卡死——2026-06 真實故障）。
- **字體**：全站 Noto Sans TC（=思源黑體），字重 400/500/600/700；禁用宋體/微軟正黑/Segoe UI
  當主字型；備援 `sans-serif`；font-family 尾端可加 emoji 字型備援。
- **語言**：對使用者一律繁體中文；程式碼註解繁中；變數/函式/檔名英文。
- **日期**：前端日期字串用 `getFullYear()/getMonth()/getDate()` 手組，禁用 `toISOString()`（時區差一天）。
- SSRS 查詢有 120 秒 TTL 快取（`cache_get/cache_set`）。
