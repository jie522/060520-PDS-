# ZUMEN 圖面（管理頁子頁，唯讀檢視）

> 何時需要讀這份文件：修改管理頁面「ZUMEN 圖面」子頁或 `/api/zume/*` 路由時。

- ZUMEN（`zume-n.com`）是**第三方雲端圖面系統**，本專案僅做**唯讀整合**。
- 資料來源：使用者從 ZUMEN 匯出的 `zume-n_data_list_*.csv`，放到「下載」資料夾後啟動時自動匯入（`_auto_import_zume_csv`），或前端按「重新匯入」（`/api/zume/scan`）。
- 存於 `zume_drawings.db` 的 `drawings` 表：`part_no`(圖號)、`part_name`(品名)、`url`、`line`(生產線別)、`prod_group`(生產群組)、`category`(分類)、`vendor`(廠商)。後四欄以 **CSV 標題關鍵字動態偵測**（`_zume_header_indices`），CSV 沒有就留空、前端自動隱藏空欄。
- API：`/api/zume/list`（清單+篩選選項）、`/api/zume/open`（開啟圖號對應 ZUMEN 頁）、`/api/zume/lookup`、`/api/zume/scan`、`/api/zume/import`。
- **不做寫入/上傳**：ZUMEN 無官方 API。曾評估其內部 API（Next.js + `https://zume-n.com/api`，Auth0 Bearer token，GCS signed-url 上傳），技術上可行但需存帳密自動登入＋維護＋條款風險，**已決定不實作**；「快速新增/草稿」功能也已移除。

## 線上圖面管理 GUI 整合（2026-07-05）

管理頁 ZUMEN 子頁的「🖼 線上圖面管理」按鈕會啟動並開啟本機的 260704-zumen Node 工具
（https://github.com/jie522/260704-zumen ，Express+Playwright 自動登入 zume-n.com）。

- 設定：`config.py` 的 `ZUMEN_GUI_DIR`（本機安裝目錄）與 `ZUMEN_GUI_PORT`（3456）
- API：`/api/zumen_gui/status`（installed/has_env/running）、`/api/zumen_gui/start`（背景啟動 node server.js）
- 這是「選配」：沒裝 Node/工具的電腦按鈕會提示未安裝，不影響其他功能
- 每台電腦要自己準備 `.env`（ZUMEN_EMAIL/ZUMEN_PASSWORD，不在 git 裡）
- 2026-07-05：ZUMEN GUI 詳情頁底部加「PDS 整合卡」——以圖號向 PDS（port 5088）查
  PDM 圖面（`/api/drawing/search`，可 SW 開啟）與 ERP 途程（`/api/bom/routing`，製程代號表）。
  app.py 的 after_request 對 localhost:3456 開了 CORS；改動在 260704-zumen 專案的 public/index.html。
- 2026-07-05：ZUMEN 改為頂層導覽分頁 `/zumen`（templates/zumen.html，淺藍色 tab，
  全頁 iframe＋自動啟動 Node 工具），管理頁的 ZUMEN 子頁已移除。
  ZUMEN 建立圖面流程順序＝建立→附範本→寫基本資料（ZUMEN 對剛建立的圖面
  PATCH 會回 200 但不寫入，需等約 1 分鐘，範本附掛正好當等待）。
- 2026-07-05｜症狀：ZUMEN 分頁點下去畫面整個空白（連載入中文字都看不到）
  根因：`zumenShowFrame()` 把 `frame.style.display = ''`（空字串），但 CSS 已宣告
  `#zumen-frame{display:none}`；inline style 設空字串不會覆蓋樣式表規則，iframe 永遠顯示不出來
  修法：明確設成 `frame.style.display = 'flex'`
  排錯方式：preview 工具在此環境對本頁 screenshot/navigate 持續逾時（連根目錄 / 也一樣，
  屬環境問題非本頁 bug），改用專案既有 Playwright（`260704 ZUMEN圖面管理` 的 node_modules，
  channel:'msedge'）寫獨立腳本 render + getComputedStyle 直接檢查才定位到問題
