# 加工部查詢系統 PDS（製令查詢）

詠基公司加工課用的製令查詢 Web 應用。Flask 後端 + Bootstrap 5 前端，以 Edge/Chrome --app 模式呈現為桌面視窗。

## 技術架構

- **後端**：Flask（Python 3.11）
- **前端**：HTML5 + Bootstrap 5.3 + Vanilla JS
- **資料來源**：SSRS（ERP 報表）、PDM Vault（SolidWorks）、ServCloud API（設備）、ERP SQL Server（Computech YC01）、Google 試算表（管理頁：請假/出勤/加班/報廢/生產日報/分類對照）
- **本地 DB**：SQLite（pdm_search.db 圖面索引含 dcn_index 表、zume_drawings.db ZUMEN 圖面對照表）
- **部署**：Python Embeddable Package，免安裝，double-click VBScript 啟動

**第三方 JS/CSS 函式庫一律下載到本地 `static/vendor/`，禁止用外部 CDN 連結**（如 `cdn.jsdelivr.net`）。原因：部署環境是內網電腦，有些電腦沒有外網連線，CDN 載入失敗時 Bootstrap 的 `.modal{display:none}` 規則不生效，會讓彈窗內容（含寫死的「搜尋中...」loading 文字）直接裸露顯示在頁面上卡住，且 JS 失敗導致按鈕也沒反應——这是實際發生過的真實故障（2026-06 修正）。目前 `static/vendor/` 已有 `bootstrap.min.css`、`bootstrap.bundle.min.js`、`JsBarcode.all.min.js`；新增任何前端函式庫都要照同樣方式下載進來，模板裡只能用 `/static/vendor/xxx` 引用，不可寫 `https://cdn...`。

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

## PDM 索引（build_pdm_index.py）

- 連線 SolidWorks MAXCLAW Vault（COM API）
- 品名/發圖日期優先從 PDM SQL Server 直連讀取（ADO via win32com）
- SQL 失敗時備援使用 COM IEdmEnumeratorVariable10 多策略讀取
- 掃描階段只存路徑字串，不保留 COM 物件（避免狀態污染）
- 索引存於 `%LOCALAPPDATA%\PDMSearch\pdm_search.db`

## 批成本計算（batch_cost.html）

頂層導覽列「批成本」分頁（位於「產品途程查詢」與「管理」之間）。依製令計算單批的刀具耗用成本＋加工費用，左側選單分四個子頁：

- **批成本計算**：①輸入製令查詢（支援不打 `-`，純數字會自動補回 4位單位代碼+`-`+序號格式）→ 自動帶出品號/品名/製程代號/製程名稱/報工人員/加工秒數/機台名稱代號（依「製令+製程代號」比對 K1_P2 生產報工統計與生產日報表P5.3）→ ②填完成數量/每秒鐘生產費用 → ③自動帶出該品號刀表，沒有則可現場新增/編輯 → ④逐刀填使用次數即時試算成本 → ⑤確認送出。
- **批成本明細**：讀『批成本計算』分頁全部記錄，支援搜尋+月份篩選+刪除；每列可展開查看該批次每把刀的使用明細（依「建立日期+製令+製程代號+品號」比對『製令與刀具壽命』分頁，並 join『刀具資料』補項目/供應商）；明細列依「建立日期」由新到舊排序。內含「統計」分頁：柱狀+折線組合圖（SVG 動態疊加折線於柱狀圖上）+平均值虛線參考線，月份成本趨勢、依品名平均加工成本比較。
- **刀表維護／刀具資料維護**：對『刀表』『刀具資料』分頁的完整 CRUD（新增/編輯/刪除），不透過製令查詢流程也能直接維護。

資料來源是共用區一份 Excel 範本（`config.py` 的 `BATCH_COST_FILE_PATH`），固定 4 個分頁：

| 分頁 | 設定常數 | 內容 |
|------|------|------|
| 刀具資料 | `BATCH_COST_TOOL_SHEET` | 刀具名稱→單價/刃數/項目/供應商對照 |
| 刀表 | `BATCH_COST_TOOLMAP_SHEET` | 品號→T1~T39 刀具配置（不分製程代號） |
| 批成本計算 | `BATCH_COST_RECORD_SHEET` | 每筆計算結果（不存在時自動建立） |
| 製令與刀具壽命 | `BATCH_COST_LIFESPAN_SHEET` | 每筆計算對應的逐刀使用次數明細（不存在時自動建立） |

成本公式：
```
刀具耗用成本（單把） = (使用次數 × 單價) / 刃數
刀具成本（單位） = 刀具費用總計 / 完成數量
加工費用（單位） = (加工秒數 / 完成數量) × 每秒鐘生產費用
加工費用(含刀具成本) = 加工費用 + 刀具成本
```

實作慣例：
- **批成本計算/製令與刀具壽命分頁一律依「表頭欄名」動態對應寫入**（`app.py` 讀第一列表頭建 `name→col` map，照名稱寫入，不假設固定欄位順序），因為這兩個分頁的欄位順序會隨需求調整；新增欄位時只要 value_map 加對應 key 即可，不用管欄位實際排第幾欄。
- **建立日期優先採用 K1_P2 該筆的「出站時間」**（實際生產完成日期），查不到才退回今天日期——不是單純記錄送出當下的日期。
- 製令查詢：優先查 K1_P2（含已完工製令），查不到才備援查 SSRS 未完工製令報表；P2 的「製令」欄位本身帶序號尾碼（如 `-0010`），比對生產日報表P5.3 機台名稱/代號時需先去尾碼。
- 寫檔沿用既有 `try/except (PermissionError, OSError)` fallback 慣例（檔案被 Excel 開著時改存桌面並回傳 `warning` 訊息），見 `_save_batchcost_wb()`。
- 刀號輸入框支援省略 `T`（純數字 blur 時自動補上）與按 ↓ 鍵跳下一列，見 `bcNormalizeSlotStr`/`bcSlotInputKeydown`。

## 管理頁面（management.html）

頂層導覽列「管理」分頁，左側選單分三群組（各以 📁 圖示 + 標題）：

- **人員勤務**：請假統計、出勤統計、加班統計、申請單
- **生產管理**：生產日報表P5.3、生產報工統計P2、報廢統計
- **技術文件**：治檢具清單、設計變更通知單、ZUMEN 圖面

左側選單 active 子頁為實心深藍底（`#1565C0`）+ 白字 + 左側深藍直條，醒目易辨識。

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

## 分類配色標準（A/B/C/L/未分類）

全站所有「依分類（A.CNC / B.鑽床 / C.沖床 / L.車床 / 未分類）」呈現的圖表，一律使用同一套配色，**不可各自定義新色票**：

```js
const CAT_COLOR = { A:'#3B82F6', B:'#22C55E', C:'#F59E0B', L:'#A855F7', other:'#6B7280' };
// A 藍／B 綠／C 橙／L 紫／other(未分類) 灰
```

- 基準定義在 `management.html` 的 `K1P2_CAT_COLORS`（生產報工統計P2／生產日報表P5.3 的分類圖表使用）；`CAT_COLOR`（報廢統計用）、`EMP_GROUP_COLORS`（員工名冊組別徽章用）皆須對齊此色票。
- 涵蓋範圍：分類佔比橫條（`.sc-cat-row-fill`）、月份堆疊長條圖（`scStackedVbar`）、月份彙整橫向堆疊圖（`mkStackedBar`）、KPI 卡片強調色等「圖表」類呈現。
- `.sc-cat-badge .sc-cat-A/B/C/L`（分類佔比文字徽章，淺底深字）已是對應的 Tailwind 100/700 色階，與上述標準同色系，新增分類圖表時可直接沿用。
- 不在此規範內：`K1P2_MACARON`／`K1P2_MACARON_DK`（P2 表格內馬卡龍漸層色，刻意做柔和變化的特殊樣式）、`.cat-badge.cat-A/B/C/L`（P5.3/P2 明細表格內 分類欄的小圓標，屬表格欄位樣式而非圖表）。

## 資料表格設計規範

管理頁面（management.html）明細表格的標準樣式，以「生產日報表P5.3」表格（`#prod-content` 範圍內）為基準，新增明細表格時應沿用：

### 卡片標題列
- class：`.card-header.today-leave-header`
- `font-size:0.95rem；font-weight:800；padding:0.7rem 1rem`
- `background:linear-gradient(90deg, var(--blue-mid) 0%, var(--blue-d) 100%)`（`#1976D2 → #0D47A1`）＋ `border-left:4px solid var(--blue-d)`，文字白色

### 篩選列
- class：`.filter-bar`，`padding:0.5rem 0.9rem`
- 內部 `select`/`input` 字級 `font-size:0.82rem；padding:0.3rem 0.5rem`
- 各篩選元件用 inline `width` 各自設定（例如日期框 `130px`、搜尋框 `180px`），無統一固定寬

### 表格本體
- 基底 class：`table.pro-table`；P5.3 透過 `#prod-content` 前綴覆寫加強對比與字級
- `<th>`：`padding:0.5rem 0.75rem；font-size:.84rem（基底.78rem）；font-weight:600；text-align:left（數字欄用 .center 置中）`；底色 `linear-gradient(180deg, var(--blue-mid), var(--blue-d))`，文字 `rgba(255,255,255,.95)`；`position:sticky;top:0;z-index:2` 做表頭固定
- `<td>`：`padding:0.45rem 0.75rem；font-size:.93rem；color:var(--text-2)`；`tr:nth-child(even) td` 底色 `var(--row-even)`(`#FAF9F5`)；`tr:hover td` 底色 `rgba(21,101,192,.05)` + 文字轉 `var(--text)`
- 容器：`.table-wrap { flex:1; overflow:auto }`，不設 `max-height`（隨版面自動撐滿，搭配 sticky `<th>` 卷動時表頭不跑掉）

### 欄寬（inline px，依內容類型分配，非等寬）
窄欄（日期/人員/製程/數字類，52~90px）：生產日期88px、人員90px、製程75px、生產數62px、秒數62px、實際秒數70px、標工52px、機台代號68px、分類80px；中等欄（120~170px）：製令120px、品號170px、機台名稱90px；唯一不限寬欄：品名 `min-width:200px`（隨內容延展）。

### 數字強調徽章（如「實際秒數」欄）
- 命名模式：`td.{欄位簡稱} span`（例：`td.pr-actsec span`）
- `padding:.1rem .5rem；border-radius:10px；font-weight:700；font-size:.85rem；background:#FFF3E0；color:#E65100`（橘色系，用於提醒「實際秒數」等需要特別注意的換算值）

### 分類圓標（表格欄位內，非圖表）
- class：`.cat-badge` + `.cat-A/.cat-B/.cat-C/.cat-L`
- 基底：`padding:.15rem .45rem；border-radius:4px；font-size:.82rem；font-weight:700`
- 配色：A `#E3F2FD/#1565C0`、B `#E8F5E9/#2E7D32`、C `#FFF3E0/#E65100`、L `#F3E5F5/#6A1B9A`
- **注意**：這套配色跟「分類配色標準」章節的 `CAT_COLOR`／`K1P2_CAT_COLORS` 是兩套不同色票（`.cat-badge` 專屬表格欄位小圓標，不可互相套用，已在前一章節明確排除）。

## 徽章點擊篩選慣例（單擊篩選／雙擊清除）

管理頁面（management.html）明細表格中，任何「狀態」「分類」「會計科目」等彩色徽章，只要該欄位在篩選列已有對應的下拉選單（`<select>`），就必須支援：

- **單擊徽章** → 把該下拉選單的值設成這個徽章的值，並重新渲染（等同直接點下拉選單選同一個選項）
- **雙擊徽章** → 把該下拉選單的值清空（設回 `''`，即「全部」），並重新渲染

實作方式固定一個模式：寫一個 `xxxFilterByY(value)` 函式（設值＋重新渲染），徽章的 `onclick` 呼叫它帶實際值，`ondblclick` 呼叫**同一個函式**帶空字串 `''` 即可清除，不需要另外寫一個清除用的函式。徽章本身要加 `style="cursor:pointer"` 與 `title="單擊篩選／雙擊清除"` 提示。

已套用的位置：
- 採購登入表：狀態徽章 → `puFilterByStatus()`／會計科目徽章 → `puFilterByAccount()`
- 報廢統計：分類徽章 → `scFilterByCat()`
- 生產日報表P5.3：分類徽章 → `prFilterByCat()`
- 生產報工統計P2：分類徽章 → `p2FilterByCat()`

**新增任何帶篩選下拉選單的彩色徽章時，一律比照此慣例加上單擊／雙擊事件**，不可只做單擊。若該欄位目前還沒有對應的篩選下拉選單（例如治檢具清單/DCN/員工名冊的狀態欄），要先補上下拉選單再套用此慣例，不可省略下拉選單只做徽章互動。

## 設計變更通知單索引（build_dcn_index.py）

- 掃描 PDM Vault『00-研發部\02-文件資料\02-設計變更通知單』子資料夾，下載 xlsm 申請表。
- **PDM Card 變數值讀取自 Office 檔的 `docProps/custom.xml`**（Office Custom Properties），不是 worksheet cell——RR 新格式（2021+）的 cell 是空的，唯一可靠來源是 custom.xml。
- 讀取欄位：機型、提出人員、經辦、審核/核決、申請原因（PP_R_004_*_tasky 勾選項）、申請變更內容說明、工作流程狀態。
- 寫入 `pdm_search.db` 的 `dcn_index` 表。前端可點列開卡片彈窗、篩選提出人員、顯示已作廢/已結案。

## ZUMEN 圖面（管理頁子頁，唯讀檢視）

- ZUMEN（`zume-n.com`）是**第三方雲端圖面系統**，本專案僅做**唯讀整合**。
- 資料來源：使用者從 ZUMEN 匯出的 `zume-n_data_list_*.csv`，放到「下載」資料夾後啟動時自動匯入（`_auto_import_zume_csv`），或前端按「重新匯入」（`/api/zume/scan`）。
- 存於 `zume_drawings.db` 的 `drawings` 表：`part_no`(圖號)、`part_name`(品名)、`url`、`line`(生產線別)、`prod_group`(生產群組)、`category`(分類)、`vendor`(廠商)。後四欄以 **CSV 標題關鍵字動態偵測**（`_zume_header_indices`），CSV 沒有就留空、前端自動隱藏空欄。
- API：`/api/zume/list`（清單+篩選選項）、`/api/zume/open`（開啟圖號對應 ZUMEN 頁）、`/api/zume/lookup`、`/api/zume/scan`、`/api/zume/import`。
- **不做寫入/上傳**：ZUMEN 無官方 API。曾評估其內部 API（Next.js + `https://zume-n.com/api`，Auth0 Bearer token，GCS signed-url 上傳），技術上可行但需存帳密自動登入＋維護＋條款風險，**已決定不實作**；「快速新增/草稿」功能也已移除。

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
- **emoji 例外**：font-family 尾端可加 `'Segoe UI Emoji','Apple Color Emoji','Noto Color Emoji'` 作為 emoji 彩色顯示備援（只在遇到 emoji 字元時生效，不影響中文/英文文字字型）。management.html 已套用，避免選單圖示在無彩色 emoji 字型的機器上顯示成怪方框。

## 語言規則

與使用者溝通（說明、進度、總結、錯誤訊息）一律使用繁體中文，不可使用英文或簡體中文。程式碼中的變數/函式/檔名等技術識別字維持英文，但註解一律繁體中文。
