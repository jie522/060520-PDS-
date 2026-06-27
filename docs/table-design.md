# 資料表格設計規範

> 何時需要讀這份文件：在管理頁面或其他頁面新增明細表格時，需要套用一致的視覺樣式。

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
- **注意**：這套配色跟 `docs/category-colors.md` 的 `CAT_COLOR`／`K1P2_CAT_COLORS` 是兩套不同色票（`.cat-badge` 專屬表格欄位小圓標，不可互相套用）。
