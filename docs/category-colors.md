# 分類配色標準（A/B/C/L/未分類）

> 何時需要讀這份文件：新增或修改任何「依分類」呈現的圖表（柱狀圖、堆疊圖、KPI 卡片強調色等）時。

全站所有「依分類（A.CNC / B.鑽床 / C.沖床 / L.車床 / 未分類）」呈現的圖表，一律使用同一套配色，**不可各自定義新色票**：

```js
const CAT_COLOR = { A:'#3B82F6', B:'#22C55E', C:'#F59E0B', L:'#A855F7', other:'#6B7280' };
// A 藍／B 綠／C 橙／L 紫／other(未分類) 灰
```

- 基準定義在 `management.html` 的 `K1P2_CAT_COLORS`（生產報工統計P2／生產日報表P5.3 的分類圖表使用）；`CAT_COLOR`（報廢統計用）、`EMP_GROUP_COLORS`（員工名冊組別徽章用）皆須對齊此色票。
- 涵蓋範圍：分類佔比橫條（`.sc-cat-row-fill`）、月份堆疊長條圖（`scStackedVbar`）、月份彙整橫向堆疊圖（`mkStackedBar`）、KPI 卡片強調色等「圖表」類呈現。
- `.sc-cat-badge .sc-cat-A/B/C/L`（分類佔比文字徽章，淺底深字）已是對應的 Tailwind 100/700 色階，與上述標準同色系，新增分類圖表時可直接沿用。
- 不在此規範內：`K1P2_MACARON`／`K1P2_MACARON_DK`（P2 表格內馬卡龍漸層色，刻意做柔和變化的特殊樣式）、`.cat-badge.cat-A/B/C/L`（P5.3/P2 明細表格內 分類欄的小圓標，屬表格欄位樣式而非圖表，配色定義見 `docs/table-design.md`）。
