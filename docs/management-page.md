# 管理頁面（management.html）

> 何時需要讀這份文件：修改管理頁面（`templates/management.html`）裡任何子頁面時。

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
| 生產日報表P5.3 | `PROD_REPORT_SHEET_ID` + `PROD_REPORT_SHEET_NAME`（`'P5.3生產日報表data_ref'`，用分頁名稱取代 GID） | 明細與圖表同源；此分頁無機台名稱/備註 |
| 生產報工統計P2 | `CATEGORY_SHEET_ID`（K1_P2.ref，出站數量） | 與生產日報表是**不同**資料來源，數字本就不同 |
| 員工工號對照 | `EMPLOYEE_SHEET_ID` | 工號 → 姓名 |
| ABC 分類對照 | `CATEGORY_SHEET_ID` | key = 品號+製程代號 → A/B/C/L |

- Google 試算表以「知道連結的人可檢視」共用，後端用 gviz CSV 端點讀取（`fetch_google_sheet_csv`），無需金鑰。
- 生產日報/報工的分類：① 品號+製程對照表 → ② 機台名稱/代號/製程關鍵字備援（CNC→A、鑽→B、沖→C、車→L）。

相關延伸文件：分類配色標準見 `docs/category-colors.md`、表格樣式見 `docs/table-design.md`、徽章篩選慣例見 `docs/badge-filter-convention.md`、治檢具/PDM 索引見 `docs/pdm-index.md`、DCN 索引見 `docs/dcn-index.md`、ZUMEN 圖面見 `docs/zumen.md`。
