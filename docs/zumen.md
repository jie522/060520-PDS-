# ZUMEN 圖面（管理頁子頁，唯讀檢視）

> 何時需要讀這份文件：修改管理頁面「ZUMEN 圖面」子頁或 `/api/zume/*` 路由時。

- ZUMEN（`zume-n.com`）是**第三方雲端圖面系統**，本專案僅做**唯讀整合**。
- 資料來源：使用者從 ZUMEN 匯出的 `zume-n_data_list_*.csv`，放到「下載」資料夾後啟動時自動匯入（`_auto_import_zume_csv`），或前端按「重新匯入」（`/api/zume/scan`）。
- 存於 `zume_drawings.db` 的 `drawings` 表：`part_no`(圖號)、`part_name`(品名)、`url`、`line`(生產線別)、`prod_group`(生產群組)、`category`(分類)、`vendor`(廠商)。後四欄以 **CSV 標題關鍵字動態偵測**（`_zume_header_indices`），CSV 沒有就留空、前端自動隱藏空欄。
- API：`/api/zume/list`（清單+篩選選項）、`/api/zume/open`（開啟圖號對應 ZUMEN 頁）、`/api/zume/lookup`、`/api/zume/scan`、`/api/zume/import`。
- **不做寫入/上傳**：ZUMEN 無官方 API。曾評估其內部 API（Next.js + `https://zume-n.com/api`，Auth0 Bearer token，GCS signed-url 上傳），技術上可行但需存帳密自動登入＋維護＋條款風險，**已決定不實作**；「快速新增/草稿」功能也已移除。
