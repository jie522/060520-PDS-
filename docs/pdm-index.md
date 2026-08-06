# PDM 索引（build_pdm_index.py）

> 何時需要讀這份文件：修改 `build_pdm_index.py` 或排查 PDM 圖面索引（pdm_search.db）資料不準確的問題時。

- 連線 SolidWorks MAXCLAW Vault（COM API）
- 品名/發圖日期優先從 PDM SQL Server 直連讀取（ADO via win32com）
- SQL 失敗時備援使用 COM IEdmEnumeratorVariable10 多策略讀取
- 掃描階段只存路徑字串，不保留 COM 物件（避免狀態污染）
- 索引存於 `%LOCALAPPDATA%\PDMSearch\pdm_search.db`

## 踩雷記錄

- 2026-07-05｜症狀：TCW 系列圖面在 PDM圖面查詢頁的機型/型號全部顯示成「TCW」
  根因：機型原本只靠資料夾路徑推導（`derive_ji_xing`），TCW 系列圖檔直接放在系列資料夾下、
  沒有各機型子資料夾，fallback 變成資料夾名稱本身
  修法：「機型」加入 TARGET_VARS 改從 PDM 資料卡讀取（SQL 直查，與品名同批），
  路徑推導只當卡片沒填值時的備援
  注意：`--update` 增量模式只補「品名/發圖日期為空」的列，修欄位值錯誤要跑完整重建
- 2026-07-05｜症狀：機型修好後，搜尋「TCW030180」仍查無資料
  根因：`/api/drawing/search` 只比對品號/品名/檔案路徑，沒比對 ji_xing 欄位
  （以前搜機型有結果是碰巧路徑含機型資料夾名，TCW 系列無此結構）
  修法：SQL LIKE 初篩與 match_and_not 的 combined 字串都加入 ji_xing
