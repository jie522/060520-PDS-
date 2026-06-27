# PDM 索引（build_pdm_index.py）

> 何時需要讀這份文件：修改 `build_pdm_index.py` 或排查 PDM 圖面索引（pdm_search.db）資料不準確的問題時。

- 連線 SolidWorks MAXCLAW Vault（COM API）
- 品名/發圖日期優先從 PDM SQL Server 直連讀取（ADO via win32com）
- SQL 失敗時備援使用 COM IEdmEnumeratorVariable10 多策略讀取
- 掃描階段只存路徑字串，不保留 COM 物件（避免狀態污染）
- 索引存於 `%LOCALAPPDATA%\PDMSearch\pdm_search.db`
