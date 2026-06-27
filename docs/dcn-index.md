# 設計變更通知單索引（build_dcn_index.py）

> 何時需要讀這份文件：修改 `build_dcn_index.py` 或管理頁面「設計變更通知單」子頁時。

- 掃描 PDM Vault『00-研發部\02-文件資料\02-設計變更通知單』子資料夾，下載 xlsm 申請表。
- **PDM Card 變數值讀取自 Office 檔的 `docProps/custom.xml`**（Office Custom Properties），不是 worksheet cell——RR 新格式（2021+）的 cell 是空的，唯一可靠來源是 custom.xml。
- 讀取欄位：機型、提出人員、經辦、審核/核決、申請原因（PP_R_004_*_tasky 勾選項）、申請變更內容說明、工作流程狀態。
- 寫入 `pdm_search.db` 的 `dcn_index` 表。前端可點列開卡片彈窗、篩選提出人員、顯示已作廢/已結案。
