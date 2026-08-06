# 本環境快速診斷（每個新 session 先讀這份）

> 讀者：接手本專案的任何 AI session。寫於 2026-07-03（Fable 5 實戰一整天後的總結）。
> 這三條是本 harness 最燒 token、最易失焦、最易出錯的前三名，各附可照做的修法。

## 1. Token 最漏：直接 Read 整個大檔案

本專案的檔案很大：`templates/management.html` 約 6,600 行、`app.py` 約 4,800 行、
`templates/batch_cost.html` 約 1,500 行。整檔讀進主對話一次就吃掉大量 context。

**修法（照做即可）：**
1. 永遠先 `Grep`（pattern + `-C 3` + `head_limit`）鎖定行號，再 `Read` 該區段，單次 `limit` ≤ 120 行。
2. 「這功能在哪裡實作」這類跨檔問題 → 派 `Explore` subagent（model=haiku，breadth=quick 或 medium），主對話只收「檔案:行號＋一句結論」。
3. 例外：docs/ 底下的說明檔都很短（多數 <60 行），直接整檔 Read 沒關係。

## 2. 最易失焦：dist_embed 雙副本同步

repo 是源頭，`dist_embed/PDS系統/` 是部署副本。改完源頭忘記同步、或同步不完整，
使用者測到的是舊行為，來回鬼打牆。本專案實際踩過的三個坑：

- **根目錄 config.py 才是執行時載入的那份**：`dist_embed/PDS系統/config.py`（使用者可編輯版）
  會蓋過 `_app/config.py`。config.py 有新增常數時**兩份都要同步**，漏掉會 AttributeError。
- **程式執行中會鎖檔**：`Copy-Item` 報 "user-mapped section open" 就是 dist 版程式還在跑。
  先 `Get-Process | Where-Object { $_.Path -like "*dist_embed*" }` 找到 pythonw，請使用者關閉或（經同意）Stop-Process。
- **`sync_to_dist.py` 在檔案被鎖時會失敗**，此時改用 `Copy-Item` 逐檔複製沒被鎖的部分。

**修法（每次改完的固定收尾動作）：**
```
1. python -m py_compile app.py          # 語法檢查
2. 改 APP_VERSION 為當日 VYYYYMMDD      # 一天只需改一次
3. Copy-Item 同步所有改過的檔到 dist_embed/PDS系統/_app/
4. 若 config.py 有改 → 額外同步到 dist_embed/PDS系統/config.py（根目錄）
5. 回覆裡明講「請重啟程式再測」
```

## 3. 最易出錯：憑記憶寫外部系統的 API

本專案掛著四個外部系統：PDM（SolidWorks COM API）、ERP SQL Server、SSRS、Google 試算表。
憑訓練記憶寫這些 API 幾乎必錯。實際踩過：PDM COM 的 gen_py 預設參數 `0` 無法轉型（要明確傳
`None`）、卡片變數在 `docProps/custom.xml` 不在儲存格、workflow 分派靠 DB 變數而非檔案內值。

**修法（照做即可）：**
1. 動 PDM 前先讀 `docs/pdm-api-cookbook.md`（本次 session 驗證過的所有 API 用法都在裡面），
   cookbook 沒有的用法 → 先寫**最小唯讀測試腳本**跑通，再寫進 app.py。
2. 測試腳本一律寫成 UTF-8 檔案放 scratchpad 再執行，**不要用 bash heredoc 直接塞中文**
   （Windows 主控台 cp950 會 mojibake，且 heredoc 內的中文可能以錯誤編碼傳給 COM）。
3. 任何「寫入正式共用資源」的第一次（PDM vault、共用區 Excel）：先做唯讀驗證＋讀
   `C:\Users\user\.claude\projects\D--claude-code-04---SQL-SEARCH\memory\` 的相關教訓，
   並讓使用者知道你要寫入什麼。

## 附：兩個小但常見的坑

- **時區**：`new Date().toISOString()` 是 UTC，台灣（UTC+8）會差一天。日期字串一律用
  `getFullYear()/getMonth()/getDate()` 手組。
- **`import openpyxl` 第一次可能耗 30 秒**（防毒掃描），不是程式卡死，不要因此判定 hang。
  判定 hang 前先用 `faulthandler.dump_traceback_later()` 看卡在哪。
