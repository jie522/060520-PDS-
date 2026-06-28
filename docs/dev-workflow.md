# 開發與部署操作慣例

> 何時需要讀這份文件：要 push 到 GitHub，或要用 preview 工具（screenshot/inspect/click）視覺驗證前端改動時。

## GitHub

遠端：`https://github.com/jie522/060520-PDS-.git`，直接 push 到 `master` 分支（單人開發，沒有 PR 流程）。

使用者說「PUSH GITHUB」時，代表要把目前所有未 commit 的變更整理成 commit 並 push 到 `origin/master`：

1. `git status` 確認範圍，排除明顯跟目前工作無關的未追蹤檔案（例如使用者自己留的壓縮檔）
2. 寫繁體中文 commit message（風格參考 `git log`，條列式描述做了什麼，而非怎麼做的）
3. commit → push

## preview 工具與 5088 埠衝突

這個專案的 Flask 預設埠是 5088（`config.py` 的 `FLASK_PORT`，`.claude/launch.json` 的 `"port"` 也設定同一個值）。使用者常常自己開著桌面應用（雙擊 vbs 啟動的視窗）在用，這時 5088 會被佔用，`preview_start` 會跟它衝突。

需要用 preview 工具視覺驗證前端改動時：

1. 先 `netstat -ano | grep ":5088"` 確認是否真的被佔用
2. 若被佔用，暫時把 `config.py` 的 `FLASK_PORT` 跟 `.claude/launch.json` 的 `"port"` 都改成 5099
3. `preview_start` → 驗證
4. `preview_stop`，把兩個檔案的埠**改回 5088**，用 `git diff config.py .claude/launch.json` 確認沒有殘留變更才算清乾淨——忘記還原會讓使用者下次啟動桌面應用失敗或行為不一致

## 寫入共用 Excel 檔案的功能要先用本地副本測試

任何會寫入 `\\192.168.1.99\共用區\...` 共用 Excel 檔案的功能（目前主要是批成本計算），測試方式見 `docs/batch-cost.md`「共用 Excel 範本的結構會持續變動」一節——核心原則：絕對不要直接對正式共用檔案做寫入測試。
