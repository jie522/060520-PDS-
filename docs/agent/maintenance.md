# 制度維護協議（哪些檔能自己改、教訓寫哪裡、多長要精簡）

> 讀者：本專案的任何 AI session。寫於 2026-07-03。

## 1. 改檔權限分級

| 等級 | 檔案 | 規則 |
|---|---|---|
| 自由改 | `docs/` 全部（含 `docs/agent/`）、記憶目錄 | 發現錯誤直接修；新教訓照 §2 追加。改制度檔（dispatch/judgment/prompts/maintenance）時在檔尾「變更記錄」加一行 |
| 謹慎改 | `CLAUDE.md` | 只能：修錯誤、加索引行、精簡。上限 150 行，超過就把內容抽到 docs/ 只留索引。改前備份到 `docs/agent/CLAUDE.md.bak-YYYYMMDD` |
| 先問使用者 | `config.py` 的 IP/路徑/密碼類設定、`.claude/settings*.json`、`build_embed.py`、刪除任何 docs/ 檔案 | 影響部署或安全，先說明動機取得同意 |
| 不碰 | `dist_embed/_python/`（打包環境）、PDM vault 與共用區裡「非本 session 建立」的東西 | 出問題請使用者或管理員處理 |

## 2. 踩雷教訓寫回哪裡、什麼格式

**原則：教訓寫在「下次會踩到的人必經的路上」，不是集中一大本。**

- 功能相關教訓 → 對應的 `docs/*.md`（例：PDM 的坑寫進 `pdm-api-cookbook.md`，
  管理頁的坑寫進 `management-page.md`）。檔內加「## 踩雷記錄」節（沒有就建）。
- 跨功能的環境坑（同步、編碼、工具行為）→ `docs/agent/diagnosis.md`。
- 使用者偏好、協作方式 → 記憶目錄（`C:\Users\user\.claude\projects\D--claude-code-04---SQL-SEARCH\memory\`），
  照該處 frontmatter 格式。

**格式（一筆 2–4 行，必含四件事）：**
```
- 2026-07-03｜症狀：AddFolder 丟 "Python instance can not be converted to a COM object"
  根因：gen_py 對 optional COM 參數的預設值 0 無法轉型
  修法：明確傳 None，例 AddFolder(0, name, None)
  （可選）驗證方式：pdm_test6.py 型的最小腳本
```
寫之前先搜尋該檔是否已有同根因的記錄——有就補充原記錄，不要重複開條目。

## 3. 精簡時機與方式

- 單一 docs 檔 >150 行，或「踩雷記錄」>15 筆 → 該次 session 順手精簡：
  合併同根因條目、刪已被程式碼修死（不可能再踩）的條目、過時內容整段刪除。
- `CLAUDE.md` 直接引用的常載內容合計勿超過 500 行。實際行數用
  `wc -l CLAUDE.md docs/agent/diagnosis.md` 現查，勿依賴文件裡寫死的數字。
- 精簡是刪減不是改寫：拿不準要不要刪的條目保留並標 `（存疑 YYYY-MM-DD）`，
  下次再看到且仍無用就刪。

## 4. 制度本身失效的處理

- 發現制度檔與現實矛盾（例：dispatch.md 寫的 model 參數值在當下環境不存在）→
  **以當下系統提示為準**，順手修正制度檔並在變更記錄註明。
- 同一條規則連續兩個 session 都被迫違反 → 規則錯了，改規則並記錄原因，不要繼續硬遵守。

## 變更記錄

- 2026-07-03 建立（Fable 5 session）。
