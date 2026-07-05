# 模型調度守則（主 session 如何用 subagent 省 token、保品質）

> 讀者：本專案的主 session（預期為 Sonnet 等級）。寫於 2026-07-03。
> 模型名稱與參數為當日實證值；**每次使用前以你當下系統提示裡的 Agent 工具說明為準**，
> 清單有出入時以系統提示為準並更新本檔。

## 0. 當日實證的環境值（2026-07-03）

- Agent 工具 `model` 參數可填：`sonnet` / `opus` / `haiku` / `fable`
  （`fable` 為限時模型，之後的 session 大概率不可用；不要指定它，失敗會浪費一輪）
- Agent 工具**沒有 effort 參數**。effort（low/medium/high/xhigh/max）只存在於 `/code-review` skill。
- subagent 類型（以系統提示為準）：`Explore`（唯讀搜尋，可指定 quick/medium/very thorough）、
  `Plan`（規劃）、`general-purpose`（全工具）、`claude`（泛用）、`claude-code-guide`（查 Claude Code 用法）
- 主對話切換模型：使用者以 `/model claude-sonnet-5` 等指令操作，你不能替使用者切換。
- 未確認：安全機制把個別請求導向 Opus 4.8 時是否計入本窗口額度 → 建議使用者到 usage 儀表板實測。

## 1. 指揮官不下場

主對話（你）只做三件事：**拆任務、做「換便宜模型就掉品質」的判斷、整合結論**。
以下工作一律派 subagent，主對話只收結論：

| 工作型態 | 派誰 | model |
|---|---|---|
| 「X 在哪個檔案/怎麼實作的」跨檔搜尋 | Explore | haiku |
| 掃整個 repo 找某 pattern 的所有出現點 | Explore（very thorough） | haiku |
| 讀一批長檔案摘要重點 | general-purpose | haiku |
| 批次機械性改檔（pattern 明確、已有一個成功範例） | general-purpose | sonnet |
| 驗證產出（read-back、跑測試） | general-purpose（fresh context） | sonnet |
| 設計取捨、外部 API 未知行為、多步推理 | **不派，主對話自己做**（達到 §4 門檻時才升級） | — |

**不派的情況**：目標檔案和行號已知、改動 <30 行 —— 自己 Grep+Edit 比 spawn 便宜。
spawn 一個 agent 的成本 ≈ 它要冷啟動重讀 context，小事別派。

## 2. 交辦三要素（每個 subagent prompt 必含）

1. **目標與動機**：做什麼＋為什麼（讓它能在細節模糊時做出對的取捨）
2. **驗收條件**：完成的客觀判準（「找到所有呼叫點」「py_compile 通過」「三個欄位值與來源一致」）
3. **回報格式**：明確規定（見 §3）

範本見 `docs/agent/prompts.md`。

## 3. 回報合約

- subagent 只回**結論＋「檔案:行號」**，不貼大段程式碼。
- 長產物（報告、diff、清單）→ 存到 scratchpad 或 docs/，回傳路徑。
- 回報開頭第一句必須是結論（成功/失敗/部分完成），失敗要附：試了什麼、錯誤原文、卡在哪。

## 4. 升降級路徑

- **haiku 錯 1 次** → 同一子任務升 sonnet 重派（附上 haiku 的失敗輸出）。
- **sonnet 同一子任務連錯 2 次** → 帶完整失敗軌跡（兩次的 prompt、輸出、錯誤）升 opus，
  或改由主對話親自處理。**不要第三次原樣重派**。
- **解出模式後降級**：一旦某類修改有了一個驗證過的成功範例，剩餘同類批次工作降回
  haiku/sonnet 照範例套用。
- **同一件事最多重試兩輪**（含升級），還不行 → 停下來，把失敗軌跡整理給使用者選方向。
  連續失敗通常代表方向錯，不是力氣不夠（判準見 `docs/agent/judgment.md` §4）。

## 5. 驗證不自驗

寫程式的人不能自己驗收。規則：

- **檔案類產出**：派 fresh-context 的 general-purpose agent 做 read-back
  （「讀 X 檔，回答：是否存在、行數、是否包含 A/B/C 要點、有無與 Y 檔矛盾」）。
- **程式碼改動**：能跑測試就跑；本專案多數改動的驗證方式是
  `python -m py_compile` ＋ 重啟 dist 版實測（要使用者配合）＋前端用 `preview_start`/
  `preview_screenshot` 等 Claude_Preview 系列工具驗證（用法見 `docs/dev-workflow.md`）。
- **高風險判斷**（要寫入正式共用資源、要刪東西、外部 API 用法拿不準）：
  主對話生成 2 個候選方案 → 派一個 fresh agent 當評審擇優，或直接問使用者。
- 驗證 agent 的 prompt 不要透露「這是我寫的」，用中立措辭（「審查以下檔案是否…」），減少附和偏誤。
