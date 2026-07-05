# Subagent 交辦 prompt 範本（複製後填空）

> 讀者：主 session。搭配 `docs/agent/dispatch.md` 使用。每份範本都含三要素：
> 目標與動機／驗收條件／回報格式。`{...}` 是填空位。

## 1. 搜尋（agent=Explore，model=haiku）

```
在本專案（D:\claude code\04 製令SQL_SEARCH）找出 {要找的東西，例：所有呼叫 /api/jig/list 的前端程式碼}。
動機：{例：我要在回傳裡加欄位，需要知道所有讀取端}。
搜尋範圍提示：{例：templates/*.html 和 app.py；treat dist_embed/ 為副本，跳過}。
breadth: {quick|medium|very thorough}
驗收：列出「每一個」出現點，不確定是否相關的也列出並標註存疑。
回報格式：每行一筆「檔案:行號 — 一句話說明用途」，最後一行寫總數。不要貼程式碼區塊。
```

## 2. 實作（agent=general-purpose，model=sonnet）

```
任務：{做什麼，例：在 templates/xxx.html 的 YYY 表格加一欄「分類」}。
動機：{為什麼，例：使用者要能按分類篩選}。
改前必讀：CLAUDE.md 的「硬規則」與「dist_embed 同步鐵則」；{對應的 docs/ 檔，例：docs/management-page.md}。
已知脈絡：{行號線索、既有同類實作可參考的位置，例：同檔 renderJigTable() 是同樣模式}。
限制：只改 {檔案清單}；不要動 dist_embed/（同步由我做）；遵守繁中註解與字體/CDN 規範。
驗收條件：{客觀判準，例：py_compile 通過；新欄位對 K1P2_CAT_LABELS 找不到值時顯示空白不報錯}。
回報格式：第一句寫成功或失敗；列出改動的「檔案:行號範圍」與 15 字內摘要；失敗附錯誤原文。
```

## 3. 重構（agent=general-purpose，model=sonnet；先有成功範例才派批次）

```
任務：把 {pattern A} 改成 {pattern B}，套用到 {檔案清單}。
動機：{例：統一徽章篩選互動}。
成功範例：{檔案:行號} 已改好並驗證過，照同樣方式套用。
禁止：不要「順手」改範例以外的東西；遇到跟範例結構不同、套不上去的案例，跳過並回報，不要自行發明變體。
驗收條件：每個檔案改完後 {驗證方式，例：py_compile / 該頁面仍可 render}；跳過清單另列。
回報格式：改動清單（檔案:行號）＋跳過清單（檔案＋原因）。
```

## 4. 研究（agent=general-purpose，model=sonnet；需要外部查證時 model=opus）

```
問題：{要回答什麼，例：openpyxl 能否在不破壞 xlsm 的 VBA 與 custom properties 下插入圖片}。
動機：{這個答案會決定什麼，例：決定申請單示意圖用哪種寫入方式}。
方法要求：優先實驗驗證（在 scratchpad 寫最小測試腳本跑），其次查專案內既有程式碼，
最後才靠一般知識；一般知識得出的結論必須標註「未實測」。
本專案注意：測試腳本寫成 UTF-8 檔案執行，不要 heredoc 塞中文；不得寫入任何正式共用資源
（PDM vault、共用區 Excel）——只能唯讀或在 scratchpad 操作副本。
驗收條件：結論有實驗證據或明確標註未實測；含「建議做法」一段。
回報格式：結論（3 行內）→ 證據（實驗輸出摘要）→ 建議做法。長輸出存 scratchpad 回傳路徑。
```

## 5. 審查（agent=general-purpose，model=sonnet，fresh context）

```
審查對象：{檔案清單或 diff 範圍}。（中立措辭，不要說是誰寫的）
審查目標：{例：找正確性 bug；或：驗證文件互相一致}。
檢查清單：
- {具體檢查點 1，例：所有 SetVar 的組態參數是否為 ''}
- {具體檢查點 2，例：新增 API 是否有對應的前端錯誤處理}
- 與 CLAUDE.md 硬規則的牴觸（CDN／字體／toISOString）
驗收條件：每個檢查點都有明確結論（通過/不通過/不適用），不允許「大致沒問題」。
回報格式：問題列表，每筆含「檔案:行號、問題一句話、具體失敗情境」；沒問題就回「全部通過＋各檢查點結論」。
```

## 通用附註（貼在任何 prompt 結尾都適用）

```
回報第一句必須是結論。長產物存到 {scratchpad 路徑} 回傳路徑即可。
你查到的任何與本 prompt 假設矛盾的事實，直接回報矛盾，不要遷就假設硬做。
```
