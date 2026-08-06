# 工作日行事曆（管理頁「🗓️ 行事曆」）

> 何時要讀這份文件：改管理頁行事曆子頁、`/api/calendar/*`，或任何需要「這段期間有幾個
> 工作天」的統計功能時。

建於 2026-08-02。目的是讓使用者自己設定每個月的實際工作日（國定假日、颱風假、補班），
供妥善率等統計使用，不用把公司行事曆寫死在程式裡。

## 一、只存「例外日」，不是每天存一筆

```
預設規則：週一～週五 = 上班、週六日 = 放假
cal_day  ：只存跟預設不一樣的日子
```

| 情況 | 資料庫 | 畫面 |
|---|---|---|
| 平日上班（最常見） | **不存** | 白底 |
| 週六日放假 | **不存** | 灰底 |
| 平日放假（國定假日／颱風假） | 存 `kind='off'` | 紅底 + 「放假」標籤 |
| 週末上班（補班） | 存 `kind='work'` | 綠底 + 「補班」標籤 |

好處：**行事曆完全空白時，行為跟導入前一模一樣**（週一～週五），不需要先建一整年的
資料才能用；使用者只要標出真正的例外即可。`/api/calendar/set` 收到跟預設相同且沒有
備註的設定時會直接 `DELETE` 該筆，表裡不會累積無意義的資料。

## 二、資料庫：`calendar.db`（執行期資料庫，絕對不可同步覆蓋）

```sql
cal_day (date TEXT PK,     -- YYYY-MM-DD
         kind TEXT,        -- 'work' | 'off'
         note TEXT,        -- 假日名稱：中秋節／颱風假／補班
         updated_at TEXT)
```

`CAL_DB_PATH = _APP_DIR/calendar.db`，`_cal_conn()` 開檔時自動 `CREATE TABLE IF NOT EXISTS`，
所以第一次使用不需要任何前置作業。

**這是三個執行期資料庫裡最脆弱的一個**：`equipment.db`／`cnc_program_index.db` 至少
還有 `build_*_index.py` 可以從 Excel／網芳重建，但 `calendar.db` **沒有任何 build 工具會
產生它**——資料是使用者一天一天點出來的，被開發機的版本蓋掉就真的救不回來。
`sync_to_dist.py` 只複製明確列出的 `.py`／`config.py`／`templates/`／`static/`，
不會碰 `*.db`；日後要加同步項目時千萬不要把資料庫加進去（背景見 CLAUDE.md 的同步鐵則）。

## 三、API

| 端點 | 說明 |
|---|---|
| `GET /api/calendar/month?ym=2026-08` | 該月每一天：`date/day/dow/kind/note/is_override/default_kind`，另回傳 `first_dow`（月初對齊用）與 `workdays`／`offdays` |
| `POST /api/calendar/set` | `{date, kind, note}`；`kind` 只接受 `work`／`off`，跟預設相同且無備註時改為刪除 |
| `POST /api/calendar/reset_month` | `{ym}` 清掉整個月的例外日，回到預設 |
| `GET /api/calendar/workdays?start=&end=` | 任意區間的工作天數，給其他功能呼叫 |

核心函式 `_cal_workdays(start, end)`：逐日查例外，沒設定就用 `_cal_default_kind()`。

## 四、已接上的功能

**設備妥善率**（`docs/equipment-master.md` ⑥）：`_eq_workdays()` 已改成直接呼叫
`_cal_workdays()`，所以在行事曆標了國定假日，妥善率的「工作天／應有稼動時數」會立刻跟著變。
實測 2026-08：預設 21 天 → 標中秋節放假後 20 天 → 妥善率頁的應有稼動時數同步變成
26 台 × 8h × 20 天 = 4,160 小時。

日後要接其他統計（產能、稼動、人均產出…）一律呼叫 `_cal_workdays()`／
`/api/calendar/workdays`，不要各自再寫一份 `weekday() < 5`。

## 五、前端

管理頁「工具」群組下的 `🗓️ 行事曆`（`initCalendar()` / `loadCalendar()`）。

- 月曆用 CSS grid 7 欄，月初用 `first_dow` 個 `.cal-cell.blank` 補齊對位
- **點日期格切換上班／放假**；變成例外日時格子裡才會出現小的「假日名稱」輸入框
  （31 格都放輸入框太雜，只有真的需要命名的那幾天才顯示）
- 點輸入框本身不會觸發切換（`event.stopPropagation()` + `calToggle` 內再擋一次 `INPUT`）
- 每次存檔後整月重畫，工作日統計才會即時更新
- 日期字串一律 `getFullYear()/getMonth()/getDate()` 手組，**禁用 `toISOString()`**（時區差一天）
