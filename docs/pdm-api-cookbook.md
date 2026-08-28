# PDM COM API 實戰手冊（2026-07-03 全數實機驗證）

> 何時要讀：任何要透過程式讀寫 PDM（SolidWorks PDM 2021, vault=MAXCLAW）的時候。
> 這裡每一條都是在本機實測通過的，**不要憑訓練記憶改寫**；手冊沒有的用法，先寫最小唯讀腳本驗證。
> 相關功能：治檢具申請單自動建立（app.py 搜 `jig_apply`）、PDM 圖面索引（build_pdm_index.py）。

## 連線與登入

```python
import win32com.client
from win32com.client import gencache
vault = gencache.EnsureDispatch('ConisioLib.EdmVault')   # 必須用 gencache（早期繫結），否則 CastTo 全部失敗
vault.LoginAuto('MAXCLAW', 0)                            # 沿用本機已登入的 PDM 工作階段，不需帳密
```
- Flask 執行緒內使用前要 `pythoncom.CoInitialize()`，finally 裡 `CoUninitialize()`。
- 型別庫是「PDMWorks Enterprise 2021 Type Library」5.24；完整介面（IEdmSerNoGen7 等）都在其中，
  但部分介面名稱與官方文件版本不同（例如沒有 IEdmSerNoGen5，只有 6/7）。

## ★ 最大陷阱：gen_py 預設參數會炸

pywin32 產生的方法簽名裡，optional 的 COM 物件參數預設值是 `0`，實際呼叫會丟
**「The Python instance can not be converted to a COM object」**。修法：明確傳 `None`。

| 呼叫 | 錯誤寫法 | 正確寫法 |
|---|---|---|
| 建資料夾 | `folder.AddFolder(0, name)` | `folder.AddFolder(0, name, None)` |
| 以路徑取檔 | `vault.GetFileFromPath(path)` | `vault.GetFileFromPath(path, None)`（回傳 tuple：`(file, parent_folder)`） |

## 物件與介面對照（CastTo 用）

| 目的 | 寫法 |
|---|---|
| 資料夾 | `f5 = win32com.client.CastTo(vault.GetFolderFromPath(path), 'IEdmFolder5')` |
| 資料夾卡片變數 | `ev = win32com.client.CastTo(folder_obj, 'IEdmEnumeratorVariable5')`（資料夾物件直接 cast） |
| 檔案卡片變數 | `ev = file5.GetEnumeratorVariable()`（**不能**對檔案物件 CastTo，會「找不到成員」）；要 CloseFile 時再 `CastTo(ev, 'IEdmEnumeratorVariable8')` |
| 序號產生器 | `sg = win32com.client.CastTo(vault, 'IEdmSerNoGen7')` |
| 登入者姓名 | `u = CastTo(vault,'IEdmUserMgr5').GetLoggedInUser()`；工號=`u.Name`，中文姓名要 `CastTo(u,'IEdmUser6').FullName` |
| 檔案 by ID | `fobj = vault.GetObject(1, file_id)`（1=EdmObject_File），再 CastTo IEdmFile5 |

## 序號（PT 專案代號）

```python
sv = sg.AllocSerNoValue('PT_PT專案序號(2026後)')  # 取號（跟檔案總管範本同一個計數器，不會撞號）
pt_no = sv.Value                                   # 例 'PT2607001'
sv.Rollback()                                      # 不用時還號（預覽下一號的安全作法）；正式使用就不要 rollback
```

## 卡片變數：三個存放處，缺一不可

這是整個 PDM 整合最容易搞錯的概念。同一個變數值存在三個地方，寫入時機不同：

1. **資料庫（DB）** — `ev.SetVar(名稱, '', 值)` + `ev.Flush()`。組態名稱一律用空字串 `''`
   （用 `'@'` 會多建一個空殼組態，卡片上多出一個空白分頁）。
   **workflow 自動分派的條件只看 DB 值**。
2. **檔案本體**（xlsm 的 `docProps/custom.xml`）— `IEdmEnumeratorVariable8.CloseFile(True)` 會把
   SetVar 的值寫進檔案。**取出中的檔案，資料卡顯示的是這裡的值**；只 Flush 不 CloseFile，卡片看起來沒存到。
   也可以在檔案入庫前直接用 zipfile 改 custom.xml（見 app.py `_fill_xlsm_custom_vars`）。
3. **Excel 工作表儲存格**（Defined Names 如 `YC_機型`）— PDM 不會自動同步到這裡；openpyxl 依
   defined_names 寫入（見 app.py `_write_xlsm_defined_names`）。**治檢具索引（build_jig_index.py）讀的是儲存格**，
   讀不到會 fallback 到 custom.xml。

### 陷阱：卡片上的「勾選框」通常不是綁 ■/□ 那個變數（2026-08-18 實測）

設計變更申請單每個「申請設變原因」其實有**兩個**變數：

| 變數 | 值 | 存在哪 | 誰在看 |
|------|----|--------|--------|
| `PP_R_004_製造問題_tasky` | `■製造問題` / `□製造問題` | DB + custom.xml + 儲存格 | Excel 列印出來的樣子 |
| `PP_R_004_製造問題F_tasky` | `1` / `0` | **只有 DB** | **PDM 資料卡上的勾選框** |

`F` 系列在 xlsm 的 custom.xml 與 defined names 裡都查無此名，所以：
- 寫：`ev.SetVar(名稱, '', '1')` + `ev.Flush()`（CloseFile 對它沒意義）。
- 驗：只能 `GetVar` 讀 DB，**不能**沿用讀檔案 custom.xml 那套讀回比對（會永遠 mismatch）。

只寫 `_tasky` 不寫 `F_tasky` 的症狀：Excel 印出來有 ■，但資料卡七個框全空。
碰到別張卡片的勾選框沒反應時，第一件事就是**把人工建立的同類文件全部變數 dump 出來 diff**
（`IEdmVariableMgr5` 列 vault 全部變數 → 對兩份檔案逐一 `GetVar`），一次就會看到多出來的那組。

### 多行文字要寫 CRLF

資料卡的多行文字框只認 `\r\n`，只寫 `\n` 會擠成一整段；Excel 儲存格則慣用 `\n`
（人工建單也是這樣）。所以 SetVar 給 CRLF、`_write_xlsm_defined_names` 給 LF。
另外讀回 custom.xml 比對時，XML 解析器會把 CRLF 正規化成 LF，比對前要先正規化換行
（app.py `_dcn_same_text`）。

## 「貼上為參考」：IEdmAddCustomRefs（2026-08-18 打通）

把附件掛成申請單 xlsm 的參考（＝檔案總管的複製→貼上為參考）：

```python
ob = vault.CreateUtility(8)          # EdmUtil_AddCustomRefs
ob = ob._oleobj_
ob.InvokeTypes(2, 0, pythoncom.DISPATCH_METHOD, (24, 0),      # AddReferencesPath
               ((3, 1), (0x4000 | 0x2000 | 8, 1)), 主檔FileID, [附件完整路徑])
ob.InvokeTypes(4, 0, pythoncom.DISPATCH_METHOD, (11, 0), ((3, 1),), 0)   # CreateTree
ob.InvokeTypes(6, 0, pythoncom.DISPATCH_METHOD, (11, 0), ())             # CreateReferences
```

踩過的坑：
- **gen_py 包不出這個介面**：`CastTo(..., 'IEdmAddCustomRefs')` 拿到的物件只有 CLSID，
  沒有任何方法（`AttributeError: no attribute 'AddReferencesID'`）。
- **dynamic dispatch 會回「類型不符」**：`AddReferencesPath` 第二個參數的型別是
  「SAFEARRAY(BSTR) 的**指標**」，pywin32 推不出來，必須用 `InvokeTypes` 明確指定
  `VT_BYREF|VT_ARRAY|VT_BSTR` = `0x4000|0x2000|8` = 24584。
- memid 是從型別庫查的：`AddReferencesPath=2`、`CreateTree=4`、`CreateReferences=6`。
- 主檔取出中做沒問題，簽入後參考仍在（實測 RR2608017）。
- 唯讀確認結果：`f5.GetReferenceTree(folder.ID, 0)` → `IEdmReference5.GetFirstChildPosition('', True, True, 0)`
  （這個回傳的是 tuple `(pos, projectName)`，要取出有 `IsNull` 的那個才是 position）。

## Workflow（申請單流程）

- 檔案 AddFile 後是「取出」狀態、**沒有 workflow 狀態**（CurrentState.Name 是空字串）。
- **第一次簽入（UnlockFile）時才決定初始狀態**：由「00_文件分類」等分派流程依 DB 變數
  （關鍵是 `00文件分類`）自動轉送。簽入前沒把變數寫進 DB → 條件不成立 → 掉進
  **「其他文件歸檔」死路狀態**（沒有任何轉換能出去、一般使用者無取出/刪除權限，只能請管理員救）。
- 正確順序：`AddFile → SetVar(全部變數,'') → Flush → UnlockFile → 等自動分派（實測約 2 秒，
  輪詢 CurrentState）→ 需要繼續編輯就 LockFile(folderID, 0) 取出`。
- 變更狀態（2026-07-07 實測更正，舊記載錯誤）：
  - **Python IDispatch 呼叫 `f5.ChangeState` 一律回 DISP_E_MEMBERNOTFOUND「找不到成員」**，
    不管 IEdmFile5/IEdmFile17、早期/晚期繫結、VARIANT 包裝都一樣——此方法只在 vtable 上，
    必須走 .NET Interop（`C:\Program Files\SOLIDWORKS PDM\EPDM.Interop.epdm.dll`）。
    app.py 的做法：subprocess 呼叫 `pdm_change_state.ps1`（PowerShell Add-Type 編 C#）。
  - 只帶目標狀態的 `ChangeState`/`ChangeState2` 會回「指定的工作流程狀態資料庫 ID 是無效的」，
    要用 **`ChangeState3(ref 狀態名, ref 轉換名, folderID, 註解, 0, 0, 密碼)`** 同時指定轉換。
  - 「00-提出申請」轉換設有**身分驗證**：密碼參數必須是操作者的 PDM 登入密碼，
    空密碼回「密碼是無效的」。密碼不落地，由前端表單即時輸入傳入。
  - 列出某狀態可走的轉換（唯讀）：`IEdmState6.GetFirstTransitionPosition(True)`（True=送出方向，
    False 是列「進入」此狀態的轉換）→ `GetNextTransition`。

## 權限現況（帳號 990602 楊良捷，2026-07 實測）

- 可以：建資料夾、加檔案、寫卡片、簽入/取出、ChangeState（有轉換權限者）
- **不可以：刪除資料夾、刪除檔案**（丟 -2147220989「您沒有執行此動作的許可權限」）
  → 建錯的東西自己刪不掉，要請 PDM 管理員。所以**寫入前想清楚，測試一律唯讀**；
  必須驗證寫入行為時，依 `docs/agent/judgment.md` §3 先讓使用者知道要寫什麼再動手。

## openpyxl 與 PDM 檔案共處

- openpyxl 重寫 xlsm 會改變 `docProps/custom.xml` 的序列化格式（PDM 的 linkTarget 屬性等）。
  安全做法：openpyxl save 前先把 custom.xml 讀出備份，save 後用 zipfile 整包重寫把原始
  custom.xml 塞回去（見 app.py `_write_xlsm_defined_names` 的實作）。
- 開 xlsm 一律 `keep_vba=True`（保 VBA）；插圖需要 Pillow（dist_embed 已裝 12.3.0）。

## ★ 多台電腦共用同一份 dist_embed 時：gen_py 快取可能損毀（2026-07 實測）

現象：某台電腦點「治檢具索引更新」或「設計變更通知單索引更新」時炸出類似
`AttributeError: module 'win32com.gen_py.<LIBID>x0x5x24' has no attribute 'CLSIDToPackageMap'`
的錯誤，但在其他電腦上一切正常。**實測過至少兩種變體**：使用者回報的是缺
`CLSIDToPackageMap`，但同一類問題在開發機上重現時缺的是 `CLSIDToClassMap`——
共同特徵是錯誤訊息裡一定有 `win32com.gen_py`，不是單一固定的屬性名稱，判斷式
不要只比對某一個屬性字串。

原因：pywin32 的 `gencache.EnsureDispatch`（早期繫結）第一次呼叫時會把「PDMWorks
Enterprise 2021 Type Library」的 wrapper 產生並快取到本機的 gen_py 資料夾。多台電腦
如果透過網路共用資料夾（`\\192.168.1.99\...\dist_embed\...`）跑**同一份**內嵌 Python，
理論上 gen_py 快取路徑（`%TEMP%\gen_py\<py版本>\`）各自獨立在每台電腦本機，但只要某次
產生過程被中斷、或該機器登錄的 COM 型別庫版本跟快取內容對不上，就會產生一份不完整的
package，且不只影響第一次連線，**後面任何 `CastTo` 呼叫都可能觸發同一種錯誤**
（因為整個型別庫的 wrapper package 都沒正確產生，不是單一介面的問題）。

修法（已寫進程式碼，不需要手動介入）：`build_pdm_index.py` 的 `connect_vault()` 與
`app.py` 的 `_pdm_vault_login()` 現在會攔截這整類特徵錯誤（`AttributeError` 且訊息含
`win32com.gen_py`），自動刪除 `win32com.__gen_path__` 指向的 gen_py 快取資料夾，
讓 pywin32 在下一次 `EnsureDispatch` 時重新產生一份乾淨、跟該機器實際登錄版本吻合的
wrapper，再重試一次連線——使用者端完全無感，不需要手動去 `%TEMP%\gen_py` 刪檔。
**已在開發機上實際重現＋驗證自動修復成功**（第一次 `EnsureDispatch` 炸出
`CLSIDToClassMap` 缺失，清快取重試後第二次成功回傳可用的 COM 物件）。

**不能**改成單純 `except Exception` 就 fallback 成晚期繫結（`win32com.client.Dispatch`
不帶 gencache）長期使用，因為 `CastTo` 一定要早期繫結才能用（見本文件開頭）；只有在
清快取重試也失敗時，才把晚期繫結當最後手段（此時多數功能會直接失敗，但至少不會整支
腳本崩潰無回應）。

## 已知遺留物（待 PDM 管理員清理）

- `01-模檢治具(2026後)/PTTEST_API_DELETEME` 資料夾（API 測試遺留）
- `PT2607004/PT2607004機器模檢治具申請單.xlsm`（卡「其他文件歸檔」死路的舊檔）
- `PT2607004/..._TEST.xlsm`（流程實驗檔，狀態正確、含測試資料，可改名沿用或刪）
