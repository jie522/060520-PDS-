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

## Workflow（申請單流程）

- 檔案 AddFile 後是「取出」狀態、**沒有 workflow 狀態**（CurrentState.Name 是空字串）。
- **第一次簽入（UnlockFile）時才決定初始狀態**：由「00_文件分類」等分派流程依 DB 變數
  （關鍵是 `00文件分類`）自動轉送。簽入前沒把變數寫進 DB → 條件不成立 → 掉進
  **「其他文件歸檔」死路狀態**（沒有任何轉換能出去、一般使用者無取出/刪除權限，只能請管理員救）。
- 正確順序：`AddFile → SetVar(全部變數,'') → Flush → UnlockFile → 等自動分派（實測約 2 秒，
  輪詢 CurrentState）→ 需要繼續編輯就 LockFile(folderID, 0) 取出`。
- 變更狀態：`f5.ChangeState('目標狀態名', folder.ID, 註解, 0, 0)`，例如提出申請=
  `ChangeState('單位主管審核', ...)`（轉換「00-提出申請」的目標狀態）。

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

## 已知遺留物（待 PDM 管理員清理）

- `01-模檢治具(2026後)/PTTEST_API_DELETEME` 資料夾（API 測試遺留）
- `PT2607004/PT2607004機器模檢治具申請單.xlsm`（卡「其他文件歸檔」死路的舊檔）
- `PT2607004/..._TEST.xlsm`（流程實驗檔，狀態正確、含測試資料，可改名沿用或刪）
