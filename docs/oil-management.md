# 油品管理（油品主檔／MSDS／使用單位／更換記錄）

> 何時要讀這份文件：修改 `templates/oil.html`、`/api/oil/*`、或 `build_oil_index.py` 時。

頂層分頁「油品管理」`/oil`，左側選單版面比照設備管理（`docs/equipment-master.md` 第六節）。

## 一、資料來源

```
\\192.168.1.99\加工部-資料夾\【技術資料】\O.油品\      ← config.OIL_ROOT_PATH
├── MSDS\                    ← 油品主檔的來源：檔名就是資料（見下方二）
│   ├── 報廢\                ← 這裡的檔案 = 停用油品的舊版 MSDS
│   └── 2025.3.5暫存\        ← 根目錄同批檔案的副本，**整個略過**（SKIP_DIRS）
├── 油品簡介\                ← 型錄／使用注意事項（.docx 等），只建索引
├── 油品更換記錄表.xlsx      ← 一次性匯入來源（見下方五）
├── 油品記事本.txt / 200415 油品規範.pptx   ← 系統不讀，人看的
├── oil.db                   ← ★正本（跟 equipment.db 一樣放網芳共用）
└── 油品主檔.xlsx            ← 系統匯出的檔案（每次重建，給不用系統的同事看）
```

路徑一律寫在 `config.py` 的 `OIL_*`，不要寫死在程式裡。

**oil.db 是正本，MSDS 檔名只是第一次的匯入來源**：匯入後所有新增／修改都在系統內做。
`source='manual'`（在系統內編輯過）的油品，重新掃描時完全不會被覆蓋。
`origin` 則永不變更，用來擋「掃描來的油品被硬刪」——這兩個欄位不可合併，
理由與設備主檔完全相同（見 `docs/equipment-master.md` 第二節的實測教訓）。

## 二、MSDS 檔名解析（`parse_msds_name`）

檔名裡的中括號就是結構化欄位，實際存在四種寫法：

| 檔名 | 解析結果 |
|---|---|
| `[唯勝][AW-30]STORK金屬加工液 SW5031-030_2025.1.9.pdf` | 供應商=唯勝、代號=AW-30、版本日期=2025-01-09 |
| `[美科][9520]SDS SUPERSYN 9520詠基…pdf` | 供應商=美科、代號=9520 |
| `[油污清潔劑]MSDS.pdf` | 代號=油污清潔劑（沒有供應商） |
| `[唯勝]SE0551-2 STORK抗磨損液壓油 AW 32,被DS32取代.pdf` | 供應商=唯勝、代號取括號外第一段=SE0551-2、替代油品=DS32 |

**第一個中括號是供應商還是代號，靠 `KNOWN_SUPPLIERS` 名單分辨**（唯勝／美科／昱礽／
久昌／快密刀）。名單以外的一律當成代號——寧可代號怪一點讓使用者進系統改，
也不要把品名誤判成供應商。新的供應商進來時記得加進這份名單。

**`_clean_name()` 會把 SDS/MSDS/GHS/安全標示這類沒有情報量的字清掉**，清完只剩數字
（`SDS1080225` → `1080225`）就當作沒有品名，改用代號當品名。否則清單上會看到
一支叫「GHS」的油品（煤油的 MSDS 檔名就長那樣）。

**「報廢」資料夾裡的檔案不參與更新**（只在建立新記錄時決定初始狀態＝停用）：
那些檔名常帶「,被DS32取代」「_S7取代」的註記，拿去覆蓋現行品名只會把好資料弄髒。

**停用的油品不會因為 MSDS 根目錄又出現新檔案就自動改回使用中**——
「這支油還在不在用」是人的決策，不該由檔案位置反推。

## 三、資料庫（`oil.db`，schema 定義在 `build_oil_index.py` 的 `SCHEMA`）

```sql
oil        (code PK, name, brand, supplier, category, spec, pack,
            status,        -- 使用中 / 停用
            replaced_by, usage_note, remark,
            source, origin, created_at, updated_at)
oil_file   (relpath PK, root, code, folder, filename, ext, size, mtime,
            file_date, obsolete, claimed)      -- root = 'msds' | 'doc'
oil_unit   (code, unit, equip_code, note, sort)          -- 使用單位
oil_change (id PK, code, equip_code, equip_name, date, qty, operator, note, user)
oil_history(code, date, action, detail, user)
```

`app.py` 不另外抄一份 schema，`_oil_conn()` 直接 `from build_oil_index import SCHEMA`
再 `executescript`（全是 `CREATE TABLE IF NOT EXISTS`），沒跑過匯入的電腦也能開頁面。

**`relpath` 是 `oil_file` 的主鍵，格式必須與掃描時一致**（`os.path.relpath` 的原生分隔符，
Windows 上是反斜線）。上傳端點原本寫成 `.replace(os.sep, '/')`，結果下次重新掃描時
同一個檔案被當成新的一筆插進去，清單上出現兩份——2026-08-11 實測踩到，已修正。

**`GROUP BY` 一定要寫完整運算式，不能寫 `GROUP BY name`**：`oil` 表本身就有 `name` 欄位
（品名），別名會被同名欄位蓋掉，分類統計會變成「每支油品各一組、每組 1 支」。
同樣實測踩過。

## 四、頁面結構（`templates/oil.html`）

- 📁 **油品資料**：🛢️ 油品清單、➕ 新增油品
- 📁 **文件與記錄**：📄 MSDS 文件、🔁 更換記錄

### 油品清單（預設子頁）
- 搜尋語法同全站：空格=AND、`-` 前綴=NOT；跨 代號／品名／品牌／供應商／分類／規格／
  說明／使用單位 全文搜
- **排序：分類優先 → 油品代號**（`_oil_sort_key`，清單與匯出共用）。
  分類順序由 `_OIL_CATEGORY_ORDER` 明列（切削類→潤滑類→其他用途，與新增表單的分類
  下拉選單同一順序），**不可以直接 `ORDER BY category`**——中文是照 Unicode 碼位排，
  「主軸油」會排到「切削油」前面，同性質的油品也不會相鄰。
  代號前掛的中文說明（`「全合成切削液」CS-1010`）排序時會先剝掉，照真正的代號排，
  否則加了說明的油品會全部被擠到該分類最後面。
- **狀態預設篩「使用中」**（2026-08-11 使用者要求，`FIRST_LOAD` 那段）：停用的油品
  已經不能領用，平常不該佔畫面。下拉選單保留「全部／停用」，切過之後就不再強制，
  按「清除篩選」也會回到全部
- 篩選列是分類／供應商／狀態／MSDS 四個下拉（標題帶支數），表格徽章單擊＝設定選單、
  雙擊＝清空（`docs/badge-filter-convention.md`）
- 工具列紅字提示「N 支使用中的油品沒有現行 MSDS」，點下去直接篩出那幾支——
  這是這個模組最該被看到的風險（法規要求現場備有 MSDS）
- 點列 → 右側詳情抽屜：基本資料／說明／使用單位／MSDS／相關文件／更換記錄／異動歷程，
  抽屜工具列可直接上傳 MSDS 或其他文件
- **「🔒 快速編輯」解鎖後，油品代號與品名可以在表格裡直接改**（`/api/oil/inline_save`，
  2026-08-11 使用者要求）：工具列按鈕 → 輸入密碼 `maxclaw`（跟設備管理「編碼鎖定」同一組）
  → 兩欄變成可編輯，按鈕轉橘色提示解鎖中。平常看起來就是純文字，滑鼠移過去才浮出輸入框的
  框線，改完按 Enter 或點別的地方自動存檔，綠框閃一下代表存好了（比照設備管理維修記錄的
  「停機時數」欄）。**鎖著時完全不輸出 `contenteditable`**，不是靠 CSS 擋。
  這只是防手滑，不是權限控制（密碼寫在前端 JS 裡）。

  **改代號不是單純 UPDATE**：代號是主鍵，`inline_save` 的 `field='code'` 分支會比照
  `/api/oil/save` 把 `oil_file`／`oil_unit`／`oil_change`／`oil_history` 一起搬過去，
  寫一筆「重新編碼」異動歷程（品名不寫歷程，改代號要寫——那是會牽動一整串子表的異動），
  並呼叫 `_oil_sync_maint_code()` 更新設備保養基準書的 `mt_item.oil_code`。前端另外加一次
  `confirm()`：品名打錯改回來就好，改代號值得停下來確認。重複代號／空白代號後端擋下來後，
  前端會把儲存格還原成原值，畫面上不留假資料。

  三個實作重點：
  - 用 `contenteditable` **不是** `<input>`：品名常常比欄寬長
    （`STORK金屬加工液 SW5031-030`），input 會把後面切掉看不到，contenteditable
    跟原本的純文字一樣會自動折行。Enter 要 `preventDefault()`（品名是單行），
    Esc 還原成 `dataset.orig`
  - 欄位的 `click`/`keydown` 都要 `stopPropagation()`，否則點下去會順便把詳情抽屜打開
  - 存檔成功後只更新 `ALL` 裡那一筆（`rec.name`、`rec.source='manual'`），
    **不要重新 render**，否則游標會跳掉
- 後端 `_OIL_INLINE_FIELDS` 是白名單（目前只有 `name`），不接受任意欄位名；
  存檔一律把 `source` 標成 `manual`——品名正是重新掃描會從檔名覆蓋回去的欄位，
  使用者親手改過就不該再被蓋掉。品名刻意不寫進異動歷程（同 `_OIL_TRACKED_FIELDS` 的取捨）

### MSDS 文件
跨油品列出 `oil_file` 全部檔案。「歸屬油品」欄是個下拉選單，掃描時歸不了位的檔案
（例如 `油品簡介\美科(久昌代理)\250102 …注意事項.docx`）可以直接在這裡指定給某支油品，
指定後標 `claimed=1`，重新掃描不會被搶走。版本篩選預設「僅現行」，報廢區的舊版要看才切。

### 更換記錄
跨油品的更換歷史，可篩油品／設備／來源。「＋ 登錄更換」展開表單登錄新的一筆；
設備下拉來自設備主檔（`/api/oil/equipment_options`，只列使用中與閒置）。
**只能刪 `user='user'` 的記錄**，Excel 匯入的（`user='excel'`）是舊資料軌跡，
要整批換掉請重跑 `build_oil_index.py`（它會 `DELETE WHERE user='excel'` 後重寫）。

## 五、舊「油品更換記錄表.xlsx」的結構

一張分頁一支油品，A1 標題寫著代號（`【加工課】油品記錄表(水性) AW30`，**沒有連字號**，
所以比對前要 `norm_code()` 去掉 `-`）。第 3 列是表頭（設備 / 1 / 2 / 3…），
第 4 列起 A=設備舊編號（綜銑01）、B=機台型號、C 欄以後是第 1、2、3… 次的更換日期。

設備舊編號會透過 `equipment.old_code` 對回 PDM 編碼（綜銑01 → A03-101）；
設備主檔讀不到時只留名稱，不影響匯入。實測匯入 10 筆（該表只有 AW-30 一張分頁，
且只有前 8 台機器填了日期）。

## 六、上傳檔案（含 Excel）

`/api/oil/file/upload` 的 `kind`：
- `msds` → 存進 `MSDS\` 根目錄（沿用既有的扁平結構）。檔名沒有 `[` 開頭時會自動補成
  `[供應商][代號]原檔名`，下次重新掃描才認得回這支油品。
- `doc` → 存進 `油品簡介\<代號>\`。

允許的副檔名見 `_OIL_UPLOAD_EXT`，**含 .xlsx/.xlsm/.csv**——使用者要求「平台可以放
EXCEL 資料存放」，油品的用量表／濃度量測記錄這類 Excel 就是丟這裡，跟著油品一起看。
撞名自動加時間戳記，不覆蓋既有檔案。

## 七、匯出

`/api/oil/export` 用 `openpyxl.Workbook()` 開**全新活頁簿**寫到 `config.OIL_EXPORT_XLSX`
（`油品主檔.xlsx`，兩張分頁：油品主檔／更換記錄），整份重建。

**刻意不寫回 `油品更換記錄表.xlsx`**：本專案的鐵則是「不要用 openpyxl 開既有檔案再存回去」，
有公式的檔案會被洗掉快取值（設備主檔實測過，見 `docs/equipment-master.md`）。
檔案被開啟中時退存到桌面，訊息回報實際存檔位置。

## 八、已完成的端點

| 端點 | 說明 |
|---|---|
| `GET /oil` | 油品管理頁面 |
| `GET /api/oil/search?q=` | 油品清單（空格=AND、`-`=NOT） |
| `GET /api/oil/stats` | 分類／供應商／狀態支數、缺 MSDS 數、未歸位檔案數 |
| `GET /api/oil/detail?code=` | 單支詳情（檔案／使用單位／更換記錄／歷程） |
| `POST /api/oil/save` | 新增或編輯（含改代號，子表一起搬移） |
| `POST /api/oil/delete` | 預設軟刪除（改停用）；`hard=true` 僅限 `origin='manual'` |
| `GET /api/oil/file?root=&relpath=` | 開啟 MSDS／文件（PDF 內嵌、其餘下載，`_eq_safe_path` 擋路徑穿越） |
| `POST /api/oil/file/upload` | 上傳 MSDS／文件／Excel |
| `POST /api/oil/file/claim` | 人工指定檔案歸屬油品 |
| `GET /api/oil/file/list?q=` | 跨油品檔案清單 |
| `GET /api/oil/change/list?q=` | 更換記錄清單 |
| `POST /api/oil/change/save` / `delete` | 登錄／刪除更換記錄 |
| `GET /api/oil/equipment_options` | 設備下拉選項（來自 equipment.db） |
| `POST /api/oil/rebuild` | 重新掃描（直接呼叫 build_oil_index 的函式，不另開 process） |
| `POST /api/oil/export` | 匯出 `油品主檔.xlsx` |
| `POST /api/oil/open_folder` | 用檔案總管開啟油品資料夾 |

## 八之一、被設備保養基準書引用

設備保養基準書的保養項目有「使用油品」欄，選單直接讀這裡的油品主檔
（`GET /api/equipment_master/maint/oil_options`），存的是 `oil.code`，**品名一律以主檔
當下的值顯示、不存快照**——這裡改品名或改成停用，基準書上立刻跟著變（停用的會標紅字）。
詳見 `docs/equipment-maintenance.md`。

因此本頁 `/oil` 支援 `?code=` 深連結（檔尾 `_deepCode`）：帶代號進來會直接開那支油品的
詳情抽屜，讓保養基準書上的油品徽章一點就能看到 MSDS。

**改代號會自動同步基準書的引用**（`_oil_sync_maint_code()`）：`oil.db` 與 `equipment.db`
是兩個不同的資料庫檔案、沒有外鍵，所以在改完代號後另外開 `equipment.db` 把
`mt_item.oil_code` 一起更新，回應訊息會告訴你更新了幾筆。**清單快速編輯與 `/api/oil/save`
的改代號都要記得走這條**——漏掉的話基準書上會顯示成紅色「（主檔查無）」。
（硬刪除油品沒有做這層同步，刪掉的油品在基準書上就是要顯示成查無，那才是對的。）

異動歷程只追蹤 `_OIL_TRACKED_FIELDS`（狀態／供應商／分類／規格／包裝／替代油品），
品名與說明這種隨手改的刻意不記，理由同設備主檔——記進去只會把歷程洗版。

## 九、`oil.db` 是執行期資料庫，不要加進 sync_to_dist.py

跟 `equipment.db` 一樣放在網芳共用，開發機與桌面版讀寫同一份，**沒有本機副本**，
所以 `sync_to_dist.py` 只同步 `build_oil_index.py` 原始碼，不碰資料庫。
`_oil_conn()` 有 `timeout=15`、刻意不開 WAL（Windows SMB 對 WAL 支援不穩定），
理由與限制見 `docs/equipment-master.md` 第十一節。

## 十、首次匯入的資料現況（2026-08-11）

油品 **23 支**（使用中 18、停用 5）、MSDS 檔案 31 個、簡介文件 2 個（都未歸位）、
更換記錄 10 筆。供應商：唯勝 16、美科 3、昱礽 1、快密刀 1、未填 2。

待使用者補的資料：
- 規格（稀釋濃度／黏度）、包裝、使用單位全部是空的，要在系統內逐支補
- `9520`、`CS-1010` 分類是「其他」（檔名看不出來），要人工指定
- `SUPERSYN 900AL` 這種帶空白的代號、`SE0551-2` 這類用產品編號當代號的，
  可在編輯頁改成正式代號（會連帶搬移 MSDS 對應與更換記錄）
- 兩個簡介文件（美科 900AL 注意事項）要在 MSDS 文件頁指定歸屬油品
