# 導覽列設計規範

> 何時需要讀這份文件：新增頂層分頁（新的 .html 檔案）或修改任何頁面的導覽列時。

所有頁面的導覽列 (`<div class="header">`) 應遵循以下規範，確保視覺一致性。設計參考 ZUMEN 導覽列風格。

### 顏色
- **背景色**：純藍色 `#1565C0`（參考 ZUMEN）
- **標題文字色**：白色 `#FFFFFF`
- **Tab 文字色**（未激活）：半透明白色 `rgba(255,255,255,.85)`
- **Tab 背景色**（未激活）：半透明白色 `rgba(255,255,255,.15)`
- **Tab 文字色 & 背景**（激活/hover）：白色背景 + 深藍文字 `#1565C0`

### 字體 & 字號
- **全頁面字體**：`'Noto Sans TC', 'Source Han Sans TC', sans-serif`
- **標題**：字號 `1.1rem`、加粗 `700`、不換行 `nowrap`
- **Tab**：字號 `1rem`（比預設大 2 號）、加粗 `600`

### 間距與尺寸
- **Header padding**：上下 `0.4rem`、左右 `1rem`（以 `padding: 0.4rem 1rem 0` 為基準）
- **標題 padding-bottom**：`0.45rem`
- **標題右外邊距**：`0.5rem`
- **Tab padding**：上下 `0.3rem`、左右 `1rem`
- **Tab 間距**：`gap: 2px`
- **Tab border-radius**：`0`（直角，無圓角，參考 ZUMEN 風格）

### 動畫 & 其他效果
- **過渡效果**：`transition: background 0.15s`（hover/active 狀態）
- **陰影**：`box-shadow: 0 2px 6px rgba(0,0,0,.3)`
- **位置**：`position: sticky; top: 0; z-index: 200`（黏性定位，頁面置頂）
