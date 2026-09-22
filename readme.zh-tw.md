# Office2PDF v5.x（Office 全能 PDF 工具）繁體中文說明

> [English](./README.md) | [日本語](./readme.ja.md) | [简体中文](./readme.zh-cn.md) | 繁體中文

> 為需要可靠、離線批次 PDF 處理的辦公人士設計。

**一款面向辦公場景的批次 PDF 轉換與處理工具。**
批次將 Office 檔案和圖片轉換為 PDF，並可加入浮水印、頁碼、分割/合併，以及靈活的命名規則，幾步點擊即可完成。

> **離線轉換：** 本應用透過 COM 自動化呼叫 Microsoft Office 或 WPS Office。
> Office 文件需安裝可供 COM 呼叫的對應元件；PDF、圖片處理不需要 Office/WPS。

執行按鈕及進度固定在底部，設定區可捲動，日誌可收合。可選「自動（優先 Office）」「Microsoft Office」「WPS Office」，並隨設定及預設儲存。自動模式僅於 Office 元件無法建立時嘗試 WPS；開啟或匯出失敗會記錄錯誤。WPS 支援取決於安裝版本的 COM 能力，目前不新增 `.wps/.et/.dps` 格式。

---

## 🌟 主要功能（v5.x）

### 批次轉換
- 支援以下檔案類型轉換為 PDF：
  - **Word**：`.doc`、`.docx`
  - **Excel**：`.xls`、`.xlsx`、`.xlsm`
  - **PowerPoint**：`.ppt`、`.pptx`
  - **圖片**：`.jpg`、`.jpeg`、`.png`
  - **PDF**：重新儲存/標準化處理

### 靈活的分割 / 合併
- **全部合併**為單一 PDF
- **按頁分割**（Word / PowerPoint / PDF / Excel）
- **按工作表分割**（Excel）

### 範圍選擇
- 雙擊清單項目指定目標頁碼範圍（Word / PowerPoint / PDF）：
  - 範例：`1-3,5,8-` / `2` / `-3` / `1-`
- Excel：選擇一個或多個**工作表名稱**

### 浮水印與頁碼
- 最多 **2 個浮水印**，可設定位置、不透明度、字型、大小和顏色
- **頁碼**支援格式範本（如 `- {n} / {total} -`）
  - *（目前版本）* 頁碼以**黑色、固定 10.5pt** 渲染
- **預覽**：批次執行前可用**第一頁**產生預覽 PDF

### 命名規則與預設
- 使用標籤的靈活輸出命名範本（見下文）
- 儲存/讀取預設（JSON 設定）

### 安全與最佳化
- 批次設定 PDF 密碼
- 移除中繼資料
- 壓縮 PDF（取決於 pypdf 版本支援）

### 易用性
- 拖曳加入檔案
- 清單排序、移除、清空
- 輸出選項：
  - 與原始檔相同資料夾，或自訂資料夾
- 後處理操作：
  - 開啟 PDF / 開啟輸出資料夾 / 清空清單

### 🌐 多語言介面
- 支援日語 / 英語 / 簡體中文 / 繁體中文
- 首次啟動按系統語言自動選擇，也可透過選單列 **Language** 手動切換（儲存後重新啟動生效）

---

## 🏷️ 可用標籤

標籤可用於**命名規則**和**浮水印文字**，處理時自動替換。

- `{name}`：原始檔案名稱（不含副檔名）
- `{sheet}`：Excel 工作表名稱（非 Excel 為空）
- `{parent}`：上層資料夾名稱
- `{seq}`：全域流水號
- `{fseq}`：單一輸入檔案內的流水號
- `{pseq}`：頁面流水號（分割/統計頁數時）
- `{total}`：輸入檔案總數
- `{ptotal}`：目前檔案/單元的總頁數
- `{username}`：Windows 登入使用者名稱
- `{rand}`：4 位隨機數字（如 `0427`）

### 日期/時間標籤
- `{date:yyyy-mm-dd}` → `2026-02-06`
- `{date:yyyy/mm/dd HH:MM}` → `2026/02/06 16:54`

支援的佔位符：
- `yyyy` 年 / `mm` 月 / `dd` 日
- `HH` 時 / `MM` 分 / `SS` 秒

---

## 🚀 使用方法

### 系統需求
- Windows 10 / 11（64 位元）
- **已安裝 Microsoft Office 或支援 COM 的 WPS Office**（所需的文字 / 表格 / 簡報元件）

### 基本步驟
1. 啟動應用程式。
2. 拖曳檔案到清單。
3. （可選）雙擊清單項目設定頁碼範圍 / 工作表選擇。
4. 設定：
   - 浮水印 / 頁碼
   - 分割/合併選項
   - 命名規則和輸出資料夾
5. 點擊 **開始轉換 PDF**。

---

## 📦 安裝

1. 從 **Releases** 下載最新版本。
2. 解壓縮 `Office2PDF.zip`。
3. 執行 `Office2PDF.exe`（無需安裝）。

---

## 🛠️ 從原始碼執行 / 建置可執行檔

### 1) 環境
- Python **3.9+**（Windows）
- 已安裝 Microsoft Office 或支援 COM 的 WPS Office

### 2) 安裝相依性
```bash
pip install pywin32 pypdf Pillow reportlab tkinterdnd2
```

### 3) 執行
```bash
python office2pdf_v5.py
```

### 4) 建置 `.exe`（PyInstaller）
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 --name Office2PDF office2pdf_v5.py
```

建置完成後，可執行檔位於：
- `dist/Office2PDF.exe`

> 提示：某些環境下轉換 PowerPoint 時需保持其可見。
> 如果 PowerPoint 轉換失敗，請避免強制 `Visible=False`。

---

## ⚖️ 授權條款

MIT License — 詳見 ./LICENSE
