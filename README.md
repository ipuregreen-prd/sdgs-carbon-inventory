# 碳足跡排放係數資料庫 每日定時爬蟲與 Excel 匯出系統

本專案專為 **綠色公益基金會 (SDGs 碳盤查計畫)** 開發，自動針對中華民國環境部「產品碳足跡資訊網」排放係數資料庫進行每日定時爬取，並將完整資料結構化匯出為專業美觀的 Excel 活頁簿 (`data.xlsx`)。

---

## 📌 資料來源與目標檔案
* **資料來源**：[中華民國環境部 產品碳足跡資訊網 - 碳足跡排放係數資料庫](https://cfp.moenv.gov.tw/WebPage/WebSites/CoefficientDB.aspx)
* **匯出路徑**：`C:\0-Learning\0-綠色公益基金會\程式\sdgs-carbon-inventory\data.xlsx`
* **目前資料筆數**：**1,164 筆**（涵蓋 2013 年至 2025 年最新公布之碳足跡排放係數）

---

## 📂 檔案清單與功能說明

| 檔案名稱 | 說明 |
| :--- | :--- |
| [`crawler.py`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/crawler.py) | **核心爬蟲模組**：自動處理 ASP.NET AJAX 狀態機制，擷取全部 1,164 筆係數與分類清單，格式化並匯出至 `data.xlsx`。 |
| [`scheduler.py`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/scheduler.py) | **定時排程服務**：純 Python 實現的每日定時常駐排程器，支援自訂執行時間（預設每天凌晨 03:00）。 |
| [`run_now.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/run_now.bat) | **一鍵立即執行**：雙擊即可立即手動觸發爬蟲並更新 `data.xlsx`。 |
| [`start_scheduler.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/start_scheduler.bat) | **一鍵啟動定時服務**：雙擊開啟終端視窗常駐執行定時排程（每天 03:00 自動執行）。 |
| [`setup_windows_task.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/setup_windows_task.bat) | **一鍵註冊 Windows 工作排程器**：將爬蟲註冊為系統原生工作排程，無須開啟視窗即可在每日背景自動執行。 |
| [`uninstall_windows_task.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/uninstall_windows_task.bat) | **一鍵移除 Windows 工作排程器**：如需取消定時任務可雙擊此檔案移除。 |
| [`run_task.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/run_task.bat) | Windows 工作排程器背後呼叫的靜默執行腳本。 |
| [`data.xlsx`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/data.xlsx) | **匯出之碳足跡排放係數 Excel 報表**。 |
| [`crawler.log`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/crawler.log) | 爬蟲執行日誌。 |

---

## 📊 Excel 檔案結構 (`data.xlsx`)

匯出的 Excel 檔案採用綠色主題排版（呼應綠色公益基金會永續風格），包含 3 個工作表：

1. **碳足跡排放係數清冊**：
   * `序號`
   * `係數編號` (官方 `cd_id`)
   * `碳係數名稱`
   * `生產區域` (如：臺灣、台灣)
   * `係數數值(數值)` (可直接計算之純數值型態)
   * `係數數值(科學記號)` (如：`6.058E-01`)
   * `碳排單位` (如：`kgCO₂e`)
   * `宣告單位` (如：`度(kwh)`、`公斤(kg)`、`公升(L)`)
   * `完整數值與單位` (如：`6.058E-01 kgCO₂e / 度(kwh)`)
   * `公告年份` (如：`2025`、`2023`、`2022`...)
   * `數據抓取時間`
   * *內建功能：首列凍結窗格（Freeze Panes）、自動篩選箭頭（AutoFilter）、交替淺綠斑馬紋、最適欄寬自適應調整。*

2. **係數分類架構**：
   * 官方資料庫完整階層分類結構（中類、小類、分類代碼 `cc_id`）。

3. **更新日誌與說明**：
   * 記錄來源網址、最新同步時間、總筆數、年份涵蓋區間與使用說明。

---

## 🚀 使用方式

### 方法 A：手動單次更新
在檔案總管中直接雙擊 [`run_now.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/run_now.bat)，或在終端機執行：
```bash
python crawler.py
```

### 方法 B：使用 Windows 原生工作排程器（推薦，免開視窗背景執行）
1. 在 [`setup_windows_task.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/setup_windows_task.bat) 上按右鍵，選擇 **「以系統管理員身分執行」**。
2. 系統將自動在 Windows 排程器建立名為 `CarbonFootprintDailyCrawler` 的任務。
3. 每天凌晨 **03:00** 電腦開機時自動在背景靜默更新 `data.xlsx`，不需要手動開啟任何黑視窗！
*(如欲移除，請以系統管理員身分執行 [`uninstall_windows_task.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/uninstall_windows_task.bat))*

### 方法 C：使用 Python 內建排程服務常駐執行
若習慣以命令列常駐運行，可雙擊 [`start_scheduler.bat`](file:///c:/0-Learning/0-綠色公益基金會/程式/sdgs-carbon-inventory/start_scheduler.bat) 或在終端機執行：
```bash
# 預設每天凌晨 03:00 執行，啟動時先執行一次
python scheduler.py --time 03:00 --run-now

# 自訂每日執行時間（例如每天早上 08:30）
python scheduler.py --time 08:30 --run-now
```

---

## 🛡️ 反爬蟲與容錯機制
1. **ASP.NET ViewState 動態解析**：即時解析 ASP.NET WebForms 之 `__VIEWSTATE`、`__VIEWSTATEGENERATOR`、`__EVENTVALIDATION`。
2. **AJAX Delta 解碼**：精確解析 MicrosoftAjax UpdatePanel 局部重新整理封包，避免無謂的多餘流量。
3. **頻率限制與 429 退避重試**：針對環境部伺服器的頻率限制（Rate Limiting），內建指數退避（Exponential Backoff）與智慧間隔延遲機制。
