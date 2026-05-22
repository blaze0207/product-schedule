# 撚二科生產戰情室 4.0 旗艦品質版技術手冊 (GEMINI4.0.md)

本文件定義 4.0 「旗艦品質版」的最終架構。此版本完美整合了品質數據、計畫備註、多重連動篩選以及自動化美化報表。

## 1. 系統架構 (System Architecture)
1.  **`RealityLogAnalyzer` (多源引擎 - reality_analyzer.py)**：
    *   **備註整合**：自動抓取並向下填充（ffill）產銷計畫 C 欄備註，確保任務資訊完整。
    *   **品質監控**：整合日報表，追蹤 A 級率（門檻 95%）與定重率（門檻 90.5%）。
    *   **特殊狀態支援**：新增 `SAMPLE` 與 `TEST` 狀態判定。系統會自動識別 AA 欄位的特殊字眼，並在儀表板即時呈現，同時在累計庫存計算中自動排除此類非正式產量。
    *   **對接邏輯**：使用標準化 ID 與 Inventory Key 實施 A/AX 庫存智慧分攤。
    *   **機台對接修正**：自動將生產資訊中的 `S01`/`S02` 映射至計畫中的 `S1`/`S2`。
    *   **智慧狀態顯示 (eff_status)**：區分「機台停機類」與「特殊作業類」關鍵字。即使在機台運行時，若有 `SAMPLE`/`TEST` 紀錄也會優先顯示，避免資訊被遮蔽。
2.  **`ExportEngine` (美化匯出 - export_cleaned_plan.py)**：
    *   **自動生成**：每次同步自動產生美化版 Excel (`*_clear.xlsx`)。
    *   **旗艦設計**：深藍專業表頭、凍結首行、中文字長度自適應欄寬。
3.  **`DashboardRenderer` (渲染層 - generate_dashboard_v3.py)**：
    *   **多重連動篩選**：整合「工程師」與「機台狀態」雙排控制鈕。
    *   **S1/S2 顯示保護**：針對 S1, S2 機台實施顯示豁免。
    *   **智慧搜尋**：支援機台號與 DTY 批號之即時模糊過濾。

## 2. 關鍵維護與同步邏輯 (Maintenance Logic)
*   **全目錄同步 (Full Portability)**：
    *   `更新生產進度表.bat` 已修改為 `git add .`，確保所有腳本、原始 Excel 資料及技術文件均同步至 GitHub。
    *   目的：支援多台電腦無縫切換，只需 Git Pull 即可直接執行完整環境。
*   **排除規則 (.gitignore)**：
    *   自動忽略 `__pycache__`、`~$*.xlsx` 等暫存與臨時檔案，保持儲存庫乾淨。
*   **環境清理**：
    *   已移除所有舊版 (v2.0, v3.0, v3.5) 資料夾與偵錯腳本，目前僅保留 v4.0 核心檔案。

## 3. UI/UX 規範
*   **佈局防禦**：手機端標籤區縮減至 110px，增加數據展示空間。
*   **即時反饋**：狀態指示燈（閃爍綠、紅、灰）直觀呈現現場動態。

## 4. 核心檔案清單 (v4.0 Cleanup Version)
*   **核心引擎**：`reality_analyzer.py`, `generate_dashboard_v3.py`, `export_cleaned_plan.py`
*   **自動指令**：`更新生產進度表.bat`
*   **設定文件**：`.gitignore`, `GEMINI4.0.md`
*   **輸出成果**：`production_dashboard.html`, `*_clear.xlsx`
*   **穩定備份**：`backup_v4.0_stable/`
