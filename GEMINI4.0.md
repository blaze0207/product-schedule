# 撚二科生產戰情室 4.0 旗艦品質版技術手冊 (GEMINI4.0.md)

本文件定義 4.0 「旗艦品質版」的最終架構。此版本完美整合了品質數據、計畫備註、多重連動篩選以及自動化美化報表。

## 1. 系統架構 (System Architecture)
1.  **`RealityLogAnalyzer` (多源引擎 - reality_analyzer.py)**：
    *   **備註整合**：自動抓取並向下填充（ffill）產銷計畫 C 欄備註，確保任務資訊完整。
    *   **品質監控**：整合日報表，追蹤 A 級率（門檻 95%）與定重率（門檻 90.5%）。
    *   **對接邏輯**：使用標準化 ID 與 Inventory Key 實施 A/AX 庫存智慧分攤。
2.  **`ExportEngine` (美化匯出 - export_cleaned_plan.py)**：
    *   **自動生成**：每次同步自動產生美化版 Excel。
    *   **旗艦設計**：深藍專業表頭、凍結首行、中文字長度自適應欄寬。
3.  **`DashboardRenderer` (渲染層 - generate_dashboard_v3.py)**：
    *   **多重連動篩選**：整合「工程師」與「機台狀態（運行/停機/完工）」雙排控制鈕。
    *   **智慧搜尋**：支援機台號與 DTY 批號之即時模糊過濾。
    *   **視覺整合**：低侵入式「📝 備註」區塊設計。

## 2. UI/UX 規範
*   **導航優化**：右下角「回到頂部」按鈕、左上角「原始產銷(精簡版)」快捷開啟。
*   **佈局防禦**：
    *   **RWD 3.5**：手機端標籤區縮減至 110px，增加數據展示空間。
    *   **溢出保護**：全面實施 `overflow-wrap: break-word`，確保長備註與機台號不重疊。
*   **即時反饋**：狀態指示燈（閃爍綠、紅、灰）直觀呈現現場動態。

## 3. 自動化維護 (更新生產進度表.bat)
*   **一鍵同步**：Git Pull -> 資料清洗 -> 生成 Excel -> 生成 HTML -> Git Push。
*   **衝突預警**：若 Excel 檔案未關閉導致無法寫入，系統會彈出紅色警告並暫停。

## 4. 關鍵檔案備份 (v4.0 Final)
*   核心引擎：`reality_analyzer.py`
*   網頁渲染：`generate_dashboard_v3.py`
*   報表工具：`export_cleaned_plan.py`
*   自動指令：`更新生產進度表.bat`
*   輸出成果：`production_dashboard.html`
