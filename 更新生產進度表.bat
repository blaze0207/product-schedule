@echo off
:: 設定 UTF-8 編碼
chcp 65001 > nul
title 撚二科產銷戰情室 v4.0 - 自動化同步系統 (偵錯強化版)

:: 取得批次檔所在目錄
set BASE_DIR=%~dp0
cd /d "%BASE_DIR%"

echo ==================================================
echo [步驟 1] 正在從 GitHub 同步遠端狀態...
echo --------------------------------------------------
git pull origin main
if %ERRORLEVEL% NEQ 0 (
    echo [提示] Pull 過程有警示，可能是本地已有最新資料或需要手動處理衝突。
)
echo [完成] 同步步驟。
echo.

echo ==================================================
echo [步驟 2] 執行數據清洗與戰情室生成 (v4.0 旗艦版)
echo --------------------------------------------------
:: 執行核心腳本
python generate_dashboard_v3.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ [錯誤] 數據處理失敗！請檢查 Excel 是否已關閉。
    pause
    goto :end
)
echo [完成] 網頁與美化版 Excel 已生成。
echo.

echo ==================================================
echo [步驟 3] 正在上傳至 GitHub...
echo --------------------------------------------------
git add .

:: 檢查是否有東西需要提交
git diff --cached --exit-code > nul
if %ERRORLEVEL% EQU 0 (
    echo [資訊] 偵測到內容無變動，跳過提交。
) else (
    echo [資訊] 偵測到更新，正在提交變更...
    git commit -m "Update: v4.0 Flagship Version (Auto-sync)"
)

echo [資訊] 正在執行 Push...
git push origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✨ [成功] 數據已全面同步！
    echo --------------------------------------------------
    echo 網頁網址: https://blaze0207.github.io/product-schedule/production_dashboard.html
    echo ==================================================
    :: 成功後等待 10 秒才關閉，讓您看清楚結果
    timeout /t 10
) else (
    echo.
    echo ❌ [失敗] GitHub 上傳失敗！可能是網路問題或權限問題。
    echo 請查看上方 Git 錯誤訊息。
    pause
)

:end
