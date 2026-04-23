@echo off
:: 設定 UTF-8 編碼以顯示中文
chcp 65001 > nul
title 產銷表更新系統 - 自動同步版

:: 取得批次檔所在目錄
set BASE_DIR=%~dp0
cd /d "%BASE_DIR%"

echo ==================================================
echo [準備] 正在處理本地變更並與 GitHub 同步...
echo --------------------------------------------------

:: 1. 先將本地所有變更 (包含腳本、網頁、GEMINI.md) 加入存檔區
git add .

:: 2. 提交一個臨時紀錄，確保工作區是乾淨的 (若無變更會自動跳過)
git commit -m "Save local changes before sync" 2>nul

:: 3. 抓取遠端變更 (您在網頁版刪除檔案的紀錄) 並合併
echo 正在抓取遠端更新...
git pull origin main --rebase

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [警報] 同步失敗！請檢查是否有手動修改導致的衝突。
    echo 嘗試執行: git rebase --abort 或手動處理衝突。
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [步驟 1/3] 正在執行 Python 資料處理...
echo --------------------------------------------------
"C:\Users\blaze\AppData\Local\Programs\Python\Python313\python.exe" "generate_dashboard.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [錯誤] generate_dashboard.py 執行失敗！
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [步驟 2/3] 正在同步最新產出至 GitHub...
echo --------------------------------------------------
:: 確保產出的網頁和更新的手冊也被加入
git add production_dashboard.html GEMINI.md
git commit -m "Auto Update: %date% %time%" 2>nul

echo 正在推送到 GitHub...
git push origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo [步驟 3/3] 全部更新成功！
    echo 線上網址: https://blaze0207.github.io/product-schedule/production_dashboard.html
    echo ==================================================
) else (
    echo.
    echo [失敗] GitHub 上傳失敗！
    pause
)

timeout /t 5
