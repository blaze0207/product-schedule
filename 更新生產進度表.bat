@echo off
:: 設定 UTF-8 編碼
chcp 65001 > nul
title 產銷表更新系統 - 強制同步修復版

:: 取得批次檔所在目錄
set BASE_DIR=%~dp0
cd /d "%BASE_DIR%"

echo ==================================================
echo [準備] 正在強制修復版本分歧並同步...
echo --------------------------------------------------

:: 1. 先中止任何可能卡住的 rebase 狀態
git rebase --abort >nul 2>&1

:: 2. 暫存本地目前的變更，確保 pull 不會失敗
echo 正在暫存本地變更...
git add .
git commit -m "Pre-sync save" 2>nul

:: 3. 執行強制拉取並合併 (不使用 rebase，降低衝突風險)
echo 正在與 GitHub 合併...
git pull origin main --no-rebase --no-edit

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [警報] 自動合併失敗！可能發生了檔案衝突。
    echo 請嘗試執行: git reset --hard origin/main (注意：這會覆蓋本地所有未上傳的修改)
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
