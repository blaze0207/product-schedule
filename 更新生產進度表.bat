@echo off
chcp 65001 > nul
title 產銷表更新 - 自動雲端同步版
echo --------------------------------------------------
echo [1/3] 正在執行 Python 更新資料...
cd /d C:\product_schedule_test
C:\Users\blaze\AppData\Local\Programs\Python\Python313\python.exe generate_dashboard.py

if %ERRORLEVEL% NEQ 0 (
    echo [錯誤] Python 執行失敗，請檢查 Excel！
    pause
    exit
)

echo [2/3] 正在同步到 GitHub 雲端...
:: 執行 Git 同步指令
git add production_dashboard.html
git commit -m "自動更新產銷表: %date% %time%"
git push origin main

if %ERRORLEVEL% EQU 0 (
    echo [3/3] 成功！雲端網頁已更新。
    echo --------------------------------------------------
    echo 雲端網址: https://github.com/您的帳號/product-schedule/
    echo (請將您的帳號填入上方路徑)
    start production_dashboard.html
) else (
    echo [失敗] 無法上傳到 GitHub，請檢查網路或 Git 設定。
    pause
)

timeout /t 5
