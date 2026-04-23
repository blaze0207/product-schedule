@echo off
chcp 65001 > nul
title 產銷表更新 - GitHub 自動同步版
echo --------------------------------------------------
echo [1/3] 正在更新產銷表數據...
cd /d C:\product_schedule_test

:: 確保 Git 記住憑證
git config --global credential.helper wincred

C:\Users\blaze\AppData\Local\Programs\Python\Python313\python.exe generate_dashboard.py

if %ERRORLEVEL% NEQ 0 (
    echo [錯誤] 資料產出失敗！
    pause
    exit
)

echo [2/3] 正在同步至 GitHub (blaze0207/product-schedule)...
git add production_dashboard.html
git commit -m "Auto Update: %date% %time%"
git push origin main

if %ERRORLEVEL% EQU 0 (
    echo [3/3] 成功！請查看: https://blaze0207.github.io/product-schedule/production_dashboard.html
    echo --------------------------------------------------
) else (
    echo [失敗] 上傳失敗！請先手動執行一次 git push 以驗證帳密。
    pause
)

timeout /t 5
