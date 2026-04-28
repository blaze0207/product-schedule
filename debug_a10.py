import pandas as pd
import glob
import os

files = glob.glob("*撚二科生產資訊.xlsx")
if not files:
    print("找不到檔案")
else:
    file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
    print(f"正在檢查檔案: {file_path}")
    xl = pd.ExcelFile(file_path)
    df = xl.parse(xl.sheet_names[1])
    
    # 搜尋 A10 在 2026-04-23 的資料 (假設日期在 Index 0, 機台在 Index 1)
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
    target_date = pd.to_datetime('2026-04-23')
    
    a10_data = df[(df.iloc[:, 1].astype(str).str.contains('A10')) & (df.iloc[:, 0] == target_date)]
    
    if a10_data.empty:
        print("在 4/23 找不到 A10 的資料")
    else:
        print("\n--- A10 在 4/23 的原始資料 ---")
        # 顯示關鍵欄位: 日期(0), 機台(1), 批號(2), 規格(4), 
        # L(11), M(12), AA(26), AT(45), AV(47)
        cols_to_show = [0, 1, 2, 4, 11, 12, 26, 45, 47]
        print(a10_data.iloc[:, cols_to_show].to_string())
