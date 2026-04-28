import pandas as pd
import glob
import os

def verify_precision():
    files = [f for f in glob.glob("*撚二科生產資訊.xlsx") if not os.path.basename(f).startswith('~$')]
    if not files: return
    file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
    
    # 讀取資料
    df = pd.read_excel(file_path, sheet_name=1)
    cols = list(df.columns)
    
    # 定位最新日期與目標機台
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
    latest_date = df.iloc[:, 0].max()
    target_machines = ['A15', 'A19']
    target_data = df[(df.iloc[:, 0] == latest_date) & (df.iloc[:, 1].isin(target_machines))]
    
    print(f"分析檔案: {file_path}")
    print(f"最新日期: {latest_date}")
    
    print("\n=== [核實 1] 關鍵欄位索引對應表 ===")
    # 我們特別關注 11, 12, 26, 44, 47 這幾個點
    key_indices = [2, 11, 12, 26, 44, 45, 47]
    for i in range(min(50, len(cols))):
        marker = " <--- 重要" if i in key_indices else ""
        print(f"Index {i:2d}: {cols[i]}{marker}")

    print("\n=== [核實 2] A15/A19 原始數據真相 ===")
    for _, row in target_data.iterrows():
        print(f"\n--- 機台: {row.iloc[1]} ---")
        for i in key_indices:
            if i < len(cols):
                print(f"[{i:2d}] {cols[i]}: {row.iloc[i]}")
        # 額外搜尋是否有「全」或「半」出現在這一行
        for i, val in enumerate(row):
            if str(val) in ['全', '半']:
                print(f"!! 發現關鍵字 '{val}' 位於 Index {i} ({cols[i]})")

if __name__ == "__main__":
    verify_precision()
