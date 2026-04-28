import pandas as pd
import glob
import os
import re
from reality_analyzer import RealityLogAnalyzer

def debug_latest_data():
    files = [f for f in glob.glob("*撚二科生產資訊.xlsx") if not os.path.basename(f).startswith('~$')]
    if not files:
        print("Error: No Excel file found.")
        return
    
    file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
    print(f"=== 1. 原始檔案分析: {file_path} ===")
    
    df = pd.read_excel(file_path, sheet_name=1)
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
    latest_date = df.iloc[:, 0].max()
    print(f"最新日期: {latest_date}")

    # 針對問題機台 A22, T05, J07 等提取最新日期的原始資料
    target_machines = ['A10', 'A22', 'A24', 'J07', 'J09', 'T05']
    print("\n--- 最新日期的原始資料 (關鍵欄位) ---")
    # 欄位說明: 0:日期, 1:機台, 2:批號, 11:L側, 12:R側, 26:效率/狀態, 44:全/半, 45:全, 47:產量
    raw_latest = df[(df.iloc[:, 0] == latest_date) & (df.iloc[:, 1].isin(target_machines))]
    print(raw_latest.iloc[:, [0, 1, 2, 11, 12, 26, 44, 45, 47]])

    print("\n=== 2. 程式清洗後數據分析 (RealityLogAnalyzer) ===")
    analyzer = RealityLogAnalyzer(os.getcwd())
    cleaned_tasks = analyzer.get_reality_tasks()
    
    # 找出對應機台的清洗結果
    print("\n--- 清洗後的任務狀態 ---")
    for t in cleaned_tasks:
        if t['machine'] in target_machines:
            status = "活動中" if t['is_active'] else "歷史"
            print(f"機台: {t['machine']:3s} | 批號: {t['dty_batch']:8s} | 側邊: {t['produced_sides']} | 狀態: {status} | 天數: {t['days']} | 最後狀態: {t['last_status']}")

if __name__ == "__main__":
    debug_latest_data()
