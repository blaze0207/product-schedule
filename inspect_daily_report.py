
import pandas as pd
import glob
import os

def analyze_daily_report():
    # 動態獲取最新的生產日報表
    files = [f for f in glob.glob("生產日報表(*).xlsx") if not os.path.basename(f).startswith('~$')]
    if not files:
        print("錯誤：找不到生產日報表檔案")
        return
    
    target_file = sorted(files, key=os.path.getmtime, reverse=True)[0]
    print(f"正在讀取檔案: {target_file}")
    
    xl = pd.ExcelFile(target_file)
    if '日報表' not in xl.sheet_names:
        print(f"錯誤：在檔案中找不到 '日報表' 分頁。可用分頁為: {xl.sheet_names}")
        return
    
    # 讀取日報表內容
    df = xl.parse('日報表')
    
    print("\n--- 欄位清單 ---")
    print(df.columns.tolist())
    
    print("\n--- 資料前 10 列樣貌 ---")
    print(df.head(10).to_string())
    
    print("\n--- 資料型態分析 ---")
    print(df.info())

if __name__ == "__main__":
    analyze_daily_report()
