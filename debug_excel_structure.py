import pandas as pd
import glob
import os

def debug_excel():
    files = [f for f in glob.glob("*撚二科生產資訊.xlsx") if not os.path.basename(f).startswith('~$')]
    if not files:
        print("Error: No Excel file found.")
        return
    
    file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
    print(f"Analyzing file: {file_path}")
    
    # 讀取 Sheet 2 (Index 1)
    df = pd.read_excel(file_path, sheet_name=1)
    
    print("\n--- Column Mapping (First 60 columns) ---")
    for i, col in enumerate(df.columns[:60]):
        # 顯示索引、欄位名稱，以及該欄位第一筆非空值的內容（作為參考）
        sample_val = df.iloc[:, i].dropna().iloc[0] if not df.iloc[:, i].dropna().empty else "Empty"
        print(f"Index {i:2d}: {col} (Sample: {sample_val})")

    # 搜尋關鍵字位置
    print("\n--- Searching for Keyword '全' in all columns ---")
    for i in range(len(df.columns)):
        if df.iloc[:, i].astype(str).str.contains('全').any():
            print(f"Found '全' in Column Index {i} ({df.columns[i]})")

    # 針對 A22, T05 進行具體行數據檢查
    print("\n--- Detail Check for A22 and T05 (Latest Records) ---")
    machines = ['A22', 'T05']
    df_filtered = df[df.iloc[:, 1].astype(str).str.contains('|'.join(machines), na=False)].tail(5)
    print(df_filtered.iloc[:, [0, 1, 2, 11, 12, 26, 44, 47]]) # 暫定索引

if __name__ == "__main__":
    debug_excel()
