import pandas as pd
import os

def inspect_stock():
    file_name = '每日-最新庫存(DTY-LISA).xlsx'
    if not os.path.exists(file_name):
        print(f"Error: {file_name} not found.")
        return
    
    # 讀取總庫存
    df = pd.read_excel(file_name, sheet_name='總庫存')
    
    # 欄位名稱處理 (避免亂碼影響)
    # B: 加工批號, C: 規格, D: 等級, L: 結存重量, M: 結存箱數
    # 根據之前的診斷，我們鎖定物理位置
    
    print("--- 數據樣本 (批號 vs 等級) ---")
    # 挑選有內容的行
    valid_df = df.dropna(subset=[df.columns[1], df.columns[3]]).head(100)
    
    # 打印前 50 筆
    print(valid_df.iloc[:, [1, 2, 3, 11, 12]].to_string())

    # 深入分析：尋找特定批號的等級分佈
    # 例如尋找包含 'FD2F01' 的所有紀錄
    search_keyword = 'FD2F'
    print(f"\n--- 搜尋包含 '{search_keyword}' 的批號等級分佈 ---")
    matches = df[df.iloc[:, 1].astype(str).str.contains(search_keyword, na=False)]
    
    # 依批號排序方便觀察
    matches_sorted = matches.sort_values(by=df.columns[1])
    print(matches_sorted.iloc[:, [1, 2, 3, 11, 12]].head(30).to_string())

if __name__ == "__main__":
    inspect_stock()
