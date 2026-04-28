import pandas as pd
import os
import re

def analyze_grade_rules():
    file_name = '每日-最新庫存(DTY-LISA).xlsx'
    if not os.path.exists(file_name): return
    
    df = pd.read_excel(file_name, sheet_name='總庫存')
    
    results = []
    for _, row in df.iterrows():
        raw_batch = str(row.iloc[1]).strip().upper()
        if raw_batch == 'NAN' or raw_batch == '': continue
        
        # 移除 FD 前綴
        dty_batch = raw_batch[2:] if raw_batch.startswith('FD') else raw_batch
        
        grade = "Unknown"
        # 1. C級判斷 (最後一碼為 C)
        if dty_batch.endswith('C'):
            grade = "C"
        # 2. B級判斷 (第二碼為 G 且 最後一碼為 8)
        elif len(dty_batch) >= 2 and dty_batch[1] == 'G' and dty_batch.endswith('8'):
            grade = "B"
        # 3. AX級判斷 (第二碼為 G)
        elif len(dty_batch) >= 2 and dty_batch[1] == 'G':
            grade = "AX"
        # 4. A級判斷 (第二碼為 F)
        elif len(dty_batch) >= 2 and dty_batch[1] == 'F':
            grade = "A"
            
        results.append({
            'LISA_Batch': raw_batch,
            'DTY_Batch': dty_batch,
            'Determined_Grade': grade,
            'Weight': row.iloc[11] if not pd.isna(row.iloc[11]) else 0
        })
    
    res_df = pd.DataFrame(results)
    
    print("--- 規則判定結果統計 ---")
    print(res_df['Determined_Grade'].value_counts())
    
    print("\n--- 混合等級樣本 (以 2F04X2 為例) ---")
    # 搜尋可能屬於同一產品的紀錄 (忽略等級特徵碼)
    print(res_df[res_df['DTY_Batch'].str.contains('F04X2|G04X2')].to_string())

if __name__ == "__main__":
    analyze_grade_rules()
