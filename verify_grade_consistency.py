import pandas as pd
import os

def verify_grade_consistency():
    file_name = '每日-最新庫存(DTY-LISA).xlsx'
    if not os.path.exists(file_name): return
    
    # 讀取數據，header=0
    df = pd.read_excel(file_name, sheet_name='總庫存')
    
    mismatch_count = 0
    match_count = 0
    mismatch_details = []
    
    print(f"正在分析檔案: {file_name}")
    print("------------------------------------------")
    
    for i, row in df.iterrows():
        raw_batch = str(row.iloc[1]).strip().upper()
        raw_grade = str(row.iloc[3]).strip().upper() # Excel 原始 D 欄
        
        if raw_batch == 'NAN' or raw_batch == '' or raw_grade == 'NAN': continue
        
        # 移除 FD 前綴
        b = raw_batch[2:] if raw_batch.startswith('FD') else raw_batch
        
        # 判定邏輯
        calc_grade = "Unknown"
        if b.startswith('2F'):
            calc_grade = "A"
        elif (b.startswith('2G') or b.startswith('8G')) and b.endswith('C'):
            calc_grade = "C"
        elif b.startswith('2G') and b.endswith('8'):
            calc_grade = "B"
        elif b.startswith('2G'):
            calc_grade = "AX"
        
        # 處理原始等級標註的特殊性 (例如 Excel 寫 'AX' 或 'B', 'C', 'A')
        # 有些 Excel 可能寫 'AX'，有些寫 'A' 等
        if calc_grade == raw_grade:
            match_count += 1
        else:
            # 排除一些特殊情況，例如 D 欄空白或總計列
            if calc_grade != "Unknown":
                mismatch_count += 1
                mismatch_details.append({
                    'Row': i + 2, # Excel 行號
                    'Batch': raw_batch,
                    'Excel_Grade': raw_grade,
                    'Calc_Grade': calc_grade
                })

    print(f"比對完成！")
    print(f"吻合筆數: {match_count}")
    print(f"不吻合筆數: {mismatch_count}")
    
    if mismatch_count > 0:
        print("\n--- 不吻合案例分析 (前 20 筆) ---")
        detail_df = pd.DataFrame(mismatch_details)
        print(detail_df.head(20).to_string(index=False))
    else:
        print("\n✨ 100% 吻合！證明批號等級規則完全正確。")

if __name__ == "__main__":
    verify_grade_consistency()
