import pandas as pd
import os

def analyze_grade_v2():
    file_name = '每日-最新庫存(DTY-LISA).xlsx'
    if not os.path.exists(file_name): return
    df = pd.read_excel(file_name, sheet_name='總庫存')
    
    results = []
    for _, row in df.iterrows():
        raw_batch = str(row.iloc[1]).strip().upper()
        if raw_batch == 'NAN' or raw_batch == '': continue
        b = raw_batch[2:] if raw_batch.startswith('FD') else raw_batch
        
        # [規則 1] C級: (2G 或 8G 開頭) 且 結尾為 C
        if (b.startswith('2G') or b.startswith('8G')) and b.endswith('C'):
            grade = "C"
            join_key = b[2:-1] if b.startswith('2G') or b.startswith('8G') else b
        # [規則 2] B級: 2G 開頭 且 結尾為 8
        elif b.startswith('2G') and b.endswith('8'):
            grade = "B"
            join_key = b[2:-1]
        # [規則 3] AX級: 2G 開頭 (且排除上述)
        elif b.startswith('2G'):
            grade = "AX"
            join_key = b[2:]
        # [規則 4] A級: 2F 開頭
        elif b.startswith('2F'):
            grade = "A"
            join_key = b[2:]
        else:
            grade = "Other"
            join_key = b

        results.append({'Raw': raw_batch, 'DTY': b, 'Grade': grade, 'JoinKey': join_key})
    
    res_df = pd.DataFrame(results)
    print("--- 修正版規則判定統計 ---")
    print(res_df['Grade'].value_counts())
    
    print("\n--- A級特殊結尾核實 (應判定為 A) ---")
    print(res_df[(res_df['Grade'] == 'A') & (res_df['DTY'].str.endswith('C'))].head(5))

    print("\n--- 產品族群化對接測試 (以 04X2 為核心) ---")
    print(res_df[res_df['JoinKey'] == '04X2'].sort_values(by='Grade'))

if __name__ == "__main__":
    analyze_grade_v2()
