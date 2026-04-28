import pandas as pd
import glob
import os

def final_verify():
    files = [f for f in glob.glob("*撚二科生產資訊.xlsx") if not os.path.basename(f).startswith('~$')]
    if not files: return
    file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
    
    # 讀取資料，不設 header 來看前幾行，確認標題位置
    df_raw = pd.read_excel(file_path, sheet_name=1, header=None)
    
    # 尋找標題列（通常在第 0 或 1 行）
    header_row = 0
    for r in range(3):
        row_vals = [str(v) for v in df_raw.iloc[r, :]]
        if '機台' in row_vals or '批號' in row_vals or '單/雙' in row_vals:
            header_row = r
            break
            
    # 重新以正確的 header 讀取
    df = pd.read_excel(file_path, sheet_name=1, header=header_row)
    cols = list(df.columns)
    
    # 定位關鍵欄位
    def find_col(name):
        for i, c in enumerate(cols):
            if name in str(c): return i
        return -1

    idx_machine = find_col('機台')
    idx_batch = find_col('批號')
    idx_l = find_col('L側')
    idx_r = find_col('R側')
    idx_eff = find_col('效率')
    idx_single_double = find_col('單/雙')
    idx_unit = find_col('台')
    idx_prod = find_col('產量')

    print(f"檔案: {file_path} | 標題列位於: {header_row}")
    print(f"--- 欄位定位結果 ---")
    print(f"機台: {idx_machine}, 批號: {idx_batch}, L側: {idx_l}, R側: {idx_r}")
    print(f"效率: {idx_eff}, 單/雙: {idx_single_double}, 台: {idx_unit}, 產量: {idx_prod}")

    # 讀取 A15/A19 最新數據
    df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
    latest_date = df.iloc[:, 0].max()
    target_data = df[(df.iloc[:, idx_machine].isin(['A15', 'A19'])) & (df.iloc[:, 0] == latest_date)]

    print(f"\n--- [100% 正確數據核實] 最新日期: {latest_date} ---")
    for _, row in target_data.iterrows():
        print(f"\n機台: {row.iloc[idx_machine]}")
        print(f"  批號: {row.iloc[idx_batch]}")
        print(f"  L側內容: {row.iloc[idx_l]}")
        print(f"  R側內容: {row.iloc[idx_r]}")
        print(f"  效率(26): {row.iloc[idx_eff]}")
        print(f"  單/雙(44): {row.iloc[idx_single_double]}")
        print(f"  台(45): {row.iloc[idx_unit]}")
        print(f"  產量(47): {row.iloc[idx_prod]}")

if __name__ == "__main__":
    final_verify()
