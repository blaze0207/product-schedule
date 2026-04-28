import pandas as pd
import glob
import os
import re

# 1. 定義最強標準化函數 (目標: 提取核心代碼，過濾前綴/尾碼/裝飾字)
def standardize_id(val):
    if pd.isna(val): return ""
    s = str(val).strip().upper()
    
    # 移除關鍵裝飾字
    s = s.replace("LIKE", "").replace("R", "").replace("K", "").strip()
    
    # 處理 DTY 批號 (例如: FD2FA032 -> 2FA032)
    # 規則: 若 FD 後面接數字，則保留該數字開始的部分
    if s.startswith("FD"):
        match = re.search(r'\d', s)
        if match: s = s[match.start():]
        else: s = s[2:]
        
    # 處理 POY 批號 (例如: FP81386 -> 81386)
    if s.startswith("FP"):
        match = re.search(r'\d', s)
        if match: s = s[match.start():]
        else: s = s[2:]
        
    # 截取前 6 位作為核心匹配碼 (依實務經驗，前 6 碼通常最精確)
    return s[:6]

def run_alignment_test():
    print("=== 數據對接精確度診斷 ===")
    
    # A. 讀取 DTY 計畫母表
    dty_files = [f for f in glob.glob("DTY*.xlsx") if not f.startswith('~$')]
    if not dty_files: return print("Missing DTY file")
    dty_file = sorted(dty_files, key=os.path.getmtime, reverse=True)[0]
    df_dty = pd.read_excel(dty_file, sheet_name=0, header=0)
    # 取前二欄，補全機台
    df_dty.iloc[:, 0] = df_dty.iloc[:, 0].ffill()
    dty_batches = df_dty.iloc[:, 1].dropna().unique()
    
    # B. 讀取實產表
    prod_files = [f for f in glob.glob("*撚二科生產資訊.xlsx") if not f.startswith('~$')]
    prod_batches = []
    if prod_files:
        df_prod = pd.read_excel(prod_files[0], sheet_name=1)
        prod_batches = df_prod.iloc[:, 2].dropna().unique()
        
    # C. 讀取 POY 庫存
    poy_files = [f for f in glob.glob("絲八科-庫存表*.xlsx") if not f.startswith('~$')]
    poy_batches = []
    if poy_files:
        df_poy = pd.read_excel(poy_files[0], sheet_name=-1)
        poy_batches = df_poy.iloc[:, 0].dropna().unique()

    print(f"\n[1] 母表 (DTY) 原始批號範例: {list(dty_batches[:5])}")
    print(f"[2] 實產表原始批號範例: {list(prod_batches[:5])}")
    print(f"[3] POY 庫存原始批號範例: {list(poy_batches[:5])}")

    print("\n--- 開始執行標準化對接測試 ---")
    dty_std = {standardize_id(b): b for b in dty_batches}
    prod_std = {standardize_id(b) for b in prod_batches}
    poy_std = {standardize_id(b) for b in poy_batches}
    
    success_prod, success_poy = 0, 0
    fail_prod = []
    
    for std_id, orig in dty_std.items():
        if std_id in prod_std: success_prod += 1
        else: fail_prod.append(f"{orig} -> {std_id}")
            
        if any(std_id in p_id for p_id in poy_std) or std_id[:5] in "".join(poy_std): success_poy += 1

    total = len(dty_std)
    print(f"\n對接結果總結 (母表總批號數: {total})")
    print(f"✅ 實產表對接成功: {success_prod} ({success_prod/total*100:.1f}%)")
    print(f"✅ POY 庫存對接成功: {success_poy} ({success_poy/total*100:.1f}%)")
    
    if fail_prod:
        print("\n❌ 實產表對接失敗的批號 (前 10 筆):")
        for f in fail_prod[:10]: print(f"  - {f}")

if __name__ == "__main__":
    run_alignment_test()
