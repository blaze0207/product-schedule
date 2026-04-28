
import os
import pandas as pd
import glob
from reality_analyzer import RealityLogAnalyzer

analyzer = RealityLogAnalyzer(os.getcwd())

print("--- 1. 檢查 LISA 庫存原始資料 (2F29B1) ---")
lisa_file = os.path.join(os.getcwd(), "每日-最新庫存(DTY-LISA).xlsx")
if os.path.exists(lisa_file):
    df_lisa = pd.read_excel(lisa_file, sheet_name='總庫存')
    # 搜尋包含 2F29B1 的批號
    match_lisa = df_lisa[df_lisa.iloc[:, 1].astype(str).str.contains('2F29B1', na=False)]
    print(f"LISA 中找到 {len(match_lisa)} 筆相關紀錄:")
    for _, row in match_lisa.iterrows():
        print(f"  批號: {row.iloc[1]} | 重量: {row.iloc[11]} | 標準化後: {analyzer.standardize_id(row.iloc[1][2:] if str(row.iloc[1]).startswith('FD') else row.iloc[1])}")
else:
    print("找不到 LISA 庫存檔")

print("\n--- 2. 檢查生產資訊原始資料 (A21 + 2F29B1) ---")
prod_files = glob.glob(os.path.join(os.getcwd(), "*撚二科生產資訊.xlsx"))
if prod_files:
    latest_prod = sorted(prod_files, key=os.path.getmtime, reverse=True)[0]
    df_prod = pd.read_excel(latest_prod, sheet_name=1)
    match_prod = df_prod[(df_prod.iloc[:, 1].astype(str).str.contains('A21', na=False)) & 
                         (df_prod.iloc[:, 2].astype(str).str.contains('2F29B1', na=False))]
    print(f"生產資訊中找到 {len(match_prod)} 筆相關紀錄:")
    for _, row in match_prod.iterrows():
        print(f"  日期: {row.iloc[0]} | 機台: {row.iloc[1]} | 批號: {row.iloc[2]} | 標準化後: {analyzer.standardize_id(row.iloc[2])}")
else:
    print("找不到生產資訊檔")

print("\n--- 3. 檢查引擎對接結果 ---")
result = analyzer.get_reality_tasks()
found = False
for t in result['tasks']:
    if 'A21' in t['machine'] and '2F29B1' in t['dty_batch']:
        print(f"引擎任務中找到: M:{t['machine']} B:{t['dty_batch']} (Std:{t['dty_std']})")
        print(f"  繳庫量: {t['stock_summary']['total_deposit']} | 計畫目標: {t['target_kg']}")
        found = True
if not found:
    print("引擎最終任務清單中找不到 A21-2F29B1")
