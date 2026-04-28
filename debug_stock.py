import pandas as pd
import os
import glob
from core_processor import MasterTableEngine

def run_debug():
    eng = MasterTableEngine(os.getcwd())
    l_file = '每日-最新庫存(DTY-LISA).xlsx'
    if not os.path.exists(l_file):
        print("Missing LISA file")
        return
        
    df_lisa = pd.read_excel(l_file, header=None)
    df_lisa = df_lisa.iloc[:-1] # 排除最後一列
    
    # 1. 取得計畫母表中的強對接 Key
    tasks = eng.get_planned_tasks()
    task_keys = {t['batch_join']: t['batch_orig'] for t in tasks}
    
    # 2. 追蹤每一筆庫存
    total_raw = 0
    total_linked = 0
    missed_list = []
    
    for _, row in df_lisa.iterrows():
        raw_b = str(row[1]).strip().upper()
        qty = pd.to_numeric(row[11], errors='coerce') or 0
        total_raw += qty
        
        # 使用對接模組的清洗規則
        bj = eng.standardize_batch(raw_b, aggressive=True)
        
        if bj in task_keys:
            total_linked += qty
        else:
            if qty > 0:
                missed_list.append((raw_b, bj, qty))

    print("\n" + "="*30)
    print(f"LISA 總表原始公斤數: {total_raw:,.1f}")
    print(f"對接到計畫表的公斤數: {total_linked:,.1f}")
    print(f"遺失數據量: {total_raw - total_linked:,.1f}")
    print(f"遺失比例: {(total_raw - total_linked)/total_raw*100:.1f}%")
    print("="*30)
    
    print("\n--- 遺失庫存前 10 大批號 ---")
    for b, bj, q in sorted(missed_list, key=lambda x: x[2], reverse=True)[:10]:
        print(f"批號: {b:12} | Key: {bj:10} | 重量: {q:8,.0f}")

if __name__ == "__main__":
    run_debug()
