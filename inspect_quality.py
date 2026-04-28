
import pandas as pd
import glob
import os
import re

def analyze_quality_metrics():
    files = [f for f in glob.glob("生產日報表(*).xlsx") if not os.path.basename(f).startswith('~$')]
    if not files: return
    target_file = sorted(files, key=os.path.getmtime, reverse=True)[0]
    
    df = pd.read_excel(target_file, sheet_name='日報表', skiprows=2)
    
    quality_data = []
    for _, row in df.iterrows():
        machine_raw = str(row.iloc[15]).strip()
        if machine_raw == 'nan' or not machine_raw: continue
        
        # 提取 A級率 (Index 10) 與 滿管率 (Index 13)
        a_rate = pd.to_numeric(row.iloc[10], errors='coerce')
        full_spool_rate = pd.to_numeric(row.iloc[13], errors='coerce')
        
        if pd.isna(a_rate) and pd.isna(full_spool_rate): continue
        
        quality_data.append({
            'machine_tag': machine_raw,
            'batch': str(row.iloc[0]),
            'a_rate': round(a_rate * 100, 2) if not pd.isna(a_rate) else None,
            'full_rate': round(full_spool_rate * 100, 2) if not pd.isna(full_spool_rate) else None
        })
    
    print(f"--- 數據採樣 (自 {target_file}) ---")
    for item in quality_data[:10]:
        print(f"機台: {item['machine_tag']:<6} | 批號: {item['batch']:<8} | A級率: {item['a_rate']}% | 滿管率: {item['full_rate']}%")

if __name__ == "__main__":
    analyze_quality_metrics()
