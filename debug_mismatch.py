
import os
import pandas as pd
from reality_analyzer import RealityLogAnalyzer

analyzer = RealityLogAnalyzer(os.getcwd())
plan_map, target_aggregate = analyzer.get_plan_data()
result = analyzer.get_reality_tasks()

print("--- 診斷：有繳庫量但無目標計畫 ( target_kg=0 ) ---")
for t in result['tasks']:
    if t['target_kg'] == 0 and t['stock_summary']['total_deposit'] > 0:
        print(f"機台: {t['machine']} | 批號: {t['dty_batch']} | 標準化批號: {t['dty_std']}")
        
print("\n--- 檢查產銷表中的 Key 範例 ---")
keys = list(target_aggregate.keys())[:10]
for k in keys:
    print(f"計畫 Key: {k} | 目標值: {target_aggregate[k]}")
