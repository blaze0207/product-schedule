
import pandas as pd
import re
import os
from reality_analyzer import RealityLogAnalyzer

analyzer = RealityLogAnalyzer(os.getcwd())
result = analyzer.get_reality_tasks()

print(f"使用的狀態基準日: {result['production_date']}")

tasks = result['tasks']
active_tasks = [t for t in tasks if t['is_active']]

run_count = 0
running_details = []

for d in active_tasks:
    machine = d['machine']
    dty_batch = d['dty_batch']
    last_status = d.get('last_status', "")
    current_sides = d.get('current_sides', [])

    # S1 邏輯
    if machine == 'S1':
        weight = 0
        match = re.search(r'(\d+)SEC', last_status, re.I)
        if match:
            sec = int(match.group(1))
            weight = 1.0 if sec > 6 else 0.5
        elif dty_batch != '---':
            weight = 1.0
        
        if weight > 0:
            run_count += weight
            running_details.append(f"{machine}: {weight} 台 (狀態: {last_status})")
        continue

    # S2 邏輯 (排除)
    if machine == 'S2':
        continue

    # 一般機台邏輯
    if dty_batch != '---':
        weight = 1.0 if len(current_sides) >= 2 else 0.5
        run_count += weight
        running_details.append(f"{machine}: {weight} 台 (側邊: {'/'.join(current_sides)})")

print(f"\n總計運行台數: {run_count}")
print("機台明細:")
for detail in sorted(running_details):
    print(f"  - {detail}")
