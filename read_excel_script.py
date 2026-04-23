import pandas as pd
import os
import sys

# 設定輸出編碼為 utf-8 以防萬一
sys.stdout.reconfigure(encoding='utf-8')

target_dir = r'C:\product_schedule_test'
file_name = 'DTYMay26產銷260416-1450+300+30.xlsx'
file_path = os.path.join(target_dir, file_name)

output_file = os.path.join(target_dir, 'analysis_result.txt')

with open(output_file, 'w', encoding='utf-8') as f:
    try:
        df = pd.read_excel(file_path, sheet_name='May')
        f.write('--- Columns ---\n')
        f.write(str(df.columns.tolist()) + '\n\n')
        f.write('--- Data Preview (First 50 rows) ---\n')
        f.write(df.head(50).to_string())
        print(f'Analysis completed. Results saved to {output_file}')
    except Exception as e:
        f.write(f'Error: {e}')
        print(f'Error occurred: {e}')
