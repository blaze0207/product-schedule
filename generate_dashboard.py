import pandas as pd
import os
import json
import glob
from datetime import datetime

# 固定工作目錄
os.chdir(r'C:\product_schedule_test')

def simplify_id(val):
    s = str(val).strip().upper()
    if s.startswith("FD2F"): s = s[4:]
    return s

def clean_data():
    files = [f for f in glob.glob("DTYMay26*.xlsx") if not f.startswith('~$')]
    if not files: raise FileNotFoundError("找不到 Excel 檔案")
    
    files.sort(key=os.path.getmtime, reverse=True)
    excel_file = files[0]
    mtime = datetime.fromtimestamp(os.path.getmtime(excel_file)).strftime('%Y-%m-%d %H:%M:%S')
    
    xl = pd.ExcelFile(excel_file)
    df_may = xl.parse(xl.sheet_names[0], header=0)
    df_stock = xl.parse(xl.sheet_names[1], header=None)
    
    stock_map = {}
    valid_grades = ['A', 'AX', 'B', 'C']
    for _, row in df_stock.iterrows():
        bid = simplify_id(row[1])
        grade = str(row[3]).strip().upper()
        qty = pd.to_numeric(row[11], errors='coerce')
        if pd.isna(qty) or qty == 0: continue
        if bid not in stock_map:
            stock_map[bid] = {'total': 0, 'grades': {g: 0 for g in valid_grades}, 'Other': 0}
        stock_map[bid]['total'] += qty
        if grade in valid_grades: stock_map[bid]['grades'][grade] += qty
        else: stock_map[bid]['Other'] += qty

    cols = ['machine', 'batch_no', 'desc', 'mark', 'dty_spec', 'date_range', 'scheduled_days', 't_d', 'poy_batch', 'poy_spec']
    df_may = df_may.iloc[:, :len(cols)]
    df_may.columns = cols
    df_may['machine'] = df_may['machine'].ffill()
    df_may = df_may.dropna(subset=['batch_no', 'dty_spec'], how='all')
    
    output_data = []
    for _, row in df_may.iterrows():
        m_name = str(row['machine']).strip()
        m_name_upper = m_name.upper()
        if m_name_upper in ['V', 'S2'] or 'S2' in m_name_upper: continue
        if '庫存不排產' in m_name or '待排產' in m_name or 'M4 (CW)' in m_name_upper or m_name == '機台': continue
        
        raw_days = row['scheduled_days']
        if pd.isna(raw_days): continue
        days = pd.to_numeric(raw_days, errors='coerce')
        if pd.isna(days) or days <= 0: continue
        
        td = pd.to_numeric(row['t_d'], errors='coerce') or 0
        target_kg = days * td * 1000
        bid = simplify_id(row['batch_no'])
        
        s_data = stock_map.get(bid, {'total': 0, 'grades': {}})
        stored = s_data['total']
        grades = {g: q for g, q in s_data.get('grades', {}).items() if q > 0}
        if s_data.get('Other', 0) > 0: grades['Other'] = s_data['Other']
        
        demand = stored - target_kg
        pct = min(100, (stored / target_kg * 100)) if target_kg > 0 else 0
        
        # 處理日期範圍文字 (若為 datetime 則轉為字串)
        dr = row['date_range']
        if isinstance(dr, datetime):
            dr = dr.strftime('%m/%d')
        else:
            dr = str(dr).replace(' 00:00:00', '')

        output_data.append({
            'machine': m_name,
            'batch': str(row['batch_no']),
            'spec': str(row['dty_spec']),
            'date_range': dr, # 新增日期欄位
            'days': days,
            'td': td,
            'target': target_kg,
            'stored': stored,
            'grades': grades,
            'demand': demand,
            'pct': pct,
            'poy': str(row['poy_batch'])
        })

    return {
        'data': output_data,
        'update_time': datetime.now().strftime('%H:%M:%S'),
        'file_name': os.path.basename(excel_file),
        'file_time': mtime
    }

def generate_html(res):
    json_data = json.dumps(res['data'], ensure_ascii=False)
    output_html = os.path.join(r'C:\product_schedule_test', 'production_dashboard.html')
    html = f"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>5月產銷表</title>
    <style>
        :root {{ --primary: #2563eb; --bg: #f8fafc; --text: #1e293b; }}
        body {{ font-family: "Microsoft JhengHei", sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 10px; }}
        .header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; flex-wrap: wrap; gap: 10px; }}
        h1 {{ font-size: 1.5rem; margin: 0; }}
        .info-bar {{ font-size: 11px; color: #64748b; margin-bottom: 15px; background: #fff; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; line-height: 1.6; border-left: 4px solid #2563eb; }}
        .search-box {{ padding: 12px; width: 100%; max-width: 300px; border: 1px solid #ddd; border-radius: 8px; font-size: 16px; box-sizing: border-box; }}
        .filter-group {{ margin-bottom: 15px; display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }}
        .btn {{ padding: 10px 16px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; transition: 0.2s; font-size: 14px; }}
        .btn-default {{ background: #e2e8f0; color: #475569; }}
        .btn-primary {{ background: #2563eb; color: white; }}
        .card {{ background: white; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); padding: 10px; overflow-x: auto; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ background: #f1f5f9; padding: 12px; text-align: left; border-bottom: 2px solid #e2e8f0; font-size: 13px; color: #64748b; }}
        td {{ padding: 12px; border-bottom: 1px solid #f1f5f9; font-size: 14px; vertical-align: top; }}
        .badge {{ padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: bold; }}
        .machine-tag {{ background: #dbeafe; color: #1e40af; }}
        .grade-tag {{ display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; margin-right: 4px; margin-bottom: 4px; color: white; }}
        .grade-A {{ background: #22c55e; }} .grade-AX {{ background: #3b82f6; }} .grade-B {{ background: #f59e0b; }} .grade-C {{ background: #64748b; }} .grade-Other {{ background: #94a3b8; }}
        .progress-bar-container {{ width: 100%; max-width: 80px; height: 8px; background: #e2e8f0; border-radius: 4px; display: inline-block; overflow: hidden; vertical-align: middle; }}
        .progress-inner {{ height: 100%; background: #22c55e; }}
        .status-negative {{ color: #ef4444; font-weight: bold; }}
        .status-positive {{ color: #22c55e; font-weight: bold; }}
        @media screen and (max-width: 768px) {{
            thead {{ display: none; }}
            tr {{ display: block; margin-bottom: 15px; border: 1px solid #e2e8f0; border-radius: 12px; padding: 10px; background: #fff; }}
            td {{ display: block; border: none; border-bottom: 1px solid #f1f5f9; position: relative; padding-left: 40%; text-align: right; min-height: 24px; }}
            td:last-child {{ border-bottom: none; }}
            td::before {{ content: attr(data-label); position: absolute; left: 10px; width: 35%; text-align: left; font-weight: bold; color: #64748b; font-size: 13px; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>5月產銷表</h1>
        <input type="text" id="searchInput" class="search-box" placeholder="搜尋機台、批號或規格...">
    </div>
    
    <div class="info-bar">
        📄 檔案: {res['file_name']} ({res['file_time']})<br>
        🕒 網頁更新: {res['update_time']}
    </div>

    <div class="filter-group">
        <span>工程師:</span>
        <button class="btn btn-primary active" onclick="setEngineer('all', this)">全部</button>
        <button class="btn btn-default" onclick="setEngineer('wu', this)">吳</button>
        <button class="btn btn-default" onclick="setEngineer('lo', this)">羅</button>
    </div>

    <div class="card">
        <table id="dataTable">
            <thead>
                <tr>
                    <th>機台</th><th>批號</th><th>規格</th><th>排產日期</th><th>天數</th><th>T/D</th><th>排產總量 (KG)</th><th>已繳庫 (KG)</th><th>進度</th><th>需求量 (KG)</th><th>POY批號</th>
                </tr>
            </thead>
            <tbody id="tableBody"></tbody>
        </table>
    </div>

    <script>
        const rawData = {json_data};
        const tableBody = document.getElementById('tableBody');
        const searchInput = document.getElementById('searchInput');
        const engineerSpecs = {{
            wu: ["A10", "A11", "A12", "A16", "A17", "A18", "A22", "A23", "A24", "T01", "T02", "T03", "J08"],
            lo: ["A7", "A8", "A9", "A13", "A14", "A15", "A19", "A20", "A21", "T04", "T05", "T06", "J07"]
        }};
        let currentEngineer = 'all';

        function setEngineer(eng, btn) {{
            currentEngineer = eng;
            document.querySelectorAll('.btn').forEach(b => {{ b.classList.remove('btn-primary', 'active'); b.classList.add('btn-default'); }});
            btn.classList.remove('btn-default'); btn.classList.add('btn-primary', 'active');
            renderTable(searchInput.value);
        }}

        function renderTable(filter = '') {{
            tableBody.innerHTML = '';
            let filtered = rawData;
            if (currentEngineer !== 'all') {{
                const machines = engineerSpecs[currentEngineer];
                filtered = filtered.filter(item => {{
                    const m = String(item.machine).toUpperCase().trim();
                    return machines.some(target => {{
                        const targetUpper = target.toUpperCase();
                        if (m === targetUpper) return true;
                        const mNum = m.match(/[0-9]+/);
                        const tNum = targetUpper.match(/[0-9]+/);
                        if (mNum && tNum && m[0] === targetUpper[0]) {{ return parseInt(mNum[0]) === parseInt(tNum[0]); }}
                        return false;
                    }});
                }});
            }}
            filtered = filtered.filter(item => {{
                const searchStr = `${{item.machine}} ${{item.batch}} ${{item.spec}}`.toLowerCase();
                return searchStr.includes(filter.toLowerCase());
            }});
            filtered.forEach(item => {{
                const progress = item.pct;
                const demand = item.demand;
                const grades = item.grades;
                const statusClass = demand >= 0 ? 'status-positive' : 'status-negative';
                const displayDemand = demand >= 0 ? `+${{Math.round(demand).toLocaleString()}}` : Math.round(demand).toLocaleString();
                
                let gradeHtml = `<b>${{Math.round(item.stored).toLocaleString()}}</b><br>`;
                Object.keys(grades).forEach(g => {{ gradeHtml += `<span class="grade-tag grade-${{g}}">${{g}}:${{Math.round(grades[g]).toLocaleString()}}</span>`; }});
                
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td data-label="機台"><span class="badge machine-tag">${{item.machine}}</span></td>
                    <td data-label="批號"><b>${{item.batch}}</b></td>
                    <td data-label="規格"><small>${{item.spec}}</small></td>
                    <td data-label="排產日期"><small>${{item.date_range}}</small></td>
                    <td data-label="天數">${{item.days}}</td>
                    <td data-label="T/D">${{item.td}}</td>
                    <td data-label="排產總量">${{Math.round(item.target).toLocaleString()}}</td>
                    <td data-label="已繳庫">${{gradeHtml}}</td>
                    <td data-label="生產進度"><div class="progress-bar-container"><div class="progress-inner" style="width: ${{progress}}%"></div></div><span style="font-size: 11px">${{progress.toFixed(0)}}%</span></td>
                    <td data-label="需求量" class="${{statusClass}}">${{displayDemand}}</td>
                    <td data-label="POY批號"><small>${{item.poy}}</small></td>
                `;
                tableBody.appendChild(row);
            }});
        }}
        searchInput.addEventListener('input', (e) => renderTable(e.target.value));
        renderTable();
    </script>
</body>
</html>
"""
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    res = clean_data()
    generate_html(res)
    print("更新成功")
