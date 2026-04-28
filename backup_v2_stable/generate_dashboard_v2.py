import json
import os
import pandas as pd
from datetime import datetime
from core_processor import generate_dashboard_data, MasterTableEngine

def generate_html(data):
    # 獲取元數據
    update_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    json_data = json.dumps(data, ensure_ascii=False)
    
    # 1. 計算全月排產總目標 (過濾 NaN)
    total_target = sum(d['target'] for d in data if not pd.isna(d['target'])) / 1000 
    
    # 2. 核心修正：計算「本月自產繳庫量」 (總庫存 H 末尾 - 外購 H 末尾)
    self_produced_ton = 0
    try:
        lisa_path = os.path.join(os.getcwd(), "每日-最新庫存(DTY-LISA).xlsx")
        if os.path.exists(lisa_path):
            xl = pd.ExcelFile(lisa_path)
            df_total = xl.parse('總庫存', header=None)
            total_val = pd.to_numeric(df_total.iloc[:, 7], errors='coerce').dropna().iloc[-1]
            df_buy = xl.parse('外購', header=None)
            buy_val = pd.to_numeric(df_buy.iloc[:, 7], errors='coerce').dropna().iloc[-1]
            self_produced_ton = (total_val - buy_val) / 1000
        else:
            self_produced_ton = sum(d['stored'] for d in data if not pd.isna(d['stored'])) / 1000
    except Exception as e:
        print(f"計算自產繳庫量失敗: {e}")
        self_produced_ton = 0
    
    avg_progress = (self_produced_ton / total_target * 100) if total_target > 0 else 0
    
    abnormal_counts = {"改紡": 0, "了機": 0, "停機": 0, "待料": 0, "清車": 0, "檢修": 0}
    active_machines = 0
    processed_m = set()
    for d in data:
        m_base = d['machine'].split(' ')[0]
        if m_base not in processed_m:
            if d['actual_batch'] and d['actual_batch'] != '未開機':
                active_machines += 1
            processed_m.add(m_base)
        for cat in abnormal_counts:
            if cat in d['status_text']:
                abnormal_counts[cat] += 1
                break

    # 網頁模板
    html_template = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>撚二科產銷戰情室 2.0</title>
    <style>
        :root {
            --primary: #2563eb; --primary-dark: #1e40af;
            --success: #10b981; --warning: #f59e0b; --danger: #ef4444;
            --bg: #f8fafc; --card-bg: #ffffff; --text-main: #1e293b; --text-sub: #64748b;
            --border: #e2e8f0;
        }
        * { box-sizing: border-box; font-family: 'Segoe UI', 'Microsoft JhengHei', sans-serif; }
        body { background: var(--bg); color: var(--text-main); margin: 0; padding: 0; line-height: 1.5; }
        .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; flex-wrap: wrap; gap: 15px; }
        .header h1 { margin: 0; font-size: 24px; color: var(--primary-dark); display: flex; align-items: center; gap: 10px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 20px; margin-bottom: 25px; }
        .stat-card { background: var(--card-bg); padding: 20px; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid var(--border); display: flex; flex-direction: column; }
        .stat-label { font-size: 14px; color: var(--text-sub); font-weight: 600; margin-bottom: 8px; }
        .stat-value { font-size: 28px; font-weight: 800; color: var(--text-main); }
        .stat-unit { font-size: 14px; color: var(--text-sub); margin-left: 5px; }
        .stat-progress { height: 6px; background: #e2e8f0; border-radius: 3px; margin-top: 15px; overflow: hidden; }
        .stat-progress-bar { height: 100%; background: var(--primary); transition: width 0.5s ease; }
        .abnormal-panel { display: flex; gap: 12px; margin-bottom: 25px; flex-wrap: wrap; }
        .ab-tag { background: var(--card-bg); border: 1px solid var(--border); padding: 8px 16px; border-radius: 10px; font-size: 14px; font-weight: 700; display: flex; align-items: center; gap: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        .ab-count { color: var(--danger); font-size: 18px; }
        .controls { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 15px; }
        .search-group { display: flex; gap: 10px; flex-grow: 1; max-width: 600px; }
        .input-field { padding: 10px 15px; border: 1px solid var(--border); border-radius: 10px; font-size: 14px; outline: none; width: 100%; transition: all 0.2s; }
        .input-field:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
        .btn-group { display: flex; gap: 8px; }
        .btn { padding: 8px 16px; border: 1px solid var(--border); border-radius: 8px; background: #fff; cursor: pointer; font-size: 13px; font-weight: 600; color: var(--text-sub); transition: all 0.2s; }
        .btn.active { background: var(--primary); color: #fff; border-color: var(--primary); }
        .table-container { background: var(--card-bg); border-radius: 16px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); border: 1px solid var(--border); overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; min-width: 1000px; }
        th { background: #f1f5f9; padding: 15px; text-align: left; font-size: 12px; color: var(--text-sub); text-transform: uppercase; letter-spacing: 0.05em; }
        td { padding: 15px; border-bottom: 1px solid var(--border); font-size: 14px; vertical-align: top; }
        .machine-badge { background: #eff6ff; color: var(--primary-dark); padding: 4px 10px; border-radius: 6px; font-weight: 800; border: 1px solid #dbeafe; font-size: 13px; }
        .status-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 8px; }
        .dot-ok { background: var(--success); box-shadow: 0 0 8px var(--success); }
        .dot-diff { background: var(--warning); box-shadow: 0 0 8px var(--warning); }
        .dot-none { background: #cbd5e1; }
        .poy-tag { background: #f1f5f9; color: var(--text-sub); padding: 4px 8px; border-radius: 5px; font-size: 11px; margin-top: 5px; display: inline-block; font-weight: 600; }
        .poy-warning { background: #fff7ed; color: #9a3412; border: 1px solid #ffedd5; }
        .poy-danger { background: #fef2f2; color: #991b1b; border: 1px solid #fee2e2; animation: pulse 2s infinite; }
        .forecast-badge { background: #f0fdf4; color: #166534; padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: 800; border: 1px solid #dcfce7; margin-top: 5px; display: inline-block; }
        @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.6; } 100% { opacity: 1; } }

        /* RWD 修正核心：確保內容不被隱藏 */
        @media (max-width: 992px) {
            .table-container { border: none; background: transparent; box-shadow: none; overflow-x: visible; }
            table, thead, tbody, th, td, tr { display: block; width: 100%; min-width: unset; }
            thead { display: none; }
            tr { background: #fff; margin-bottom: 20px; border-radius: 16px; border: 1px solid var(--border); padding: 10px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
            td { border: none; padding: 10px 15px; display: flex !important; justify-content: space-between; align-items: flex-start; text-align: right; border-bottom: 1px solid #f8fafc; }
            td:last-child { border-bottom: none; }
            td::before { content: attr(data-label); font-weight: 800; color: var(--text-sub); text-align: left; font-size: 12px; min-width: 100px; padding-top: 2px; }
            .td-content { flex: 1; display: flex; flex-direction: column; align-items: flex-end; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 撚二科產銷戰情室 2.0 <small style="font-size: 12px; color: var(--text-sub); font-weight: normal;">Ver 2.0 (New Engine)</small></h1>
            <div style="text-align: right;">
                <div style="font-size: 12px; color: var(--text-sub);">數據同步：VAR_UPDATE_TIME</div>
                <div style="font-size: 11px; color: var(--success);">● 引擎連接正常 (LISA Core Active)</div>
            </div>
        </div>
        <div class="stats-grid">
            <div class="stat-card">
                <span class="stat-label">全月排產總目標</span>
                <span class="stat-value">VAR_TARGET<span class="stat-unit">T</span></span>
                <div class="stat-progress"><div class="stat-progress-bar" style="width: VAR_PROG_VAL%"></div></div>
            </div>
            <div class="stat-card">
                <span class="stat-label">本月自產繳庫量</span>
                <span class="stat-value">VAR_PRODUCED<span class="stat-unit">T</span></span>
                <span style="font-size: 12px; color: var(--success); font-weight: bold; margin-top: 5px;">達成率: VAR_PROG_TEXT%</span>
            </div>
            <div class="stat-card">
                <span class="stat-label">生產機台數</span>
                <span class="stat-value">VAR_ACTIVE_M<span class="stat-unit">台</span></span>
                <span style="font-size: 12px; color: var(--text-sub); margin-top: 5px;">總監控機台數: VAR_TOTAL_M</span>
            </div>
        </div>
        <div class="abnormal-panel">
            <div class="ab-tag">🔄 改紡 <span class="ab-count">VAR_AB_1</span></div>
            <div class="ab-tag">🏁 了機 <span class="ab-count">VAR_AB_2</span></div>
            <div class="ab-tag">🛑 停機 <span class="ab-count">VAR_AB_3</span></div>
            <div class="ab-tag">📦 待料 <span class="ab-count">VAR_AB_4</span></div>
            <div class="ab-tag">🧹 清車 <span class="ab-count">VAR_AB_5</span></div>
            <div class="ab-tag">🔍 檢修 <span class="ab-count">VAR_AB_6</span></div>
        </div>
        <div class="controls">
            <div class="search-group">
                <input type="text" id="mainSearch" class="input-field" placeholder="🔍 搜尋機台、批號、規格、POY...">
                <input type="text" id="poySearch" class="input-field" placeholder="🏭 POY 跨機台彙總查詢..." style="max-width: 200px;">
            </div>
            <div class="btn-group">
                <button class="btn active" onclick="filterEngineer('all', this)">全部</button>
                <button class="btn" onclick="filterEngineer('wu', this)">吳工</button>
                <button class="btn" onclick="filterEngineer('lo', this)">羅工</button>
            </div>
        </div>
        <div id="poySummary" style="display:none; margin-bottom: 20px; background: #eff6ff; padding: 15px; border-radius: 12px; border: 1px solid #bfdbfe;"></div>
        <div class="table-container">
            <table>
                <thead><tr><th>機台狀態</th><th>批號 (產銷/實際)</th><th>備註</th><th>日產(預計/實際平均)</th><th>規格需求</th><th>已繳庫量</th><th>預計完工</th><th>POY 庫存預警</th></tr></thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>
    <script>
        const rawData = VAR_JSON_DATA;
        const engineerSpecs = {wu:["A10","A11","A12","A16","A17","A18","A22","A23","A24","T01","T02","T03","J08"],lo:["A07","A08","A09","A13","A14","A15","A19","A20","A21","T04","T05","T06","J07"]};
        let currentEngineer = 'all';
        function renderTable() {
            const tbody = document.getElementById('tableBody');
            const search = document.getElementById('mainSearch').value.toLowerCase();
            const poySearch = document.getElementById('poySearch').value.toUpperCase();
            tbody.innerHTML = '';
            let filtered = rawData.filter(d => {
                const matchSearch = `${d.machine} ${d.batch} ${d.actual_batch} ${d.spec} ${d.note}`.toLowerCase().includes(search);
                const matchEng = currentEngineer === 'all' || engineerSpecs[currentEngineer].some(prefix => d.machine.startsWith(prefix));
                const matchPoy = !poySearch || d.poy.toUpperCase().includes(poySearch);
                return matchSearch && matchEng && matchPoy;
            });
            if (poySearch) updatePoySummary(poySearch, filtered);
            else document.getElementById('poySummary').style.display = 'none';
            filtered.forEach(d => {
                const tr = document.createElement('tr');
                const statusDot = d.actual_batch === '未開機' ? 'dot-none' : (d.actual_batch.includes(d.batch.replace('FD','').replace('LIKE','').trim()) ? 'dot-ok' : 'dot-diff');
                const poyWarnClass = d.poy_days < 3 ? 'poy-danger' : (d.poy_days < 5 ? 'poy-warning' : '');
                
                const demandLabel = d.demand >= 0 ? '盈餘' : '缺';
                const demandColor = d.demand >= 0 ? 'var(--success)' : 'var(--text-sub)';

                tr.innerHTML = `
                    <td data-label="機台狀態"><div class="td-content">
                        <span class="machine-badge">${d.machine}</span>
                        ${d.status_text ? `<div style="color:var(--danger); font-size:11px; font-weight:bold; margin-top:5px;">● ${d.status_text}</div>` : ''}
                    </div></td>
                    <td data-label="批號 (產銷/實際)"><div class="td-content">
                        <div><span class="status-dot ${statusDot}"></span><strong>${d.batch}</strong></div>
                        <div style="font-size:11px; color:var(--text-sub); margin-top:4px;">實際: ${d.actual_batch}</div>
                    </div></td>
                    <td data-label="備註"><div class="td-content">
                        <small style="color:var(--text-sub); font-size:12px;">${d.note || '---'}</small>
                    </div></td>
                    <td data-label="日產(預計/實際平均)"><div class="td-content">
                        <div style="font-size:12px;">預計: ${d.td_plan.toFixed(2)}</div>
                        <div style="font-size:13px; font-weight:bold; color:var(--primary);">平均: ${d.td_actual.toFixed(2)}</div>
                    </div></td>
                    <td data-label="規格需求"><div class="td-content">
                        <div style="font-size:12px; font-weight:600;">${d.spec}</div>
                        <div style="font-size:11px; color:var(--text-sub);">目標: ${Math.round(d.target).toLocaleString()} KG</div>
                    </div></td>
                    <td data-label="已繳庫量"><div class="td-content">
                        <div style="font-weight:bold;">${Math.round(d.stored).toLocaleString()} KG</div>
                        <div class="stat-progress" style="width:80px; height:4px; margin-top:5px;"><div class="stat-progress-bar" style="width:${d.pct}%; background:var(--success);"></div></div>
                        <div style="font-size:10px; color:var(--text-sub);">${d.pct.toFixed(0)}%</div>
                    </div></td>
                    <td data-label="預計完工"><div class="td-content">
                        <span class="forecast-badge">${d.finish_forecast || '---'}</span>
                        <div style="font-size:10px; color:${demandColor}; margin-top:4px;">${demandLabel}: ${Math.round(Math.abs(d.demand)).toLocaleString()}</div>
                    </div></td>
                    <td data-label="POY 庫存預警"><div class="td-content">
                        <div style="font-size:12px; font-weight:bold;">${d.poy}</div>
                        <span class="poy-tag ${poyWarnClass}">支撐: ${d.poy_days > 99 ? '---' : d.poy_days.toFixed(1) + '天'}</span>
                    </div></td>
                `;
                tbody.appendChild(tr);
            });
        }
        function updatePoySummary(poyId, activeRows) {
            const summaryDiv = document.getElementById('poySummary');
            const totalDailyUsage = activeRows.reduce((sum, r) => sum + (r.td_actual * 1000), 0);
            let totalStock = 0; const seenPoy = new Set();
            activeRows.forEach(r => { r.poy_details.forEach(p => { if (p.id.includes(poyId) && !seenPoy.has(p.id)) { totalStock += p.qty; seenPoy.add(p.id); } }); });
            const daysLeft = totalDailyUsage > 0 ? (totalStock / totalDailyUsage) : 999;
            summaryDiv.style.display = 'block';
            summaryDiv.innerHTML = `<div style="display:flex; justify-content:space-between; align-items:center;"><strong style="color:var(--primary-dark);">🏭 POY 彙總分析: ${poyId}</strong><button onclick="document.getElementById('poySearch').value=''; renderTable();" style="border:none; background:none; color:var(--primary); cursor:pointer;">[關閉]</button></div><div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap:15px; margin-top:10px;"><div><small>關聯機台</small><div style="font-weight:bold;">${activeRows.length} 台</div></div><div><small>總日消耗</small><div style="font-weight:bold; color:var(--danger);">${Math.round(totalDailyUsage).toLocaleString()} KG/日</div></div><div><small>剩餘總量</small><div style="font-weight:bold;">${Math.round(totalStock).toLocaleString()} KG</div></div><div><small>預計全場支撐</small><div style="font-size:18px; font-weight:800; color:var(--danger);">${daysLeft > 99 ? '---' : daysLeft.toFixed(1) + ' 天'}</div></div></div>`;
        }
        function filterEngineer(eng, btn) { currentEngineer = eng; document.querySelectorAll('.btn').forEach(b => b.classList.remove('active')); btn.classList.add('active'); renderTable(); }
        document.getElementById('mainSearch').addEventListener('input', renderTable);
        document.getElementById('poySearch').addEventListener('input', renderTable);
        renderTable();
    </script>
</body>
</html>
"""

    final_content = (html_template
        .replace("VAR_UPDATE_TIME", update_time)
        .replace("VAR_TARGET", f"{total_target:,.1f}")
        .replace("VAR_PRODUCED", f"{self_produced_ton:,.2f}")
        .replace("VAR_PROG_VAL", f"{avg_progress}")
        .replace("VAR_PROG_TEXT", f"{avg_progress:.1f}")
        .replace("VAR_ACTIVE_M", str(active_machines))
        .replace("VAR_TOTAL_M", str(len(processed_m)))
        .replace("VAR_AB_1", str(abnormal_counts['改紡']))
        .replace("VAR_AB_2", str(abnormal_counts['了機']))
        .replace("VAR_AB_3", str(abnormal_counts['停機']))
        .replace("VAR_AB_4", str(abnormal_counts['待料']))
        .replace("VAR_AB_5", str(abnormal_counts['清車']))
        .replace("VAR_AB_6", str(abnormal_counts['檢修']))
        .replace("VAR_JSON_DATA", json_data))

    with open(os.path.join(os.path.dirname(__file__), 'production_dashboard.html'), 'w', encoding='utf-8') as f:
        f.write(final_content)

if __name__ == "__main__":
    print("🚀 啟動數據引擎...")
    data = generate_dashboard_data()
    print("🎨 正在產生視覺化看板...")
    generate_html(data)
    print("✨ 更新成功！請打開 production_dashboard.html 查看。")
