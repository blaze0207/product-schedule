import json
import os
import pandas as pd
import re
from datetime import datetime
from reality_analyzer import RealityLogAnalyzer

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def generate_v3_html(result):
    data = result['tasks']
    summary_weight = result['self_produced_weight']
    stock_date = result.get('stock_update_date', 'Unknown')
    production_date = result.get('production_date', 'Unknown')
    daily_sum = result.get('latest_daily_sum', 0)
    summary_prod_date = result.get('summary_prod_date', 'Unknown')
    
    update_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    data.sort(key=lambda x: (natural_sort_key(x['machine']), not x['is_active']))
    
    filtered_data = []
    for d in data:
        if d.get('dty_std') == '__STATUS__' and not d.get('is_active'):
            continue
        if 'last_date' in d and not isinstance(d['last_date'], str):
            d['last_date'] = d['last_date'].strftime('%m/%d')
        filtered_data.append(d)
            
    json_data = json.dumps(filtered_data, ensure_ascii=False)
    poy_json = json.dumps(result.get('poy_analysis', {}), ensure_ascii=False)
    
    html_content = r"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>撚二科生產戰情室</title>
    <style>
        :root { 
            --primary: #0f172a; 
            --secondary: #1e293b;
            --accent: #2563eb; 
            --success: #059669; 
            --warning: #d97706;
            --danger: #dc2626;
            --bg: #f1f5f9; 
            --card: #ffffff; 
            --text-main: #1e293b;
            --text-muted: #64748b;
        }
        
        * { box-sizing: border-box; font-family: 'Inter', 'Segoe UI', "Microsoft JhengHei", sans-serif; margin: 0; padding: 0; }
        body { background: var(--bg); padding: 0; color: var(--text-main); line-height: 1.5; }
        
        .top-header { 
            background: var(--primary); 
            color: white; 
            padding: 15px 30px; 
            display: flex; 
            justify-content: space-between; 
            align-items: center;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            position: sticky; top: 0; z-index: 100;
        }
        .top-header h1 { font-size: 24px; font-weight: 800; letter-spacing: 2px; display: flex; align-items: center; gap: 10px; }
        .top-header h1::before { content: ''; display: inline-block; width: 4px; height: 24px; background: var(--accent); border-radius: 2px; }
        
        .header-info { text-align: right; font-family: 'Consolas', monospace; }
        .header-info div { font-size: 12px; opacity: 0.8; margin-bottom: 2px; }
        .date-highlight { color: #60a5fa; font-weight: bold; }

        .main-container { padding: 25px; max-width: 1600px; margin: 0 auto; }

        .summary-bar { display: grid; grid-template-columns: repeat(2, minmax(300px, 1fr)); gap: 20px; margin-bottom: 25px; }
        .summary-card { 
            background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%); 
            color: white; 
            padding: 20px; 
            border-radius: 12px; 
            box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1);
            position: relative; overflow: hidden;
        }
        .summary-card::after { content: ''; position: absolute; right: -20px; bottom: -20px; width: 100px; height: 100px; background: rgba(255,255,255,0.05); border-radius: 50%; }
        .summary-label { font-size: 13px; text-transform: uppercase; letter-spacing: 1px; opacity: 0.7; margin-bottom: 8px; }
        .summary-value { font-size: 36px; font-weight: 900; font-family: 'Bahnschrift', sans-serif; }
        .summary-unit { font-size: 16px; margin-left: 8px; font-weight: normal; opacity: 0.6; }
        .card-accent { background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%); }

        .interactive-section { background: white; padding: 20px; border-radius: 12px; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-top: 4px solid var(--accent); }
        .filter-group { display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap; }
        .filter-btn { 
            padding: 8px 20px; border-radius: 8px; border: 1px solid #e2e8f0; 
            background: white; cursor: pointer; font-weight: bold; color: var(--text-main);
            transition: all 0.2s; font-size: 14px; display: flex; align-items: center; gap: 6px;
        }
        .filter-btn:hover { background: #f8fafc; border-color: var(--accent); color: var(--accent); }
        .filter-btn.active { background: var(--accent); border-color: var(--accent); color: white; box-shadow: 0 4px 6px rgba(37,99,235,0.2); }
        
        .interactive-bar { display: flex; gap: 20px; align-items: flex-start; }
        .search-box { flex: 0 0 350px; position: relative; display: flex; align-items: center; }
        .search-input { 
            width: 100%; padding: 14px 45px 14px 15px; border-radius: 8px; 
            border: 2px solid #e2e8f0; font-size: 16px; outline: none; transition: all 0.3s; 
            background: #f8fafc;
        }
        .search-input:focus { border-color: var(--accent); background: white; box-shadow: 0 0 0 4px rgba(37,99,235,0.1); }
        .clear-btn { position: absolute; right: 15px; cursor: pointer; color: var(--danger); font-size: 22px; font-weight: 900; display: none; transition: transform 0.1s; }
        .clear-btn:hover { transform: scale(1.2); color: #be123c; }
        
        .analysis-panel { 
            flex: 1; border-radius: 10px; padding: 15px 20px; display: none; 
            background: #f0f7ff; border: 1px solid #bae6fd; animation: slideIn 0.4s ease-out;
        }
        @keyframes slideIn { from { opacity: 0; transform: translateX(20px); } to { opacity: 1; transform: translateX(0); } }

        .table-wrapper { background: var(--card); border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.05); overflow: hidden; }
        table { width: 100%; border-collapse: collapse; table-layout: fixed; }
        th { background: #f8fafc; padding: 15px 12px; text-align: left; font-size: 12px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.1em; border-bottom: 2px solid #e2e8f0; }
        td { padding: 16px 12px; border-bottom: 1px solid #f1f5f9; font-size: 15px; vertical-align: middle; word-wrap: break-word; }
        
        tr.active { background: white; }
        tr.active:hover { background: #f8fafc; }
        tr.history { background: #fbfcfd; opacity: 0.85; }
        tr.history td { color: var(--text-muted); }

        .machine-badge { 
            background: var(--primary); color: white; padding: 6px 10px; border-radius: 6px; 
            font-weight: 800; font-size: 14px; display: inline-flex; align-items: center; gap: 5px;
        }
        
        .status-indicator { display: flex; align-items: center; gap: 8px; font-weight: bold; }
        .dot { width: 8px; height: 8px; border-radius: 50%; }
        .dot-running { background: var(--success); box-shadow: 0 0 8px var(--success); animation: pulse 2s infinite; }
        @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.4; } 100% { opacity: 1; } }
        .dot-stopped { background: var(--text-muted); }

        .batch-text { font-weight: 700; color: var(--primary); font-size: 16px; display: block; margin-bottom: 4px; }
        .spec-text { font-size: 13px; color: var(--text-muted); }

        .poy-info-box { 
            margin-bottom: 12px; padding: 10px; border-radius: 8px; 
            background: #f8fafc; border-left: 4px solid #cbd5e1; transition: all 0.2s;
        }
        .poy-info-box:hover { border-left-color: var(--accent); background: #f1f5f9; }
        .clickable-poy { cursor: pointer; color: var(--accent); font-weight: 800; font-family: 'Consolas', monospace; font-size: 15px; }
        .stock-label { font-size: 13px; font-weight: 900; margin-top: 5px; }
        .history-tag { font-size: 11px; color: var(--text-muted); margin-top: 3px; display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }

        .stock-tags { display: flex; flex-wrap: wrap; gap: 4px; }
        .tag { font-size: 11px; padding: 2px 6px; border-radius: 4px; font-weight: 700; }
        .tag-a { background: #d1fae5; color: #065f46; }
        .tag-ax { background: #fef3c7; color: #92400e; }
        .tag-b { background: #ffedd5; color: #9a3412; }
        .tag-c { background: #fee2e2; color: #991b1b; }
        .total-stock { font-size: 18px; font-weight: 900; color: var(--primary); margin-top: 5px; display: block; }
        
        @media (max-width: 768px) {
            .top-header { flex-direction: column; align-items: flex-start; gap: 10px; padding: 15px; }
            .header-info { text-align: left; }
            .main-container { padding: 15px; }
            .summary-bar { grid-template-columns: 1fr; }
            .interactive-bar { flex-direction: column; }
            .search-box { flex: 1 0 auto; width: 100%; }
            table, thead, tbody, th, td, tr { display: block; }
            thead { display: none; }
            tr { margin-bottom: 20px; border: 1px solid #e2e8f0; border-radius: 12px; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.05); overflow: hidden; }
            td { border: none; position: relative; padding-left: 40%; text-align: left; border-bottom: 1px solid #f1f5f9; min-height: 45px; }
            td::before { content: attr(data-label); position: absolute; left: 15px; width: 35%; font-size: 12px; font-weight: bold; color: var(--text-muted); text-transform: uppercase; }
            .poy-info-box { border-left-width: 2px; padding: 5px 0 5px 10px; margin-top: 5px; }
            .machine-badge { font-size: 16px; }
        }
    </style>
</head>
<body>
    <header class="top-header">
        <h1>撚二科生產戰情室</h1>
        <div class="header-info">
            <div>SYSTEM ENGINE: <span class="date-highlight">v3.2 PROFESSIONAL</span></div>
            <div>數據同步：VAR_TIME</div>
            <div>生產資訊日期：<span class="date-highlight">VAR_PROD_DATE</span></div>
            <div>庫存更新日期：<span class="date-highlight">VAR_STOCK_DATE</span></div>
        </div>
    </header>

    <div class="main-container">
        <div class="summary-bar">
            <div class="summary-card">
                <div class="summary-label">本月累計繳庫 (自產)</div>
                <div class="summary-value">VAR_SUMMARY<span class="summary-unit">KG</span></div>
            </div>
            <div class="summary-card card-accent">
                <div class="summary-label">每日產量 (VAR_SETTLE_DATE)</div>
                <div class="summary-value">VAR_DAILY_SUM<span class="summary-unit">T</span></div>
            </div>
        </div>

        <section class="interactive-section">
            <div class="filter-group">
                <button class="filter-btn active" id="btn-all" onclick="filterByEngineer('all')">🌐 全部</button>
                <button class="filter-btn" id="btn-luo" onclick="filterByEngineer('luo')">👤 羅工</button>
                <button class="filter-btn" id="btn-wu" onclick="filterByEngineer('wu')">👤 吳工</button>
            </div>
            <div class="interactive-bar">
                <div class="search-box">
                    <input type="text" id="poySearch" class="search-input" placeholder="輸入或點擊 POY 批號追蹤分析..." oninput="handleSearch(this.value)">
                    <span id="clearSearch" class="clear-btn" onclick="clearSearch()">&times;</span>
                </div>
                <div id="analysisPanel" class="analysis-panel"></div>
            </div>
        </section>

        <div class="table-wrapper">
            <table>
                <thead>
                    <tr>
                        <th style="width: 8%">機台</th>
                        <th style="width: 12%">狀態</th>
                        <th style="width: 14%">批號 / 規格</th>
                        <th style="width: 20%">POY庫存(A+A2)</th>
                        <th style="width: 10%; text-align: center;">預估支撐</th>
                        <th style="width: 8%; text-align: center;">平均日產</th>
                        <th style="width: 7%; text-align: center;">生產天數</th>
                        <th style="width: 21%">繳庫量</th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
    </div>

    <script>
        const rawData = VAR_JSON;
        const poyAnalysis = VAR_POY_JSON;
        const tbody = document.getElementById('tableBody');

        const engineerMaps = {
            'luo': ['A07','A08','A09','A13','A14','A15','A19','A20','A21','J07','T04','T05','T06'],
            'wu': ['A10','A11','A12','A16','A17','A18','A22','A23','A24','J08','T01','T02','T03']
        };

        function filterByEngineer(type) {
            document.querySelectorAll('.filter-btn').forEach(btn => btn.classList.remove('active'));
            if (type === 'all') document.getElementById('btn-all').classList.add('active');
            else if (type === 'luo') document.getElementById('btn-luo').classList.add('active');
            else if (type === 'wu') document.getElementById('btn-wu').classList.add('active');

            document.getElementById('poySearch').value = '';
            document.getElementById('analysisPanel').style.display = 'none';
            document.getElementById('clearSearch').style.display = 'none';

            const rows = tbody.getElementsByTagName('tr');
            Array.from(rows).forEach(row => {
                const badge = row.querySelector('.machine-badge');
                if(!badge) return;
                const mId = badge.innerText.split(' ')[0];
                row.style.display = (type === 'all' || engineerMaps[type].includes(mId)) ? '' : 'none';
            });
        }

        function cleanPoyId(val) {
            if (!val) return "";
            return val.trim().toUpperCase().replace(/[×*X]\d+$/, '').trim();
        }
        
        function handleSearch(val) {
            const queryRaw = val.trim().toUpperCase();
            const queryClean = cleanPoyId(queryRaw);
            const panel = document.getElementById('analysisPanel');
            const clearBtn = document.getElementById('clearSearch');
            const rows = tbody.getElementsByTagName('tr');
            
            clearBtn.style.display = queryRaw ? 'block' : 'none';

            if (!queryRaw) {
                panel.style.display = 'none';
                Array.from(rows).forEach(r => r.style.display = '');
                return;
            }

            Array.from(rows).forEach(row => {
                row.style.display = row.innerText.toUpperCase().includes(queryRaw) ? '' : 'none';
            });

            let matchedKey = Object.keys(poyAnalysis).find(k => k === queryClean);
            if (!matchedKey && queryClean.length >= 3) {
                matchedKey = Object.keys(poyAnalysis).find(k => k.includes(queryClean));
            }

            if (matchedKey) {
                const info = poyAnalysis[matchedKey];
                panel.style.display = 'block';
                const color = info.support_days < 3 ? 'var(--danger)' : 'var(--accent)';
                panel.innerHTML = `
                    <div style="display:flex; justify-content:space-between; align-items:center; height:100%;">
                        <div>
                            <div style="font-size:12px; color:var(--text-muted); font-weight:bold; margin-bottom:5px;">POY 深度分析系統:</div>
                            <div style="font-weight:900; font-size:20px; color:var(--primary);">${matchedKey} <span style="font-size:14px; font-weight:normal; color:var(--text-muted); margin-left:10px;">${info.spec}</span></div>
                            <div style="font-size:13px; margin-top:5px; color:var(--text-muted);">關聯機台: <span style="color:var(--accent); font-weight:bold;">${info.machines_text}</span></div>
                        </div>
                        <div style="text-align:right; border-left:2px solid #e2e8f0; padding-left:25px;">
                            <div style="font-size:12px; color:var(--text-muted); font-weight:bold;">全場預估支撐天數</div>
                            <div style="font-size:28px; font-weight:900; color:${color};">${info.support_days.toFixed(2)} <span style="font-size:14px;">DAYS</span></div>
                            <div style="font-size:11px; color:var(--text-muted);">庫存 ${info.stock_a_a2.toLocaleString()} KG / 日耗 ${(info.total_avg_td*1000).toLocaleString()} KG</div>
                        </div>
                    </div>
                `;
            } else {
                panel.style.display = 'none';
            }
        }

        function clearSearch() {
            document.getElementById('poySearch').value = '';
            handleSearch('');
        }

        function clickPoy(batch) {
            document.getElementById('poySearch').value = batch;
            handleSearch(batch);
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        rawData.forEach(d => {
            const tr = document.createElement('tr');
            const isActive = d.is_active;
            tr.className = isActive ? 'active' : 'history';
            
            const s = d.stock || {A:0, AX:0, B:0, C:0};
            let stockTags = "";
            if (s.A > 0) stockTags += `<span class="tag tag-a">A:${s.A.toFixed(0)}</span>`;
            if (s.AX > 0) stockTags += `<span class="tag tag-ax">AX:${s.AX.toFixed(0)}</span>`;
            if (s.B > 0) stockTags += `<span class="tag tag-b">B:${s.B.toFixed(0)}</span>`;
            if (s.C > 0) stockTags += `<span class="tag tag-c">C:${s.C.toFixed(0)}</span>`;
            
            const totalStock = d.stock_summary?.total_deposit > 0 
                ? `<span class="total-stock">總:${d.stock_summary.total_deposit.toFixed(0)}</span>` 
                : '<span style="color:#cbd5e1">---</span>';

            let statusHtml = "";
            if (isActive) {
                if (d.dty_batch === '---') {
                    const displayStatus = d.last_status || '停機';
                    const isStop = displayStatus.includes('停機');
                    statusHtml = `<div class="status-indicator"><span class="dot" style="background:${isStop ? 'var(--danger)' : '#94a3b8'}"></span><span style="color:${isStop ? 'var(--danger)' : 'var(--text-muted)'}; font-weight:bold;">${displayStatus}</span></div>`;
                } else {
                    statusHtml = `<div class="status-indicator"><span class="dot dot-running"></span><span style="color:var(--success)">運行中</span></div>`;
                    if (d.last_status) {
                        const isStop = d.last_status.includes('停機');
                        statusHtml += `<div style="font-size:12px; color:${isStop ? 'var(--danger)' : 'var(--warning)'}; margin-top:4px; font-weight:bold;">[${d.last_status}]</div>`;
                    }
                }
            } else {
                statusHtml = `<div style="font-size:13px; color:var(--text-muted)">已完工 (${d.last_date})</div>`;
            }

            let poyHtml = '<span style="color:#cbd5e1">-</span>';
            let supportHtml = '<span style="color:#cbd5e1">-</span>';
            if (d.poy_list && d.poy_list.length > 0) {
                poyHtml = d.poy_list.map(p => {
                    const stockVal = parseFloat(p.stock_a_a2).toLocaleString(undefined, {minimumFractionDigits:1});
                    const color = p.stock_a_a2 < 500 ? 'var(--danger)' : 'var(--accent)';
                    return `
                        <div class="poy-info-box">
                            <div class="clickable-poy" onclick="clickPoy('${p.batch}')">${p.batch}</div>
                            <div class="spec-text">${p.spec}</div>
                            <div class="stock-label" style="color:${color}">庫存(A+A2): ${stockVal} KG</div>
                            <div class="history-tag" title="${p.history}">履歷: ${p.history}</div>
                        </div>
                    `;
                }).join('');

                supportHtml = d.poy_list.map(p => {
                    const days = p.support_days;
                    const color = days < 3 ? 'var(--danger)' : 'var(--accent)';
                    const weight = days < 3 ? '900' : 'bold';
                    return `<div style="margin-bottom:12px; padding:10px; height:75px; display:flex; align-items:center; justify-content:center; font-size:18px; color:${color}; font-weight:${weight}; font-family:'Consolas';">
                        ${days >= 999 ? '-' : days.toFixed(2) + ' 天'}
                    </div>`;
                }).join('');
            }

            tr.innerHTML = `
                <td data-label="機台"><span class="machine-badge">${d.machine} <span style="opacity:0.6; font-size:11px;">(${isActive ? d.current_sides.join('/') : d.produced_sides.join('/')})</span></span></td>
                <td data-label="狀態">${statusHtml}</td>
                <td data-label="批號 / 規格">
                    <span class="batch-text">${d.dty_batch}</span>
                    <span class="spec-text">${d.spec}</span>
                </td>
                <td data-label="POY庫存(A+A2)">${poyHtml}</td>
                <td data-label="預估支撐" style="text-align: center;">${supportHtml}</td>
                <td data-label="平均日產" style="text-align: center; font-weight:900; font-family:'Consolas'; color:${isActive && d.dty_batch !== '---' ? 'var(--accent)' : 'var(--text-muted)'}; font-size:18px;">
                    ${d.avg_td > 0 ? d.avg_td.toFixed(2) : '---'}
                </td>
                <td data-label="生產天數" style="text-align: center; font-weight:bold;">${d.days}</td>
                <td data-label="繳庫量">
                    <div class="stock-tags">${stockTags}</div>
                    ${totalStock}
                </td>
            `;
            tbody.appendChild(tr);
        });
    </script>
</body>
</html>
"""
    final_html = html_content.replace("VAR_TIME", update_time) \
                              .replace("VAR_JSON", json_data) \
                              .replace("VAR_POY_JSON", poy_json) \
                              .replace("VAR_SUMMARY", f"{summary_weight:,.1f}") \
                              .replace("VAR_STOCK_DATE", stock_date) \
                              .replace("VAR_PROD_DATE", production_date) \
                              .replace("VAR_DAILY_SUM", f"{daily_sum:,.2f}") \
                              .replace("VAR_SETTLE_DATE", summary_prod_date)

    with open('production_dashboard.html', 'w', encoding='utf-8') as f:
        f.write(final_html)

if __name__ == "__main__":
    analyzer = RealityLogAnalyzer(os.getcwd())
    result = analyzer.get_reality_tasks()
    generate_v3_html(result)
    print("✨ 包含自產繳庫量彙總 3.0 更新成功！")
