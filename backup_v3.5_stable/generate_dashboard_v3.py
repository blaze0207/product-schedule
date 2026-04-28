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
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap');
        
        :root { 
            --primary: #0f172a; 
            --secondary: #334155;
            --accent: #3b82f6; 
            --accent-soft: #eff6ff;
            --success: #10b981; 
            --warning: #f59e0b;
            --danger: #ef4444;
            --bg: #f8fafc; 
            --card: #ffffff; 
            --border: #e2e8f0;
            --text-main: #1e293b;
            --text-muted: #64748b;
        }
        
        * { box-sizing: border-box; font-family: 'Inter', 'Segoe UI', "Microsoft JhengHei", sans-serif; margin: 0; padding: 0; }
        body { background: var(--bg); color: var(--text-main); line-height: 1.6; }
        
        /* 頂部標題列 - Premium Dark Style */
        .top-header { 
            background: var(--primary); 
            color: white; 
            padding: 24px 40px; 
            display: flex; 
            justify-content: space-between; 
            align-items: center;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            position: sticky; top: 0; z-index: 100;
        }
        .top-header h1 { font-size: 32px; font-weight: 900; letter-spacing: 2px; display: flex; align-items: center; gap: 15px; }
        .top-header h1::before { content: ''; display: inline-block; width: 8px; height: 35px; background: var(--accent); border-radius: 4px; }
        
        .header-info { text-align: right; font-family: 'Consolas', monospace; color: #cbd5e1; }
        .header-info div { font-size: 14px; margin-bottom: 4px; }
        .date-highlight { color: #60a5fa; font-weight: 800; }

        .main-container { padding: 40px; max-width: 1800px; margin: 0 auto; }

        /* KPI 數據概覽 */
        .summary-bar { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin-bottom: 30px; }
        .summary-card { 
            background: linear-gradient(145deg, #1e293b 0%, #0f172a 100%); 
            color: white; padding: 20px 30px; border-radius: 16px; 
            box-shadow: 0 10px 20px rgba(0,0,0,0.1);
            position: relative; overflow: hidden; border: 1px solid rgba(255,255,255,0.1);
        }
        .summary-label { font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: #94a3b8; margin-bottom: 8px; }
        .summary-value { font-size: 38px; font-weight: 900; font-family: 'Inter', sans-serif; letter-spacing: -0.5px; line-height: 1.2; }
        .summary-unit { font-size: 18px; margin-left: 8px; opacity: 0.5; font-weight: 400; }
        .card-accent { border-top: 5px solid var(--accent); }

        /* 互動控制區塊 */
        .interactive-section { background: white; padding: 30px; border-radius: 16px; margin-bottom: 40px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); border: 1px solid var(--border); }
        .filter-group { display: flex; gap: 15px; margin-bottom: 25px; }
        .filter-btn { 
            padding: 12px 32px; border-radius: 12px; border: 2px solid var(--border); 
            background: white; cursor: pointer; font-weight: 800; color: var(--text-main);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); font-size: 18px;
        }
        .filter-btn:hover { border-color: var(--accent); background: var(--accent-soft); color: var(--accent); transform: translateY(-2px); }
        .filter-btn.active { background: var(--accent); border-color: var(--accent); color: white; box-shadow: 0 8px 15px rgba(59,130,246,0.3); }
        
        .interactive-bar { display: flex; gap: 30px; align-items: stretch; }
        .search-box { position: relative; flex: 0 0 450px; }
        .search-input { 
            width: 100%; padding: 20px 60px 20px 25px; border-radius: 15px; 
            border: 2px solid var(--border); font-size: 20px; outline: none; transition: all 0.3s; background: #fcfdfe;
        }
        .search-input:focus { border-color: var(--accent); background: white; box-shadow: 0 0 0 6px rgba(59,130,246,0.1); }
        .clear-btn { position: absolute; right: 25px; top: 20px; cursor: pointer; color: var(--danger); font-size: 28px; font-weight: 900; display: none; }
        
        .analysis-panel { 
            flex: 1; border-radius: 15px; padding: 25px 30px; display: none; 
            background: linear-gradient(to right, #f0f9ff, #e0f2fe); border: 1px solid #bae6fd; animation: fadeIn 0.4s ease;
        }
        @keyframes fadeIn { from { opacity: 0; scale: 0.98; } to { opacity: 1; scale: 1; } }

        /* 表格排版旗艦版 */
        .table-wrapper { background: white; border-radius: 20px; box-shadow: 0 10px 40px rgba(0,0,0,0.04); overflow: hidden; border: 1px solid var(--border); }
        table { width: 100%; border-collapse: collapse; table-layout: fixed; }
        /* 提升表頭字體大小至 16px */
        th { 
            background: #f8fafc; padding: 25px 15px; text-align: left; 
            font-size: 16px; font-weight: 900; color: var(--primary); 
            text-transform: uppercase; letter-spacing: 0.1em; border-bottom: 4px solid var(--border); 
        }
        td { padding: 22px 12px; border-bottom: 1px solid #f1f5f9; font-size: 17px; vertical-align: middle; }

        tr.active { background: white; transition: background 0.2s; }
        tr.active:hover { background: #fcfdfe; }
        tr.history { opacity: 0.7; background: #fbfcfd; }

        .machine-badge { 
            background: var(--primary); color: white; padding: 10px 14px; border-radius: 8px; 
            font-weight: 900; font-size: 18px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);
            white-space: nowrap; display: inline-flex; align-items: center; gap: 8px;
        }

        
        .batch-text { font-weight: 900; color: var(--primary); font-size: 21px; display: block; margin-bottom: 8px; }
        .spec-text { font-size: 15px; color: var(--text-muted); font-weight: 500; }

        /* POY 互動膠囊 */
        .poy-info-box { 
            margin-bottom: 18px; padding: 15px; border-radius: 12px; 
            background: #f8fafc; border-left: 6px solid #cbd5e1; 
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); cursor: pointer;
        }
        .poy-info-box:hover { border-left-color: var(--accent); background: var(--accent-soft); transform: translateX(5px); box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
        .clickable-poy { color: var(--accent); font-weight: 900; font-family: 'Consolas', monospace; font-size: 18px; text-decoration: none; }
        .stock-label { font-size: 16px; font-weight: 900; margin-top: 8px; }
        .history-tag { font-size: 13px; color: var(--text-muted); margin-top: 6px; display: block; border-top: 1px solid rgba(0,0,0,0.05); padding-top: 5px; }

        .tag { font-size: 13px; padding: 4px 10px; border-radius: 6px; font-weight: 800; display: inline-block; }
        .tag-a { background: #dcfce7; color: #065f46; }
        .tag-ax { background: #fef3c7; color: #92400e; }
        .tag-b { background: #ffedd5; color: #9a3412; }
        .tag-c { background: #fee2e2; color: #991b1b; }
        .total-stock { font-size: 22px; font-weight: 900; color: var(--primary); margin-top: 10px; display: block; }

        /* 進度條樣式 */
        .progress-container {
            width: 100%; height: 8px; background: #e2e8f0; border-radius: 10px;
            margin-top: 12px; overflow: hidden; position: relative;
        }
        .progress-bar {
            height: 100%; border-radius: 10px;
            transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .progress-info {
            display: flex; justify-content: space-between; align-items: center;
            margin-top: 6px; font-size: 12px; font-weight: 800; color: var(--text-muted);
        }

        /* 指示燈 */
        .status-indicator { display: flex; align-items: center; gap: 10px; }
        .dot { width: 12px; height: 12px; border-radius: 50%; }
        .dot-running { background: var(--success); box-shadow: 0 0 12px var(--success); animation: pulse 2s infinite; }
        @keyframes pulse { 0%, 100% { opacity: 1; transform: scale(1); } 50% { opacity: 0.5; transform: scale(0.9); } }

        /* RWD 佈局 2.0 - 使用 Grid 徹底修復 */
        @media (max-width: 1200px) {
            .top-header { padding: 20px; }
            .main-container { padding: 20px; }
            .summary-bar { grid-template-columns: 1fr; }
            .search-box { flex: 1; }
            
            table, thead, tbody, th, td, tr { display: block; }
            thead { display: none; }
            tr { 
                margin-bottom: 40px; border: 3px solid var(--border); border-radius: 20px; 
                padding: 15px; position: relative; background: white;
            }
            td { 
                display: grid; grid-template-columns: 140px 1fr; gap: 15px; 
                padding: 15px 10px; border: none; border-bottom: 1px solid #f1f5f9;
                text-align: left !important; min-height: 60px; align-items: center;
            }
            td::before { 
                content: attr(data-label); font-size: 14px; font-weight: 900; 
                color: var(--text-muted); text-transform: uppercase;
            }
            td[style*="text-align: center"] { justify-content: start !important; }
            .poy-info-box { margin-bottom: 10px; }
            .batch-text { font-size: 18px; }
        }
    </style>
</head>
<body>
    <header class="top-header">
        <h1>撚二科生產戰情室</h1>
        <div class="header-info">
            <div>ENGINE: <span class="date-highlight">v3.5 PREMIUM</span></div>
            <div>同步時間: VAR_TIME</div>
            <div>生產資訊日: <span class="date-highlight">VAR_PROD_DATE</span></div>
            <div>DTY/POY庫存: <span class="date-highlight">VAR_STOCK_DATE</span></div>
        </div>
    </header>

    <div class="main-container">
        <div class="summary-bar">
            <div class="summary-card">
                <div class="summary-label">本月累計自產繳庫</div>
                <div class="summary-value">VAR_SUMMARY<span class="summary-unit">KG</span></div>
            </div>
            <div class="summary-card card-accent">
                <div class="summary-label">每日產量結算 (VAR_SETTLE_DATE)</div>
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
                        <th style="width: 16%">批號 / 規格</th>
                        <th style="width: 20%">POY庫存(A+A2)</th>
                        <th style="width: 10%; text-align: center;">預估支撐</th>
                        <th style="width: 8%; text-align: center;">平均日產</th>
                        <th style="width: 7%; text-align: center;">生產天數</th>
                        <th style="width: 19%">繳庫量</th>
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
            document.getElementById('btn-'+type)?.classList.add('active');
            
            clearSearch();

            Array.from(tbody.rows).forEach(row => {
                const badge = row.querySelector('.machine-badge');
                if(!badge) return;
                // 使用正則抓取字首的機台號 (如 A07)
                const mMatch = badge.innerText.match(/^[A-Z]\d+/);
                const mId = mMatch ? mMatch[0] : "";
                
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
            
            clearBtn.style.display = queryRaw ? 'block' : 'none';

            if (!queryRaw) {
                panel.style.display = 'none';
                Array.from(tbody.rows).forEach(r => r.style.display = '');
                return;
            }

            Array.from(tbody.rows).forEach(row => {
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
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <div>
                            <div style="font-size:14px; color:var(--text-muted); font-weight:800; margin-bottom:10px;">POY ANALYSIS SYSTEM:</div>
                            <div style="font-weight:900; font-size:26px; color:var(--primary);">${matchedKey} <span style="font-size:16px; font-weight:500; color:var(--text-muted); margin-left:15px;">${info.spec}</span></div>
                            <div style="font-size:16px; margin-top:10px;">共用機台: <span style="color:var(--accent); font-weight:900;">${info.machines_text}</span></div>
                        </div>
                        <div style="text-align:right; border-left:3px solid var(--border); padding-left:40px;">
                            <div style="font-size:15px; color:var(--text-muted); font-weight:800;">預估全場支撐</div>
                            <div style="font-size:42px; font-weight:900; color:${color}; line-height:1;">${info.support_days.toFixed(2)} <span style="font-size:18px;">DAYS</span></div>
                            <div style="font-size:14px; color:var(--text-muted); margin-top:10px;">庫存 ${info.stock_a_a2.toLocaleString()} KG / 日耗 ${(info.total_avg_td*1000).toLocaleString()} KG</div>
                        </div>
                    </div>
                `;
            } else { panel.style.display = 'none'; }
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
            if (s.A > 0) stockTags += `<span class="tag tag-a">A:${s.A.toFixed(0)}</span> `;
            if (s.AX > 0) stockTags += `<span class="tag tag-ax">AX:${s.AX.toFixed(0)}</span> `;
            if (s.B > 0) stockTags += `<span class="tag tag-b">B:${s.B.toFixed(0)}</span> `;
            if (s.C > 0) stockTags += `<span class="tag tag-c">C:${s.C.toFixed(0)}</span>`;
            const totalStock = d.stock_summary?.total_deposit > 0 ? `<span class="total-stock">總:${d.stock_summary.total_deposit.toLocaleString()}</span>` : '-';

            let statusHtml = "";
            if (isActive) {
                if (d.dty_batch === '---') {
                    const displayStatus = d.last_status || '停機';
                    const isStop = displayStatus.includes('停機');
                    statusHtml = `<div class="status-indicator"><span class="dot" style="background:${isStop ? 'var(--danger)' : '#94a3b8'}"></span><span style="color:${isStop ? 'var(--danger)' : 'var(--text-muted)'}; font-weight:900;">${displayStatus}</span></div>`;
                } else {
                    statusHtml = `<div class="status-indicator"><span class="dot dot-running"></span><span style="color:var(--success); font-weight:900;">運行中</span></div>`;
                    if (d.last_status) {
                        const isStop = d.last_status.includes('停機');
                        statusHtml += `<div style="font-size:14px; color:${isStop ? 'var(--danger)' : 'var(--warning)'}; margin-top:8px; font-weight:800;">[${d.last_status}]</div>`;
                    }
                }
            } else {
                statusHtml = `<div style="font-size:15px; color:var(--text-muted); font-weight:700;">已完工 (${d.last_date})</div>`;
            }

            let poyHtml = '-';
            let supportHtml = '-';
            if (d.poy_list && d.poy_list.length > 0) {
                poyHtml = d.poy_list.map(p => {
                    const stockVal = parseFloat(p.stock_a_a2).toLocaleString(undefined, {minimumFractionDigits:1});
                    const color = p.stock_a_a2 < 500 ? 'var(--danger)' : 'var(--accent)';
                    return `
                        <div class="poy-info-box" onclick="clickPoy('${p.batch}')">
                            <div class="clickable-poy">${p.batch}</div>
                            <div class="spec-text">${p.spec}</div>
                            <div class="stock-label" style="color:${color}">庫存: ${stockVal} KG</div>
                            <div class="history-tag" title="${p.history}">履歷: ${p.history}</div>
                        </div>
                    `;
                }).join('');

                supportHtml = d.poy_list.map(p => {
                    const days = p.support_days;
                    const color = days < 3 ? 'var(--danger)' : 'var(--accent)';
                    return `<div style="margin-bottom:15px; height:90px; display:flex; align-items:center; justify-content:center; font-size:22px; color:${color}; font-weight:900; font-family:'Consolas';">
                        ${days >= 999 ? '-' : days.toFixed(2) + '<span style="font-size:14px; margin-left:4px; opacity:0.6;">天</span>'}
                    </div>`;
                }).join('');
            }

            let nextPlanHtml = "";
            if (d.next_plan) {
                const p = d.next_plan;
                const poyPart = p.poy_plan ? ` | <span style="color:#0369a1; font-weight:800;">${p.poy_plan}</span>` : "";
                nextPlanHtml = `
                    <div style="margin-top: 15px; padding-top: 12px; border-top: 2px dashed var(--border); font-size: 14px; color: var(--text-muted); background: var(--bg); padding: 10px; border-radius: 8px;">
                        <span style="background: var(--primary); padding: 2px 6px; border-radius: 4px; font-weight: 900; color: white; margin-right: 8px; font-size: 11px;">NEXT</span>
                        <span style="font-weight:700;">(${p.date_range})</span> | <span style="font-weight: 900; color:var(--primary);">${d.machine}${p.side_mark}</span> <span style="color: var(--accent); font-weight: 900;">${p.display_batch}</span>${poyPart}
                    </div>
                `;
            }

            const progressPct = Math.min(100, d.progress_pct || 0);
            const progressBarColor = progressPct >= 100 ? 'var(--success)' : (progressPct >= 80 ? 'var(--warning)' : 'var(--accent)');
            const targetDisplay = d.target_kg > 0 ? `<div class="progress-info"><span>TARGET: ${Math.round(d.target_kg).toLocaleString()} KG</span><span>${progressPct.toFixed(1)}%</span></div>` : "";

            tr.innerHTML = `
                <td data-label="機台"><span class="machine-badge">${d.machine} <span style="opacity:0.6; font-size:12px;">(${isActive ? d.current_sides.join('/') : d.produced_sides.join('/')})</span></span></td>
                <td data-label="狀態">${statusHtml}</td>
                <td data-label="批號 / 規格">
                    <span class="batch-text">${d.dty_batch}</span>
                    <span class="spec-text">${d.spec}</span>
                    ${nextPlanHtml}
                </td>
                <td data-label="POY庫存">${poyHtml}</td>
                <td data-label="預估支撐" style="text-align: center;">${supportHtml}</td>
                <td data-label="平均日產" style="text-align: center; font-weight:900; font-family:'Consolas'; color:${isActive && d.dty_batch !== '---' ? 'var(--accent)' : 'var(--text-muted)'}; font-size:22px;">
                    ${d.avg_td > 0 ? d.avg_td.toFixed(2) : '-'}
                </td>
                <td data-label="生產天數" style="text-align: center; font-weight:900; font-size:20px;">${d.days}</td>
                <td data-label="繳庫量">
                    <div class="stock-tags">${stockTags}</div>
                    ${totalStock}
                    ${d.target_kg > 0 ? `
                        <div class="progress-container">
                            <div class="progress-bar" style="width: ${progressPct}%; background: ${progressBarColor};"></div>
                        </div>
                        ${targetDisplay}
                    ` : ""}
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
    print("✨ 旗艦版戰情室更新成功！")
