import json
import os
import pandas as pd
import re
from datetime import datetime
from reality_analyzer import RealityLogAnalyzer
from export_cleaned_plan import export_clear_plan

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
    
    # 動態執行美化匯出並取得檔名，供 HTML 連結使用
    try:
        excel_filename = export_clear_plan()
    except Exception as e:
        print(f"匯出 Excel 失敗: {e}")
        excel_filename = "#"
    
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
            padding: 16px 30px; 
            display: flex; 
            justify-content: space-between; 
            align-items: center;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            position: relative; z-index: 100;
        }
        .top-header h1 { font-size: 24px; font-weight: 900; letter-spacing: 2px; display: flex; align-items: center; gap: 15px; }
        .top-header h1::before { content: ''; display: inline-block; width: 6px; height: 25px; background: var(--accent); border-radius: 4px; }
        
        .excel-btn {
            background: rgba(255,255,255,0.1);
            color: white;
            padding: 8px 16px;
            border-radius: 8px;
            text-decoration: none;
            font-size: 14px;
            font-weight: 700;
            border: 1px solid rgba(255,255,255,0.2);
            transition: all 0.3s;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .excel-btn:hover {
            background: var(--accent);
            border-color: var(--accent);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(59,130,246,0.4);
        }

        .header-info { text-align: right; font-family: 'Consolas', monospace; color: #cbd5e1; }
        .header-info div { font-size: 13px; margin-bottom: 2px; }
        .date-highlight { color: #60a5fa; font-weight: 800; }

        /* 回到頂部按鈕 */
        #back-to-top {
            position: fixed;
            bottom: 30px;
            right: 30px;
            width: 50px;
            height: 50px;
            background: var(--accent);
            color: white;
            border-radius: 50%;
            display: flex;
            justify-content: center;
            align-items: center;
            cursor: pointer;
            box-shadow: 0 4px 15px rgba(59, 130, 246, 0.4);
            z-index: 1000;
            opacity: 0;
            visibility: hidden;
            transition: all 0.3s;
            font-size: 24px;
        }
        #back-to-top.visible {
            opacity: 1;
            visibility: visible;
        }
        #back-to-top:hover {
            transform: translateY(-5px);
            background: #2563eb;
        }

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
        .filter-group { display: flex; gap: 15px; margin-bottom: 25px; flex-wrap: wrap; }
        .filter-btn { 
            padding: 12px 32px; border-radius: 12px; border: 2px solid var(--border); 
            background: white; cursor: pointer; font-weight: 800; color: var(--text-main);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); font-size: 18px;
        }
        .filter-btn:hover { border-color: var(--accent); background: var(--accent-soft); color: var(--accent); transform: translateY(-2px); }
        .filter-btn.active { background: var(--accent); border-color: var(--accent); color: white; box-shadow: 0 8px 15px rgba(59,130,246,0.3); }
        
        /* 搜尋引擎區域 */
        .search-area { display: flex; flex-direction: column; gap: 20px; }
        .search-row { display: flex; gap: 20px; align-items: stretch; flex-wrap: wrap; }
        .search-box { position: relative; flex: 1; min-width: 300px; }
        .search-input { 
            width: 100%; padding: 18px 60px 18px 50px; border-radius: 15px; 
            border: 2px solid var(--border); font-size: 18px; outline: none; transition: all 0.3s; background: #fcfdfe;
        }
        .search-box::before {
            content: '🔍'; position: absolute; left: 20px; top: 18px; font-size: 20px; opacity: 0.5;
        }
        .search-input:focus { border-color: var(--accent); background: white; box-shadow: 0 0 0 6px rgba(59,130,246,0.1); }
        .clear-btn { position: absolute; right: 25px; top: 18px; cursor: pointer; color: var(--danger); font-size: 24px; font-weight: 900; display: none; }
        
        .poy-search-box { flex: 0 0 400px; } /* POY 搜尋框稍微窄一點，與主搜尋區隔 */
        .poy-search-box::before { content: '📦'; }

        .analysis-panel { 
            width: 100%; border-radius: 15px; padding: 25px 30px; display: none; 
            background: linear-gradient(to right, #f0f9ff, #e0f2fe); border: 1px solid #bae6fd; animation: fadeIn 0.4s ease;
            margin-top: 10px;
        }
        @keyframes fadeIn { from { opacity: 0; scale: 0.98; } to { opacity: 1; scale: 1; } }

        /* 表格排版旗艦版 */
        .table-wrapper { background: white; border-radius: 20px; box-shadow: 0 10px 40px rgba(0,0,0,0.04); overflow: hidden; border: 1px solid var(--border); }
        table { width: 100%; border-collapse: collapse; table-layout: fixed; }
        /* 提升表頭字體大小至 15px 並增加間距 */
        th { 
            background: #f8fafc; padding: 22px 15px; text-align: left; 
            font-size: 15px; font-weight: 900; color: var(--primary); 
            text-transform: uppercase; letter-spacing: 0.1em; border-bottom: 4px solid var(--border); 
        }
        td { padding: 18px 12px; border-bottom: 1px solid #f1f5f9; font-size: 16px; vertical-align: middle; overflow-wrap: break-word; }

        tr.active { background: white; transition: background 0.2s; }
        tr.active:hover { background: #fcfdfe; }
        tr.history { opacity: 0.7; background: #fbfcfd; }

        .machine-badge { 
            background: var(--primary); color: white; padding: 6px 10px; border-radius: 6px; 
            font-weight: 900; font-size: 15px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);
            white-space: nowrap; display: inline-flex; align-items: center; gap: 6px;
        }

        
        .batch-text { font-weight: 900; color: var(--primary); font-size: 21px; display: block; margin-bottom: 8px; }
        .spec-text { font-size: 15px; color: var(--text-muted); font-weight: 500; }

        /* 昨日品質指標樣式 - 精簡版 */
        .quality-strip { 
            margin-top: 10px; display: flex; gap: 4px; flex-wrap: wrap;
        }
        .q-badge {
            padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: 800;
            display: inline-flex; align-items: center; gap: 3px; border: 1px solid rgba(0,0,0,0.05);
        }
        .q-good { background: #f0fdf4; color: #166534; }
        .q-warn { background: #fffbeb; color: #92400e; }
        .q-bad { background: #fef2f2; color: #991b1b; }
        .q-label { opacity: 0.6; font-weight: 500; font-size: 10px; }

        /* POY 整合互動膠囊 */
        .poy-wrapper { display: flex; flex-direction: column; gap: 10px; }
        .poy-card { 
            padding: 12px; border-radius: 12px; background: #f8fafc; 
            border-left: 5px solid #cbd5e1; display: flex; justify-content: space-between; align-items: center;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); cursor: pointer; position: relative;
        }
        .poy-card:hover { border-left-color: var(--accent); background: var(--accent-soft); transform: translateX(4px); }
        .poy-main { flex: 1; }
        .poy-support { 
            text-align: right; border-left: 2px dashed var(--border); padding-left: 15px; margin-left: 15px;
            min-width: 85px;
        }
        .support-val { font-family: 'Consolas', monospace; font-size: 20px; font-weight: 900; }
        .support-unit { font-size: 11px; opacity: 0.6; margin-left: 2px; }

        .clickable-poy { color: var(--accent); font-weight: 900; font-family: 'Consolas', monospace; font-size: 17px; }
        .stock-label { font-size: 14px; font-weight: 800; margin-top: 4px; color: var(--text-main); }
        .history-tag { font-size: 11px; color: var(--text-muted); margin-top: 4px; display: block; opacity: 0.7; }

        .tag { font-size: 13px; padding: 4px 10px; border-radius: 6px; font-weight: 800; display: inline-block; }
        .tag-a { background: #dcfce7; color: #065f46; }
        .tag-ax { background: #fef3c7; color: #92400e; }
        .tag-b { background: #ffedd5; color: #9a3412; }
        .tag-c { background: #fee2e2; color: #991b1b; }
        .total-stock { font-size: 22px; font-weight: 900; color: var(--primary); margin-top: 10px; display: block; }

        /* 備註區塊樣式 */
        .remark-box {
            margin-top: 10px; padding: 10px 12px; border-radius: 8px;
            background: #f1f5f9; border-left: 4px solid #94a3b8;
            font-size: 12px; color: var(--text-main); font-weight: 500;
            line-height: 1.5; overflow-wrap: break-word;
        }
        .remark-tag { 
            font-weight: 900; color: var(--text-muted); font-size: 10px; 
            text-transform: uppercase; margin-right: 6px; 
        }

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
                display: grid; grid-template-columns: 110px 1fr; gap: 10px; 
                padding: 12px 10px; border: none; border-bottom: 1px solid #f1f5f9;
                text-align: left !important; min-height: 50px; align-items: center;
            }
            td::before { 
                content: attr(data-label); font-size: 13px; font-weight: 900; 
                color: var(--text-muted); text-transform: uppercase; padding-right: 5px;
            }
            td[style*="text-align: center"] { justify-content: start !important; }
            .poy-info-box { margin-bottom: 10px; }
            .batch-text { font-size: 18px; }
        }
    </style>
</head>
<body>
    <header class="top-header">
        <div style="display: flex; align-items: center; gap: 20px;">
            <h1>撚二科生產戰情室</h1>
            <a href="VAR_EXCEL_LINK" class="excel-btn" target="_blank" title="點擊開啟/下載精簡版 Excel 檔案">
                📂 原始產銷(精簡版)
            </a>
        </div>
        <div class="header-info">
            <div>生產資訊日: <span class="date-highlight">VAR_PROD_DATE</span></div>
            <div>DTY/POY庫存: <span class="date-highlight">VAR_STOCK_DATE</span></div>
        </div>
    </header>

    <div id="back-to-top" title="回到頂部">▲</div>

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
            <!-- 第一排：工程師篩選 -->
            <div class="filter-group">
                <span style="font-size:14px; font-weight:900; color:var(--text-muted); width:100%; margin-bottom:5px; display:block;">工程師篩選:</span>
                <button class="filter-btn active" id="btn-eng-all" onclick="setEngineer('all')">🌐 全部</button>
                <button class="filter-btn" id="btn-eng-luo" onclick="setEngineer('luo')">👤 羅工</button>
                <button class="filter-btn" id="btn-eng-wu" onclick="setEngineer('wu')">👤 吳工</button>
            </div>
            
            <!-- 第二排：狀態篩選 -->
            <div class="filter-group" style="margin-top: 20px;">
                <span style="font-size:14px; font-weight:900; color:var(--text-muted); width:100%; margin-bottom:5px; display:block;">機台狀態篩選:</span>
                <button class="filter-btn active" id="btn-stat-all" onclick="setStatus('all')">🌐 全部</button>
                <button class="filter-btn" id="btn-stat-run" onclick="setStatus('run')">🟢 運行中</button>
                <button class="filter-btn" id="btn-stat-stop" onclick="setStatus('stop')">🔴 停機/異常</button>
                <button class="filter-btn" id="btn-stat-off" onclick="setStatus('off')">⚪ 已完工</button>
            </div>
            
            <div class="search-area" style="margin-top: 30px;">
                <div class="search-row">
                    <!-- 通用搜尋框 -->
                    <div class="search-box">
                        <input type="text" id="generalSearch" class="search-input" placeholder="輸入機台(如A07)或DTY批號追蹤...." oninput="applyFilters()">
                        <span id="clearGeneral" class="clear-btn" onclick="clearInput('generalSearch')">&times;</span>
                    </div>
                    <!-- POY 分析框 -->
                    <div class="search-box poy-search-box">
                        <input type="text" id="poySearch" class="search-input" placeholder="輸入或點擊 POY 批號分析..." oninput="applyFilters()">
                        <span id="clearPoy" class="clear-btn" onclick="clearInput('poySearch')">&times;</span>
                    </div>
                </div>
                <div id="analysisPanel" class="analysis-panel"></div>
            </div>
        </section>

        <div class="table-wrapper">
            <table>
                <thead>
                    <tr>
                        <th style="width: 8%">機台</th>
                        <th style="width: 10%">狀態</th>
                        <th style="width: 18%">批號 / 規格</th>
                        <th style="width: 28%">物料狀態 (POY/支撐天數)</th>
                        <th style="width: 12%; text-align: center;">日產(T) / 天數</th>
                        <th style="width: 24%">繳庫進度</th>
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
        let currentEngineer = 'all';
        let currentStatus = 'all';

        const engineerMaps = {
            'luo': ['A07','A08','A09','A13','A14','A15','A19','A20','A21','J07','T04','T05','T06'],
            'wu': ['A10','A11','A12','A16','A17','A18','A22','A23','A24','J08','T01','T02','T03']
        };

        function setEngineer(type) {
            currentEngineer = type;
            document.querySelectorAll('[id^="btn-eng-"]').forEach(btn => btn.classList.remove('active'));
            document.getElementById('btn-eng-'+type)?.classList.add('active');
            applyFilters();
        }

        function setStatus(type) {
            currentStatus = type;
            document.querySelectorAll('[id^="btn-stat-"]').forEach(btn => btn.classList.remove('active'));
            document.getElementById('btn-stat-'+type)?.classList.add('active');
            applyFilters();
        }

        function clearInput(id) {
            document.getElementById(id).value = '';
            applyFilters();
        }

        function cleanPoyId(val) {
            if (!val) return "";
            return val.trim().toUpperCase().replace(/[×*X]\d+$/, '').trim();
        }

        function applyFilters() {
            const genQuery = document.getElementById('generalSearch').value.trim().toUpperCase();
            const poyQueryRaw = document.getElementById('poySearch').value.trim().toUpperCase();
            const poyQueryClean = cleanPoyId(poyQueryRaw);
            
            document.getElementById('clearGeneral').style.display = genQuery ? 'block' : 'none';
            document.getElementById('clearPoy').style.display = poyQueryRaw ? 'block' : 'none';

            // 執行表格過濾
            Array.from(tbody.rows).forEach((row, index) => {
                const data = rawData[index]; // 對應原始數據以獲得 isActive 與 dty_batch
                const badge = row.querySelector('.machine-badge');
                if(!badge || !data) return;
                
                // 1. 工程師過濾
                const mMatch = badge.innerText.match(/^[A-Z]\d+/);
                const mId = mMatch ? mMatch[0] : "";
                const engMatch = (currentEngineer === 'all' || engineerMaps[currentEngineer].includes(mId));
                
                // 2. 狀態過濾
                let statMatch = true;
                if (currentStatus === 'run') statMatch = (data.is_active && data.dty_batch !== '---');
                else if (currentStatus === 'stop') statMatch = (data.is_active && data.dty_batch === '---');
                else if (currentStatus === 'off') statMatch = (!data.is_active);
                
                // 3. 通用搜尋過濾
                const rowText = row.innerText.toUpperCase();
                const textMatch = !genQuery || rowText.includes(genQuery);
                
                // 4. POY 過濾
                const poyMatch = !poyQueryRaw || rowText.includes(poyQueryRaw);

                row.style.display = (engMatch && statMatch && textMatch && poyMatch) ? '' : 'none';
            });

            // POY 分析面板邏輯
            const panel = document.getElementById('analysisPanel');
            if (poyQueryRaw) {
                let matchedKey = Object.keys(poyAnalysis).find(k => k === poyQueryClean);
                if (!matchedKey && poyQueryClean.length >= 3) {
                    matchedKey = Object.keys(poyAnalysis).find(k => k.includes(poyQueryClean));
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
            } else { panel.style.display = 'none'; }
        }

        function clickPoy(batch) {
            document.getElementById('poySearch').value = batch;
            applyFilters();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        rawData.forEach(d => {
            const tr = document.createElement('tr');
            const isActive = d.is_active;
            tr.className = isActive ? 'active' : 'history';
            
            // 繳庫量與標籤
            const s = d.stock || {A:0, AX:0, B:0, C:0};
            let stockTags = "";
            if (s.A > 0) stockTags += `<span class="tag tag-a">A:${s.A.toFixed(0)}</span> `;
            if (s.AX > 0) stockTags += `<span class="tag tag-ax">AX:${s.AX.toFixed(0)}</span> `;
            if (s.B > 0) stockTags += `<span class="tag tag-b">B:${s.B.toFixed(0)}</span> `;
            if (s.C > 0) stockTags += `<span class="tag tag-c">C:${s.C.toFixed(0)}</span>`;
            const totalStock = d.stock_summary?.total_deposit > 0 ? `<span class="total-stock">總:${d.stock_summary.total_deposit.toLocaleString()}</span>` : '-';

            // 狀態邏輯
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

            // POY 整合卡片邏輯
            let poyCombinedHtml = '-';
            if (d.poy_list && d.poy_list.length > 0) {
                poyCombinedHtml = `<div class="poy-wrapper">` + d.poy_list.map(p => {
                    const stockVal = parseFloat(p.stock_a_a2).toLocaleString(undefined, {minimumFractionDigits:1});
                    const sDays = p.support_days;
                    const stockColor = p.stock_a_a2 < 500 ? 'var(--danger)' : 'var(--text-main)';
                    const dayColor = sDays < 3 ? 'var(--danger)' : 'var(--accent)';
                    
                    return `
                        <div class="poy-card" onclick="clickPoy('${p.batch}')">
                            <div class="poy-main">
                                <div class="clickable-poy">${p.batch}</div>
                                <div class="spec-text" style="font-size:12px;">${p.spec}</div>
                                <div class="stock-label" style="color:${stockColor}">庫存: ${stockVal} KG</div>
                            </div>
                            <div class="poy-support">
                                <div style="font-size:11px; font-weight:800; color:var(--text-muted); margin-bottom:2px;">支撐預估</div>
                                <div class="support-val" style="color:${dayColor}">${sDays >= 999 ? '-' : sDays.toFixed(1)}<span class="support-unit">天</span></div>
                            </div>
                        </div>
                    `;
                }).join('') + `</div>`;
            }

            // 接續計畫
            let nextPlanHtml = "";
            if (d.next_plan) {
                const p = d.next_plan;
                nextPlanHtml = `
                    <div style="margin-top: 15px; padding-top: 12px; border-top: 2px dashed var(--border); font-size: 13px; color: var(--text-muted); background: var(--bg); padding: 10px; border-radius: 8px;">
                        <span style="background: var(--primary); padding: 2px 6px; border-radius: 4px; font-weight: 900; color: white; margin-right: 8px; font-size: 10px;">NEXT</span>
                        <span style="font-weight:700;">(${p.date_range})</span> | <span style="font-weight: 900; color:var(--primary);">${d.machine}${p.side_mark}</span> <span style="color: var(--accent); font-weight: 900;">${p.display_batch}</span>
                    </div>
                `;
            }

            // 備註顯示
            let remarkHtml = "";
            if (d.remark) {
                remarkHtml = `
                    <div class="remark-box">
                        <span class="remark-tag">📝 備註</span>
                        ${d.remark}
                    </div>
                `;
            }

            // 品質指標
            let qualityHtml = "";
            if (d.quality && (d.quality.a_rate !== null || d.quality.fixed_weight_rate !== null)) {
                const aq = d.quality;
                let aClass = "q-bad";
                if (aq.a_rate >= 97.2) aClass = "q-good";
                else if (aq.a_rate >= 95.0) aClass = "q-warn";
                const fClass = (aq.fixed_weight_rate >= 90.5) ? "q-good" : "q-bad";
                
                qualityHtml = `
                    <div class="quality-strip">
                        ${aq.a_rate !== null ? `<span class="q-badge ${aClass}"><span class="q-label">A級:</span>${aq.a_rate.toFixed(1)}%</span>` : ""}
                        ${aq.fixed_weight_rate !== null ? `<span class="q-badge ${fClass}"><span class="q-label">定重:</span>${aq.fixed_weight_rate.toFixed(1)}%</span>` : ""}
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
                    ${qualityHtml}
                    ${nextPlanHtml}
                    ${remarkHtml}
                </td>
                <td data-label="物料與支撐">${poyCombinedHtml}</td>
                <td data-label="效率與天數" style="text-align: center;">
                    <div style="font-weight:900; font-family:'Consolas'; color:${isActive && d.dty_batch !== '---' ? 'var(--accent)' : 'var(--text-muted)'}; font-size:24px;">
                        ${d.avg_td > 0 ? d.avg_td.toFixed(2) : '-'}
                    </div>
                    <div style="font-size:12px; color:var(--text-muted); font-weight:800; margin-top:5px;">${d.days} DAYS</div>
                </td>
                <td data-label="繳庫進度">
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

        // 回到頂部邏輯
        const backToTop = document.getElementById('back-to-top');
        window.onscroll = function() {
            if (document.body.scrollTop > 300 || document.documentElement.scrollTop > 300) {
                backToTop.classList.add('visible');
            } else {
                backToTop.classList.remove('visible');
            }
        };
        backToTop.onclick = function() {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        };
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
                              .replace("VAR_SETTLE_DATE", summary_prod_date) \
                              .replace("VAR_EXCEL_LINK", excel_filename)

    with open('production_dashboard.html', 'w', encoding='utf-8') as f:
        f.write(final_html)

if __name__ == "__main__":
    analyzer = RealityLogAnalyzer(os.getcwd())
    result = analyzer.get_reality_tasks()
    generate_v3_html(result)
    print("✨ 旗艦版戰情室更新成功！")
