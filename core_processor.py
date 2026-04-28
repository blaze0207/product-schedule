import pandas as pd
import glob
import os
import re
from datetime import datetime, timedelta

class MasterTableEngine:
    def __init__(self, base_dir):
        self.base_dir = base_dir
        
    def _get_latest_file(self, pattern):
        all_files = glob.glob(os.path.join(self.base_dir, pattern))
        valid_files = [f for f in all_files if not os.path.basename(f).startswith('~$')]
        if not valid_files:
            return None
        return sorted(valid_files, key=os.path.getmtime, reverse=True)[0]

    def standardize_batch(self, val, aggressive=False):
        """
        aggressive=False: 基礎清洗 (移除 FD, FP, LIKE, 空格)
        aggressive=True:  進階對接清洗 (額外移除結尾的字母裝飾碼)
        """
        if pd.isna(val):
            return ""
        s = str(val).strip().upper()
        s = s.replace("LIKE", "").strip()
        
        if s.startswith("FD") or s.startswith("FP"):
            match = re.search(r'\d', s)
            if match:
                s = s[match.start():]
        
        if aggressive:
            s = re.sub(r'(?<=\d)[A-Z]+$', '', s)
        return s

    def get_planned_tasks(self):
        file_path = self._get_latest_file("DTY*.xlsx")
        if not file_path:
            raise FileNotFoundError("找不到 DTY 排產表")
        
        xl = pd.ExcelFile(file_path)
        df = xl.parse(xl.sheet_names[0], header=0)
        cols = ['machine', 'batch_no', 'note', 'mark', 'dty_spec', 'date_range', 'scheduled_days', 't_d', 'poy_batch', 'poy_spec']
        df = df.iloc[:, :len(cols)]
        df.columns = cols
        
        df['machine'] = df['machine'].ffill()
        df['poy_batch'] = df['poy_batch'].ffill()
        df = df.dropna(subset=['batch_no', 'dty_spec'], how='all')
        
        def is_valid_machine(m):
            m = str(m).upper()
            if any(x in m for x in ['V', 'S2', '庫存', '待排', '機台', 'M4', '(CW)']):
                return False
            if re.search(r'M0[1-8]|S0[1-2]', m):
                return False
            return True
            
        df = df[df['machine'].apply(is_valid_machine)]
        
        tasks = {}
        for _, row in df.iterrows():
            m_orig = str(row['machine']).strip().upper()
            b_orig = str(row['batch_no']).strip().upper()
            bj = self.standardize_batch(b_orig, aggressive=True)
            key = (m_orig, bj)
            
            days = pd.to_numeric(row['scheduled_days'], errors='coerce') or 0
            td_p = pd.to_numeric(row['t_d'], errors='coerce') or 0
            
            if key not in tasks:
                tasks[key] = {
                    'machine': m_orig,
                    'batch_id': self.standardize_batch(b_orig, aggressive=False),
                    'batch_join': bj,
                    'batch_orig': b_orig,
                    'is_like': "LIKE" in b_orig,
                    'dty_spec': str(row['dty_spec']),
                    'poy_batch': str(row['poy_batch']).strip().upper(),
                    'poy_spec': str(row['poy_spec']),
                    'total_target_kg': 0,
                    'total_td_plan': 0,
                    'notes': [],
                    'marks': [],
                    'date_ranges': []
                }
            
            t = tasks[key]
            t['total_target_kg'] += days * td_p * 1000
            t['total_td_plan'] += td_p
            
            for k, f in [('notes', 'note'), ('marks', 'mark'), ('date_ranges', 'date_range')]:
                val = str(row[f]).strip().replace(' 00:00:00', '')
                if val and val != 'nan' and val not in t[k]:
                    t[k].append(val)
        return list(tasks.values())

class ActualsEngine:
    def __init__(self, base_dir, m_engine):
        self.base_dir = base_dir
        self.m_engine = m_engine

    def parse_row_by_side(self, row):
        """
        核心修正：改用 L (Index 11) 與 M (Index 12) 判定側邊
        """
        l_val = row.iloc[11] # L 欄位
        m_val = row.iloc[12] # M 欄位
        aa_val = row.iloc[26] # AA 欄位
        av_val = row.iloc[47] # AV 欄位
        
        av_num = pd.to_numeric(av_val, errors='coerce')
        av_num = 0 if pd.isna(av_num) else av_num
        aa_text = str(aa_val).strip() if not pd.isna(aa_val) else ""
        
        res_sides = {}
        # 標記身分 (SIDE_1 = A/L, SIDE_2 = B/R)
        for label, val in [('SIDE_1', l_val), ('SIDE_2', m_val)]:
            if pd.isna(val) or str(val).strip() == "":
                continue
            
            val_num = pd.to_numeric(val, errors='coerce')
            if pd.isna(val_num):
                # 第一層：是文字 (如停機)
                res_sides[label] = (str(val).strip(), 0)
            else:
                # 第二層：是數字 -> 產量採 AV，狀態視 AV 而定
                st = "" if av_num > 0 else (aa_text if aa_text else "待機")
                res_sides[label] = (st, av_num)
        return res_sides, av_num

    def get_actuals_data(self):
        all_files = glob.glob(os.path.join(self.base_dir, "*撚二科生產資訊.xlsx"))
        valid_files = [f for f in all_files if not os.path.basename(f).startswith('~$')]
        if not valid_files:
            return {}, {}, None
            
        file_path = sorted(valid_files, key=os.path.getmtime, reverse=True)[0]
        xl = pd.ExcelFile(file_path)
        df = xl.parse(xl.sheet_names[1])
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')
        latest_date = df.iloc[:, 0].max()
        
        history_list = []
        status_map = {}
        
        for _, row in df.iterrows():
            d = row.iloc[0]
            if pd.isna(d):
                continue
            
            m = str(row.iloc[1]).strip().upper()
            b_raw = str(row.iloc[2]).strip().upper()
            bj = self.m_engine.standardize_batch(b_raw, aggressive=True)
            if not m:
                continue
            
            parsed_sides, row_td = self.parse_row_by_side(row)
            side_keys = ["A", "B"] if re.match(r'^A(0[7-9]|1[0-9]|2[0-4])', m) else ["L", "R"]
            
            if bj:
                history_list.append({'date': d, 'm': m, 'bj': bj, 'td': row_td})
            
            if d == latest_date:
                for i, (label, (st, td)) in enumerate(parsed_sides.items()):
                    # label 會是 'SIDE_1' 或 'SIDE_2'
                    s_key = side_keys[0] if label == 'SIDE_1' else side_keys[1]
                    status_map[(m, s_key)] = {'batch': b_raw, 'status': st, 'td': td}
                    
                    if m not in status_map:
                        status_map[m] = {'batch': b_raw, 'status': st, 'td': row_td}
                    else:
                        curr = status_map[m]
                        if b_raw not in curr['batch']:
                            curr['batch'] += f"/{b_raw}"
                        if st and st not in curr['status']:
                            curr['status'] += f" {st}"

        avg_map = {}
        if history_list:
            df_h = pd.DataFrame(history_list)
            day_sum = df_h.groupby(['date', 'm', 'bj'])['td'].sum().reset_index()
            avg_df = day_sum[day_sum['td'] > 0].groupby(['m', 'bj'])['td'].mean().reset_index()
            for _, r in avg_df.iterrows():
                avg_map[(r['m'], r['bj'])] = r['td']
                
        return avg_map, status_map, latest_date

class InventoryEngine:
    def __init__(self, base_dir, m_engine):
        self.base_dir = base_dir
        self.m_engine = m_engine

    def get_stock_data(self):
        lisa_file = os.path.join(self.base_dir, "每日-最新庫存(DTY-LISA).xlsx")
        if not os.path.exists(lisa_file):
            dty_files = [f for f in glob.glob(os.path.join(self.base_dir, "DTY*.xlsx")) if not os.path.basename(f).startswith('~$')]
            if not dty_files: return {}, {}
            lisa_file = sorted(dty_files, key=os.path.getmtime, reverse=True)[0]
            df_lisa = pd.read_excel(lisa_file, sheet_name=1, header=None)
        else:
            df_lisa = pd.read_excel(lisa_file, header=None)
        
        stock_map = {}
        for _, row in df_lisa.iterrows():
            raw_b = str(row[1]).strip().upper()
            if not raw_b or raw_b == 'NAN':
                continue
            bj = self.m_engine.standardize_batch(raw_b, aggressive=True)
            grade = str(row[3]).strip().upper()
            qty = pd.to_numeric(row[11], errors='coerce') or 0
            if bj not in stock_map:
                stock_map[bj] = {'total': 0, 'grades': {}}
            stock_map[bj]['total'] += qty
            stock_map[bj]['grades'][grade] = stock_map[bj]['grades'].get(grade, 0) + qty
            
        poy_files = [f for f in glob.glob(os.path.join(self.base_dir, "絲八科-庫存表*.xlsx")) if not os.path.basename(f).startswith('~$')]
        poy_stock = {}
        if poy_files:
            df_poy = pd.read_excel(sorted(poy_files, key=os.path.getmtime, reverse=True)[0], sheet_name=-1)
            for _, row in df_poy.iterrows():
                p_id = str(row.iloc[0]).strip().upper()[:7]
                grade = str(row.iloc[3]).strip().upper()
                qty = pd.to_numeric(row.iloc[12], errors='coerce') or 0
                if grade in ['A', 'A2']:
                    if p_id not in poy_stock:
                        poy_stock[p_id] = {'total': 0, 'A': 0, 'A2': 0}
                    poy_stock[p_id]['total'] += qty
                    if grade == 'A': poy_stock[p_id]['A'] += qty
                    else: poy_stock[p_id]['A2'] += qty
        return stock_map, poy_stock

def generate_dashboard_data():
    path = os.getcwd()
    m_eng = MasterTableEngine(path)
    a_eng = ActualsEngine(path, m_eng)
    i_eng = InventoryEngine(path, m_eng)
    
    tasks = m_eng.get_planned_tasks()
    avg_map, status_map, latest_date = a_eng.get_actuals_data()
    stock_map, poy_stock = i_eng.get_stock_data()
    
    batch_machine_count = {}
    for t in tasks:
        bj = t['batch_join']
        batch_machine_count[bj] = batch_machine_count.get(bj, 0) + 1

    final_data = []
    for t in tasks:
        m, bj = t['machine'], t['batch_join']
        curr = None
        if len(t['marks']) == 1:
            curr = status_map.get((m, t['marks'][0]))
        if not curr:
            curr = status_map.get(m, {'batch': '未開機', 'status': '', 'td': 0})
        
        td_actual = avg_map.get((m, bj), 0)
        divisor = batch_machine_count.get(bj, 1)
        s_data = stock_map.get(bj, {'total': 0, 'grades': {}})
        
        stored_total = 0 if t['is_like'] else s_data['total'] / divisor
        display_grades = {} if t['is_like'] else {g: q/divisor for g, q in s_data['grades'].items() if q > 0}
        
        target = t['total_target_kg']
        demand = stored_total - target
        pct = min(100, (stored_total / target * 100)) if target > 0 else 0
        
        p_raw = t['poy_batch']
        p_parts = re.split(r'[>＞]', p_raw)
        p_info = []
        total_p_qty = 0
        for p in p_parts:
            p_s = p.strip()
            p_c = p_s[:7]
            if len(p_c) < 7 and p_raw.startswith("FP") and p_c[0].isdigit():
                p_c = "FP" + p_s[:5]
            info = poy_stock.get(p_c, {'total': 0, 'A': 0, 'A2': 0})
            p_info.append({'id': p_s, 'qty': info['total'], 'a': info['A'], 'a2': info['A2']})
            total_p_qty += info['total']
            
        poy_days = (total_p_qty / (td_actual * 1000)) if td_actual > 0 else 999
        
        finish_f = ""
        actual_b_upper = str(curr['batch']).upper()
        if bj in actual_b_upper and td_actual > 0 and demand < 0:
            days_to_go = abs(demand) / (td_actual * 1000)
            finish_f = (latest_date + timedelta(days=days_to_go)).strftime('%m/%d')
        elif demand >= 0:
            finish_f = "達標"

        final_data.append({
            'machine': m + (f" ({'+'.join(t['marks'])})" if t['marks'] else ""),
            'batch': t['batch_orig'],
            'actual_batch': curr['batch'],
            'status_text': curr['status'],
            'td_plan': t['total_td_plan'],
            'td_actual': td_actual,
            'target': target,
            'stored': stored_total,
            'grades': display_grades,
            'demand': demand,
            'pct': pct,
            'finish_forecast': finish_f,
            'poy': p_raw,
            'poy_details': p_info,
            'poy_days': poy_days,
            'spec': t['dty_spec'],
            'date_range': ", ".join(t['date_ranges']),
            'note': "; ".join(t['notes'])
        })
    return final_data

if __name__ == "__main__":
    generate_dashboard_data()
