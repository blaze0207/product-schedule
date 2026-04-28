import pandas as pd
import glob
import os
import re
from datetime import datetime

class RealityLogAnalyzer:
    def __init__(self, base_dir):
        self.base_dir = base_dir
        self.status_keywords = ["停機", "改紡", "了機", "清車", "檢修", "待料", "待機"]

    def standardize_id(self, val):
        if pd.isna(val): return ""
        s = str(val).strip().upper()
        if any(k in s for k in self.status_keywords): return s
        s = s.replace("FD", "").replace("FP", "").strip()
        if "LIKE" in s:
            parts = s.split("LIKE")
            base = parts[0].strip()
            like_part = " LIKE"
            base = re.sub(r'(?<=\d)[A-Z]+$', '', base)
            s = base + like_part
        else:
            s = re.sub(r'(?<=\d)[A-Z]+$', '', s)
        return s

    def clean_poy_id(self, val):
        if pd.isna(val): return ""
        s = str(val).strip().upper()
        s = re.sub(r'[×*X]\d+$', '', s)
        return s.strip()

    def get_inventory_key(self, val):
        if pd.isna(val): return ""
        s = str(val).strip().upper()
        s = s.replace("FD", "").replace("FP", "").strip()
        like_suffix = ""
        if "LIKE" in s:
            s = s.split("LIKE")[0].strip()
            like_suffix = " LIKE"
        if len(s) > 2 and (s.startswith("2F") or s.startswith("2G") or s.startswith("8G")):
            s = s[2:]
        s = re.sub(r'(?<=\d)[A-Z]+$', '', s)
        return s + like_suffix

    def get_stock_data(self):
        file_name = '每日-最新庫存(DTY-LISA).xlsx'
        if not os.path.exists(os.path.join(self.base_dir, file_name)):
            return {}
        df = pd.read_excel(os.path.join(self.base_dir, file_name), sheet_name='總庫存')
        stock_map = {}
        for _, row in df.iterrows():
            raw_batch = str(row.iloc[1]).strip().upper()
            if raw_batch == 'NAN' or raw_batch == '': continue
            join_key = self.get_inventory_key(raw_batch)
            grade = "Unknown"
            b = raw_batch[2:] if raw_batch.startswith('FD') else raw_batch
            if (b.startswith('2G') or b.startswith('8G')) and b.endswith('C'): grade = "C"
            elif b.startswith('2G') and b.endswith('8'): grade = "B"
            elif b.startswith('2G'): grade = "AX"
            else: grade = "A" 
            weight = pd.to_numeric(row.iloc[11], errors='coerce') or 0
            if join_key not in stock_map: stock_map[join_key] = {'A': 0, 'AX': 0, 'B': 0, 'C': 0}
            if grade in stock_map[join_key]: stock_map[join_key][grade] += weight
        return stock_map

    def get_monthly_summary(self):
        file_name = '每日-最新庫存(DTY-LISA).xlsx'
        path = os.path.join(self.base_dir, file_name)
        if not os.path.exists(path): return 0, "Unknown"
        xl = pd.ExcelFile(path)
        df_l1 = xl.parse('總庫存', header=None, nrows=1)
        raw_l1 = str(df_l1.iloc[0, 11]) if df_l1.shape[1] > 11 else ""
        date_match = re.search(r'\d{4}/\d{2}/\d{2}', raw_l1)
        update_date = date_match.group(0) if date_match else "Unknown"
        df_total = xl.parse('總庫存')
        total_h = pd.to_numeric(df_total.iloc[-1, 7], errors='coerce') or 0
        df_purchased = xl.parse('外購')
        purchased_h = pd.to_numeric(df_purchased.iloc[-1, 7], errors='coerce') or 0
        return total_h - purchased_h, update_date

    def get_poy_data(self):
        files = [f for f in glob.glob(os.path.join(self.base_dir, "絲八科-庫存表*.xlsx")) if not os.path.basename(f).startswith('~$')]
        if not files: return {}
        file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
        xl = pd.ExcelFile(file_path)
        last_sheet = xl.sheet_names[-1]
        df = xl.parse(last_sheet)
        poy_map = {}
        for _, row in df.iterrows():
            batch = str(row.iloc[0]).strip().upper()
            spec = str(row.iloc[2]).strip()
            grade = str(row.iloc[3]).strip().upper()
            weight = pd.to_numeric(row.iloc[12], errors='coerce') or 0
            history = str(row.iloc[-1]).strip()
            if batch == 'NAN' or batch == '': continue
            if grade not in ['A', 'A2']: continue
            if batch not in poy_map: poy_map[batch] = {'spec': spec, 'stock_a_a2': 0, 'histories': set(), 'grades': set()}
            poy_map[batch]['stock_a_a2'] += weight
            poy_map[batch]['grades'].add(grade)
            if history and history != 'NAN': poy_map[batch]['histories'].add(history)
        for b in poy_map:
            poy_map[b]['history_text'] = " / ".join(sorted(list(poy_map[b]['histories'])))
            poy_map[b]['grade_text'] = "+".join(sorted(list(poy_map[b]['grades'])))
        return poy_map

    def get_plan_data(self):
        files = [f for f in glob.glob(os.path.join(self.base_dir, "*產銷*.xlsx")) if not os.path.basename(f).startswith('~$')]
        if not files: return {}, {}, {}
        file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
        xl = pd.ExcelFile(file_path)
        df = xl.parse(xl.sheet_names[0])
        df.iloc[:, 0] = df.iloc[:, 0].ffill()
        df.iloc[:, 2] = df.iloc[:, 2].ffill()
        plan_map = {}
        target_aggregate = {}
        batch_total_target = {}
        m_blacklist = ['S1-2', 'M4(CW)', '待排產', '庫存不排產', 'V']
        for _, row in df.iterrows():
            m = str(row.iloc[0]).strip().upper()
            batch_raw = str(row.iloc[1]).strip().upper()
            side_raw = str(row.iloc[3]).strip().upper()
            date_range = str(row.iloc[5]).strip()
            days_val = pd.to_numeric(row.iloc[6], errors='coerce') or 0
            td_val = pd.to_numeric(row.iloc[7], errors='coerce') or 0
            if any(k in m for k in m_blacklist): continue
            if pd.isna(row.iloc[6]): continue 
            if batch_raw in ['NAN', '']: continue
            clean_batch = self.standardize_id(batch_raw)
            target_key = (m, clean_batch)
            kg = days_val * td_val * 1000
            target_aggregate[target_key] = target_aggregate.get(target_key, 0) + kg
            batch_total_target[clean_batch] = batch_total_target.get(clean_batch, 0) + kg
            if side_raw == "A": side_mark = "(A)"
            elif side_raw == "B": side_mark = "(B)"
            else: side_mark = "(A+B)"
            display_batch = batch_raw.replace("FD", "").strip()
            poy_plan = str(row.iloc[8]).strip().upper() if not pd.isna(row.iloc[8]) else ""
            remark = str(row.iloc[2]).strip() if not pd.isna(row.iloc[2]) and str(row.iloc[2]).strip() != '0' else ""
            if m not in plan_map: plan_map[m] = []
            plan_map[m].append({'batch_key': clean_batch, 'display_batch': display_batch, 'poy_plan': poy_plan, 'side_mark': side_mark, 'date_range': date_range, 'remark': remark})
        return plan_map, target_aggregate, batch_total_target

    def get_daily_report_data(self):
        files = [f for f in glob.glob(os.path.join(self.base_dir, "生產日報表(*).xlsx")) if not os.path.basename(f).startswith('~$')]
        if not files: return {}
        file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]
        try:
            df = pd.read_excel(file_path, sheet_name='日報表', skiprows=2)
            quality_map = {}
            for _, row in df.iterrows():
                batch_raw = str(row.iloc[0]).strip().upper()
                if batch_raw == 'NAN' or not batch_raw or batch_raw == '0': continue
                clean_batch = self.standardize_id(batch_raw)
                a_rate = pd.to_numeric(row.iloc[10], errors='coerce')
                fixed_weight_rate = pd.to_numeric(row.iloc[13], errors='coerce')
                if clean_batch not in quality_map: quality_map[clean_batch] = {'a_rate': a_rate * 100 if not pd.isna(a_rate) else None, 'fixed_weight_rate': fixed_weight_rate * 100 if not pd.isna(fixed_weight_rate) else None}
            return quality_map
        except Exception as e:
            print(f"讀取日報表失敗: {e}"); return {}

    def get_reality_tasks(self):
        stock_data = self.get_stock_data(); poy_data = self.get_poy_data(); plan_data, target_aggregate, batch_total_target = self.get_plan_data()
        summary_val, stock_update_date = self.get_monthly_summary(); quality_data = self.get_daily_report_data()
        files = [f for f in glob.glob(os.path.join(self.base_dir, "*撚二科生產資訊.xlsx")) if not os.path.basename(f).startswith('~$')]
        if not files: raise FileNotFoundError("找不到生產資訊檔案")
        file_path = sorted(files, key=os.path.getmtime, reverse=True)[0]; xl = pd.ExcelFile(file_path); df = xl.parse(xl.sheet_names[1])
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce'); latest_date = df.iloc[:, 0].max()
        def is_valid_machine(m):
            m = str(m).upper(); return False if (not m or m == 'NAN' or re.search(r'M0[1-8]|S0[1-2]', m)) else True
        df['m_name'] = df.iloc[:, 1].astype(str).str.strip().str.upper(); df = df[df['m_name'].apply(is_valid_machine)].copy()
        df['occ'] = df.groupby([df.iloc[:, 0], 'm_name']).cumcount() + 1; df['d_count'] = df.groupby([df.iloc[:, 0], 'm_name'])['m_name'].transform('count')
        tasks = {}
        for _, row in df.iterrows():
            d = row.iloc[0]; m = row['m_name']; batch_raw = str(row.iloc[2]).strip().upper()
            if batch_raw in ['NAN', '']: continue
            is_status_only = any(k in batch_raw for k in self.status_keywords); dty_std = "__STATUS__" if is_status_only else self.standardize_id(batch_raw); key = (m, dty_std)
            occ, d_count = row['occ'], row['d_count']; at_val_44 = str(row.iloc[44]).strip(); at_val_45 = str(row.iloc[45]).strip(); is_full = (at_val_44 == "全" or at_val_45 == "全")
            sides = []
            if d_count >= 2:
                if occ == 1: sides.append("A" if re.match(r'^A', m) else "L")
                elif occ == 2: sides.append("B" if re.match(r'^A', m) else "R")
            else:
                if is_full: sides = ["A", "B"] if re.match(r'^A', m) else ["L", "R"]
                else:
                    a_v = [row.iloc[11], row.iloc[15], row.iloc[19], row.iloc[23]]; b_v = [row.iloc[12], row.iloc[16], row.iloc[20], row.iloc[24]]
                    if any(not pd.isna(v) and str(v).strip() != "" for v in a_v): sides.append("A" if re.match(r'^A', m) else "L")
                    if any(not pd.isna(v) and str(v).strip() != "" for v in b_v): sides.append("B" if re.match(r'^A', m) else "R")
            if key not in tasks:
                tasks[key] = {'machine': m, 'dty_batch': "---" if is_status_only else batch_raw, 'dty_std': dty_std, 'poy_list': [], 'spec': str(row.iloc[4]).strip() if not is_status_only else "---", 'sides': set(), 'active_sides': set(), 'days': 0, 'last_status': "", 'last_td': 0, 'total_td': 0, 'is_active': False, 'last_date': d, 'stock': {'A': 0, 'AX': 0, 'B': 0, 'C': 0}}
            t = tasks[key]; td = pd.to_numeric(row.iloc[47], errors='coerce') or 0; aa = str(row.iloc[26]).strip() if not pd.isna(row.iloc[26]) else ""
            l_val, r_val = str(row.iloc[11]).strip(), str(row.iloc[12]).strip()
            l_is_status, r_is_status = any(k in l_val for k in self.status_keywords), any(k in r_val for k in self.status_keywords)
            if re.match(r'^-?\d+(\.\d+)?$', aa): aa = ""
            if is_status_only: eff_status = batch_raw
            else:
                inv_key = self.get_inventory_key(batch_raw)
                if inv_key in stock_data: t['stock'] = stock_data[inv_key]
                eff_status = "全台停機" if (l_is_status and r_is_status) else ("半邊停/了機" if (l_is_status or r_is_status) else ("" if any(k in aa for k in self.status_keywords) else aa))
            poy_batch_raw = str(row.iloc[5]).strip().upper()
            if poy_batch_raw and poy_batch_raw != 'NAN' and not is_status_only:
                poy_clean_id = self.clean_poy_id(poy_batch_raw)
                if not any(p['batch'] == poy_batch_raw for p in t['poy_list']):
                    p_info = {'batch': poy_batch_raw, 'clean_id': poy_clean_id, 'stock_a_a2': 0, 'spec': '---', 'history': '---', 'grades': '---'}
                    if poy_clean_id in poy_data:
                        ext = poy_data[poy_clean_id]; p_info.update({'stock_a_a2': ext['stock_a_a2'], 'spec': ext['spec'], 'history': ext['history_text'], 'grades': ext['grade_text']})
                    t['poy_list'].append(p_info)
            if td > 0: t['days'] += 1; t['total_td'] += td
            if d == latest_date:
                t['is_active'] = True
                if eff_status and eff_status not in t['last_status']: t['last_status'] = (t['last_status'] + " " + eff_status).strip()
                t['last_td'] = td
                if sides: t['active_sides'].update(sides)
            if sides: t['sides'].update(sides)
            if d > t['last_date']: t['last_date'] = d
        for t in tasks.values():
            t['produced_sides'] = sorted(list(t['sides'])); t['current_sides'] = sorted(list(t['active_sides'])); t['avg_td'] = t['total_td'] / t['days'] if t['days'] > 0 else 0
            s = t['stock']; t['stock_summary'] = {'deposit': s['A'] + s['AX'], 'total_deposit': s['A'] + s['AX'] + s['B'] + s['C']}
            del t['sides']; del t['active_sides']
        poy_analysis = {}
        for t in tasks.values():
            if not t['is_active']: continue
            for p in t['poy_list']:
                pid = p['clean_id']
                if pid not in poy_analysis: poy_analysis[pid] = {'machines': [], 'total_avg_td': 0, 'stock_a_a2': p['stock_a_a2'], 'spec': p['spec']}
                if t['machine'] not in poy_analysis[pid]['machines']: poy_analysis[pid]['machines'].append(t['machine']); poy_analysis[pid]['total_avg_td'] += t['avg_td']
        for pid in poy_analysis:
            pa = poy_analysis[pid]; daily_kg = pa['total_avg_td'] * 1000; pa['support_days'] = pa['stock_a_a2'] / daily_kg if daily_kg > 0 else 999; pa['machines_text'] = ", ".join(sorted(pa['machines']))
        for t in tasks.values():
            for p in t['poy_list']:
                pid = p['clean_id']; p['support_days'] = poy_analysis[pid]['support_days'] if pid in poy_analysis else 999
        unique_dates = sorted(df.iloc[:, 0].dropna().unique(), reverse=True); summary_prod_date_obj = unique_dates[1] if len(unique_dates) > 1 else latest_date; latest_daily_sum = df[df.iloc[:, 0] == summary_prod_date_obj].iloc[:, 47].apply(pd.to_numeric, errors='coerce').sum()
        for t in tasks.values():
            m = t['machine']; dty_std = t['dty_std']; t['quality'] = quality_data.get(dty_std, {'a_rate': None, 'fixed_weight_rate': None})
            t['target_kg'] = target_aggregate.get((m, dty_std), 0)
            if t['target_kg'] == 0: t['target_kg'] = batch_total_target.get(dty_std, 0)
            actual_kg = t['stock_summary']['total_deposit']; t['progress_pct'] = (actual_kg / t['target_kg'] * 100) if t['target_kg'] > 0 else 0; t['next_plan'] = None; t['remark'] = ""
            if m in plan_data:
                p_queue = plan_data[m]; current_idx = -1
                for i, p_item in enumerate(p_queue):
                    if p_item['batch_key'] == dty_std:
                        current_idx = i; t['remark'] = p_item.get('remark', ""); break
                if current_idx != -1 and current_idx + 1 < len(p_queue): t['next_plan'] = p_queue[current_idx + 1]
                elif current_idx == -1 and len(p_queue) > 0: t['next_plan'] = p_queue[0]
        return {'tasks': list(tasks.values()), 'self_produced_weight': summary_val, 'stock_update_date': stock_update_date, 'production_date': latest_date.strftime('%Y/%m/%d') if pd.notnull(latest_date) else "Unknown", 'poy_analysis': poy_analysis, 'latest_daily_sum': latest_daily_sum, 'summary_prod_date': summary_prod_date_obj.strftime('%Y/%m/%d') if pd.notnull(summary_prod_date_obj) else "Unknown"}
