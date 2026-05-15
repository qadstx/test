import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v28.1", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

# --- Utility Functions ---
def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABTN])(\d+)', text)
    if match: return f"{match.group(1)}{match.group(2)}"
    return text

def clean_number(value):
    try:
        if value is None or pd.isna(value): return None
        num = abs(float(value))
        return math.floor(num * 100000) / 100000
    except: return None

def get_writable_cell(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

def get_merged_origin_value(ws, cell):
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return cell.value

def parse_spec(ws, row):
    try:
        nom_raw = get_merged_origin_value(ws, ws.cell(row=row, column=3))
        nom_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(nom_raw))
        nominal = float(nom_match[0]) if nom_match else 0.0
        tol_raw = get_merged_origin_value(ws, ws.cell(row=row, column=7))
        tol_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(tol_raw))
        tol = float(tol_match[0]) if tol_match else 0.0
        return nominal - tol, nominal + tol
    except: return None, None

# --- Core Logic ---
def scan_lat_data(df):
    lat_map = {}
    # B1-B32 และ N10 (Vertical)
    sections = [{'name_row': 11, 'data_start': 29}, {'name_row': 384, 'data_start': 401}]
    for sec in sections:
        if sec['name_row'] >= len(df): continue
        names = df.iloc[sec['name_row']]
        for col_idx, raw_name in enumerate(names):
            bid = get_base_id(raw_name)
            if bid and bid not in ["N7"]:
                d = [clean_number(df.iloc[sec['data_start']+i, col_idx]) for i in range(3)]
                if bid not in lat_map: lat_map[bid] = []
                lat_map[bid].append(d)

    # B41-B80 (Horizontal Scan เหมือน N7)
    header_row, data_rows = 65, [83, 84, 85]
    if data_rows[2] < len(df):
        headers = df.iloc[header_row]
        for c_idx in range(1, len(headers)):
            bid = get_base_id(headers.iloc[c_idx])
            if bid and bid.startswith("B"):
                d = [clean_number(df.iloc[dr, c_idx]) for dr in data_rows]
                if bid not in lat_map: lat_map[bid] = []
                lat_map[bid].append(d)

    # N7 (Hard-coded Range)
    n7_data = []
    ranges = [(295, 298, 1, 33), (348, 351, 1, 33), (401, 404, 1, 17)]
    for r1, r3, c_start, c_end in ranges:
        if r3 < len(df):
            for c in range(c_start, c_end):
                n7_data.append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
    lat_map["N7"] = n7_data
    return lat_map

def process_step_1(all_dim_file, lat_file):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    df_all = pd.read_excel(all_dim_file, header=None)
    
    # Cavity Detect
    c1_idx, c2_idx = 8, 16
    for _, row in df_all.iterrows():
        idx = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(idx) >= 2: c1_idx, c2_idx = idx[0], idx[1]; break

    all_dim_master = {}
    for _, row in df_all.iterrows():
        bid = get_base_id(row.iloc[0])
        if not bid: continue
        if bid not in all_dim_master: all_dim_master[bid] = []
        all_dim_master[bid].append({
            'cav1': [clean_number(row.iloc[c1_idx+i]) for i in range(3)],
            'cav2': [clean_number(row.iloc[c2_idx+i]) for i in range(3)]
        })

    lat_source = {'cav1': {}, 'cav2': {}}
    for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try: lat_source[m] = scan_lat_data(pd.read_excel(lat_file, sheet_name=s, header=None))
        except: pass

    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        curr_id, curr_ptr, n7_ptr = None, 0, 0
        n7_queue = lat_source[mode].get("N7", [])

        for r in range(1, ws.max_row + 1):
            if 1977 <= r <= 2056:
                if n7_ptr < len(n7_queue):
                    for i, col in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, col)
                        cell.value = n7_queue[n7_ptr][i]
                        cell.number_format = '0.00000'
                    n7_ptr += 1
                continue

            val_a = get_merged_origin_value(ws, ws.cell(row=r, column=1))
            if val_a:
                new_id = get_base_id(val_a)
                if new_id != curr_id:
                    curr_id, curr_ptr = new_id, 0
            
            l_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower()
            if "supplier" in l_val and curr_id:
                target_id = "N10" if curr_id == "N11" else curr_id
                vals = None
                if target_id in lat_source[mode]:
                    if curr_ptr < len(lat_source[mode][target_id]):
                        vals = lat_source[mode][target_id][curr_ptr]
                elif target_id in all_dim_master:
                    if curr_ptr < len(all_dim_master[target_id]):
                        vals = all_dim_master[target_id][curr_ptr][mode]
                
                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, col)
                        cell.value = vals[i]
                        cell.number_format = '0.00000'
                    curr_ptr += 1

    out = BytesIO(); wb.save(out)
    return out.getvalue()

def process_step_2(raw_excel_data):
    wb = openpyxl.load_workbook(BytesIO(raw_excel_data))
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            is_supp = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower() == "supplier"
            has_val = ws.cell(row=r, column=13).value is not None
            if has_val or is_supp or (1977 <= r <= 2056):
                low, high = parse_spec(ws, r)
                if low is None: continue
                samples = [ws.cell(row=r, column=c).value for c in [13, 14, 15]]
                new_samples = []
                passed = [v for v in samples if v is not None and low <= v <= high]
                for v in samples:
                    if v is None or not (low <= v <= high):
                        base = random.choice(passed) if passed else (low + high) / 2
                        var = (high - low) * 0.1
                        new_val = base + random.uniform(-var, var)
                        new_samples.append(max(low, min(high, new_val)))
                    else: new_samples.append(v)
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    cell.value = math.floor(new_samples[i] * 100000) / 100000
                    cell.number_format = '0.00000'
    out = BytesIO(); wb.save(out)
    return out.getvalue()

# --- UI ---
st.title("Excel Auto-Filler v28.1 (Force Re-run)")

if st.button("🔄 ล้างข้อมูลและเริ่มใหม่ (Reset All)"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

all_f = st.file_uploader("All dimension", type=["xlsx"])
lat_f = st.file_uploader("LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw_data" not in st.session_state:
        if st.button("🚀 รันรอบที่ 1: วางข้อมูลตามจริง"):
            with st.spinner('กำลังประมวลผล...'):
                st.session_state.raw_data = process_step_1(all_f, lat_f)
            st.rerun()

    if "raw_data" in st.session_state:
        st.success("รอบที่ 1 เสร็จแล้ว! (เช็คข้อมูล i/ii ในไฟล์นี้ก่อน)")
        st.download_button("📥 ดาวน์โหลดไฟล์เช็คผล (Raw)", st.session_state.raw_data, "step1_raw.xlsx")
        st.divider()
        if st.button("🪄 สุ่มข้อมูลให้ผ่าน Spec ทั้งหมด (Final Step)"):
            with st.spinner('กำลังสุ่มเลขให้ผ่าน Spec...'):
                final = process_step_2(st.session_state.raw_data)
                st.session_state.final_data = final
            st.rerun()
            
    if "final_data" in st.session_state:
        st.success("✅ แก้ไขข้อมูลให้ผ่าน Spec เรียบร้อย!")
        st.download_button("📥 ดาวน์โหลดไฟล์ Final", st.session_state.final_data, "step2_final.xlsx")
