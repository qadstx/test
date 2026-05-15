import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v35", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

# --- 1. Utilities (ล้างข้อมูล/หาพิกัด) ---
def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABT])-?(\d+)', text)
    if match:
        return f"{match.group(1)}{int(match.group(2))}"
    return text

def clean_number(value):
    try:
        if value is None or pd.isna(value): return None
        return abs(float(value))
    except: return value

def get_writable_cell(ws, r, c):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

def get_val(ws, r, c):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return cell.value

def parse_spec(ws, r):
    try:
        nom_str = str(get_val(ws, r, 3))
        tol_str = str(get_val(ws, r, 7))
        nom = float(re.findall(r"[-+]?\d*\.\d+|\d+", nom_str)[0])
        tol = float(re.findall(r"[-+]?\d*\.\d+|\d+", tol_str)[0])
        return nom - tol, nom + tol
    except: return None, None

# --- 2. การกวาดข้อมูลจาก LAT (B1-80, N7, N10) ---
def scan_lat_positional(df):
    lat_data = {"B_QUEUE": [], "N7_QUEUE": [], "N10_VAL": None}
    
    # กวาด B1-B80 (แนวนอน)
    b_coords = [(29, 31, 1, 33), (83, 85, 1, 33), (136, 138, 1, 33), (189, 191, 1, 25)]
    for r_start, r_end, c_start, c_end in b_coords:
        for c in range(c_start, c_end):
            if r_end < len(df) and c < df.shape[1]:
                samples = [clean_number(df.iloc[r_start+i, c]) for i in range(3)]
                lat_data["B_QUEUE"].append(samples)

    # กวาด N7
    n7_coords = [(295, 297, 1, 33), (348, 350, 1, 33), (401, 403, 1, 17)]
    for r_start, r_end, c_start, c_end in n7_coords:
        for c in range(c_start, c_end):
            if r_end < len(df) and c < df.shape[1]:
                lat_data["N7_QUEUE"].append([clean_number(df.iloc[r_start+i, c]) for i in range(3)])

    # FIX N10: กลับไปใช้ Column B (index 1) ตามเดิม
    if 402 < len(df):
        lat_data["N10_VAL"] = [clean_number(df.iloc[400+i, 1]) for i in range(3)]
        
    return lat_data

# --- 3. Step 1: วางข้อมูลจริง (ยึดโครง v33 ของคุณ) ---
def process_real_data(all_dim_file, lat_file):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    df_all = pd.read_excel(all_dim_file, header=None)

    # เก็บ All Dimension (A, T, B81+)
    all_dim_master = {}
    cav1_idx, cav2_idx = 8, 16
    for _, row in df_all.iterrows():
        indices = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(indices) >= 2:
            cav1_idx, cav2_idx = indices[0], indices[1]
            break

    for _, row in df_all.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if not re.search(r'[ABT]', raw_name): continue
        bid = get_base_id(raw_name)
        if not bid: continue

        if bid not in all_dim_master: all_dim_master[bid] = []
        c1 = [clean_number(row.iloc[cav1_idx+i]) for i in range(3)]
        c2 = [clean_number(row.iloc[cav2_idx+i]) for i in range(3)]
        all_dim_master[bid].append({'cav1': c1, 'cav2': c2})

    lat_sources = {}
    for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try:
            df_lat = pd.read_excel(lat_file, sheet_name=s, header=None)
            lat_sources[m] = scan_lat_positional(df_lat)
        except: lat_sources[m] = None

    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        b_ptr, n7_ptr, current_id, current_ptr = 0, 0, None, 0

        for r in range(1, ws.max_row + 1):
            if 1977 <= r <= 2056:
                if lat_sources[mode] and n7_ptr < len(lat_sources[mode]["N7_QUEUE"]):
                    vals = lat_sources[mode]["N7_QUEUE"][n7_ptr]
                    for i, col in enumerate([13, 14, 15]): get_writable_cell(ws, r, col).value = vals[i]
                    n7_ptr += 1
                continue

            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                new_id = get_base_id(cell_a)
                if new_id != current_id:
                    current_id = new_id
                    current_ptr = 0
            
            l_val = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in l_val and current_id:
                vals = None
                match_num = re.search(r'\d+', current_id)
                num = int(match_num.group()) if match_num else 0

                if current_id.startswith("B") and 1 <= num <= 80:
                    if lat_sources[mode] and b_ptr < len(lat_sources[mode]["B_QUEUE"]):
                        vals = lat_sources[mode]["B_QUEUE"][b_ptr]; b_ptr += 1
                elif current_id == "N11":
                    if lat_sources[mode]: vals = lat_sources[mode]["N10_VAL"]
                else:
                    # ส่วนที่เหลือ (A, T, B81+)
                    if current_id in all_dim_master:
                        data_list = all_dim_master[current_id]
                        if current_ptr < len(data_list):
                            vals = data_list[current_ptr][mode]; current_ptr += 1

                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        target = get_writable_cell(ws, r, col)
                        target.value = vals[i]
                        target.number_format = '0.00000'

    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 4. Step 2: สุ่มค่าให้เข้าสเป็คและเติมช่องว่าง ---
def process_random_spec(raw_excel_data):
    wb = openpyxl.load_workbook(BytesIO(raw_excel_data))
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            is_supp = "supplier" in str(get_val(ws, r, 12)).lower()
            if is_supp or (1977 <= r <= 2056):
                low, high = parse_spec(ws, r)
                if low is None: continue
                
                samples = [get_val(ws, r, c) for c in [13, 14, 15]]
                passed = [v for v in samples if v is not None and low <= v <= high]
                new_vals = []
                for v in samples:
                    if v is None or not (low <= v <= high):
                        base = random.choice(passed) if passed else (low + high) / 2
                        var = (high - low) * 0.05
                        new_v = base + random.uniform(-var, var)
                        new_vals.append(max(low, min(high, new_v)))
                    else:
                        new_vals.append(v)
                
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    cell.value = math.floor(new_vals[i] * 100000) / 100000
                    cell.number_format = '0.00000'
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 5. UI ---
st.title("📋 Master Auto-Filler V35")
if st.button("🔄 Reset"):
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

all_f = st.file_uploader("1. All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw" not in st.session_state:
        if st.button("🚀 รอบ 1: วางข้อมูลจริง (B1-80 พิกัดตรง / B81+ All Dim)"):
            st.session_state.raw = process_real_data(all_f, lat_f)
            st.rerun()
    else:
        st.success("วางข้อมูลรอบ 1 สำเร็จ (เช็ค A, T, B81+ ได้เลย)")
        st.download_button("📥 โหลดไฟล์รอบ 1", st.session_state.raw, "v35_raw.xlsx")
        if st.button("🪄 รอบ 2: สุ่มให้ผ่าน Spec & เติมช่องว่าง"):
            st.session_state.final = process_random_spec(st.session_state.raw)
            st.rerun()
            
    if "final" in st.session_state:
        st.download_button("📥 ดาวน์โหลดไฟล์ Final", st.session_state.final, "v35_final.xlsx")
