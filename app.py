import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v30", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

# --- Common Utils ---
def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABTN])(\d+)', text)
    if match: return f"{match.group(1)}{match.group(2)}"
    return text

def clean_num(v):
    try:
        if v is None or pd.isna(v): return None
        return math.floor(abs(float(v)) * 100000) / 100000
    except: return None

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
        nom = float(re.findall(r"[-+]?\d*\.\d+|\d+", str(get_val(ws, r, 3)))[0])
        tol = float(re.findall(r"[-+]?\d*\.\d+|\d+", str(get_val(ws, r, 7)))[0])
        return nom - tol, nom + tol
    except: return None, None

# --- Main Logic ---
def scan_lat_v30(df):
    data = {"B_QUEUE": [], "N7_QUEUE": [], "N10_VAL": None}
    
    # 1. กวาด B1-B80 ตามพิกัดที่ระบุ (แถวใน pandas คือ Excel-1)
    b_ranges = [
        (29, 31, 1, 33),   # Row 30-32, Col B-AG
        (83, 85, 1, 33),   # Row 84-86, Col B-AG
        (136, 138, 1, 33), # Row 137-139, Col B-AG
        (189, 191, 1, 25)  # Row 190-192, Col B-Y (B80 จบที่นี่)
    ]
    for r_start, r_end, c_start, c_end in b_ranges:
        for c in range(c_start, c_end):
            if r_end < len(df) and c < df.shape[1]:
                data["B_QUEUE"].append([clean_num(df.iloc[r_start+i, c]) for i in range(3)])

    # 2. กวาด N7 ตามพิกัดเดิม
    n7_ranges = [(295, 297, 1, 33), (348, 350, 1, 33), (401, 403, 1, 17)]
    for r_start, r_end, c_start, c_end in n7_ranges:
        for c in range(c_start, c_end):
            if r_end < len(df) and c < df.shape[1]:
                data["N7_QUEUE"].append([clean_num(df.iloc[r_start+i, c]) for i in range(3)])

    # 3. กวาด N10 (สำหรับวาง N11)
    if 403 < len(df):
        # สมมติ N10 อยู่ที่ Column N (index 13) ของแถว 401-403
        data["N10_VAL"] = [clean_num(df.iloc[400+i, 13]) for i in range(3)]
        
    return data

def process_step_1(all_dim_f, lat_f):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    df_all = pd.read_excel(all_dim_f, header=None)
    
    # ดึงข้อมูลจาก All Dimension (สำหรับ B81 ขึ้นไป และ A)
    all_dim_master = {}
    c1_idx, c2_idx = 8, 16 # Default
    for bid, row in df_all.iterrows():
        name = get_base_id(row.iloc[0])
        if not name: continue
        if name not in all_dim_master: all_dim_master[name] = []
        all_dim_master[name].append({
            'cav1': [clean_num(row.iloc[c1_idx+i]) for i in range(3)],
            'cav2': [clean_num(row.iloc[c2_idx+i]) for i in range(3)]
        })

    lat_data = {}
    for mode, sheet in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try: lat_data[mode] = scan_lat_v30(pd.read_excel(lat_f, sheet_name=sheet, header=None))
        except: lat_data[mode] = None

    for sn in ["Cav 1", "Cav 2"]:
        if sn not in wb.sheetnames: continue
        ws = wb[sn]
        m = 'cav1' if sn == "Cav 1" else 'cav2'
        if not lat_data[m]: continue

        b_idx, n7_idx, curr_id, curr_ptr = 0, 0, None, 0
        
        for r in range(1, ws.max_row + 1):
            # --- วาง N7 ---
            if 1977 <= r <= 2056:
                if n7_idx < len(lat_data[m]["N7_QUEUE"]):
                    for i, col in enumerate([13, 14, 15]):
                        get_writable_cell(ws, r, col).value = lat_data[m]["N7_QUEUE"][n7_idx][i]
                    n7_idx += 1
                continue

            # --- วาง B1-B80 และอื่นๆ ---
            id_val = get_base_id(get_val(ws, r, 1))
            if id_val:
                if id_val != curr_id: curr_id, curr_ptr = id_val, 0
            
            l_val = str(get_val(ws, r, 12)).lower()
            if "supplier" in l_val and curr_id:
                vals = None
                num_id = int(re.search(r'\d+', curr_id).group()) if re.search(r'\d+', curr_id) else 0
                
                if curr_id.startswith("B") and 1 <= num_id <= 80:
                    # ใช้ Sequential Queue สำหรับ B1-B80
                    if b_idx < len(lat_data[m]["B_QUEUE"]):
                        vals = lat_data[m]["B_QUEUE"][b_idx]
                        b_idx += 1
                elif curr_id == "N11":
                    vals = lat_data[m]["N10_VAL"]
                else:
                    # B81 ขึ้นไป หรือ A ให้ใช้ All Dimension logic เดิม
                    if curr_id in all_dim_master and curr_ptr < len(all_dim_master[curr_id]):
                        vals = all_dim_master[curr_id][curr_ptr][m]
                        curr_ptr += 1
                
                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        get_writable_cell(ws, r, col).value = vals[i]
                        get_writable_cell(ws, r, col).number_format = '0.00000'

    out = BytesIO(); wb.save(out); return out.getvalue()

def process_step_2(raw_data):
    wb = openpyxl.load_workbook(BytesIO(raw_data))
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            is_supp = "supplier" in str(get_val(ws, r, 12)).lower()
            if is_supp or (1977 <= r <= 2056):
                low, high = parse_spec(ws, r)
                if low is None: continue
                samples = [get_val(ws, r, c) for c in [13, 14, 15]]
                passed = [v for v in samples if v is not None and low <= v <= high]
                new_samples = []
                for v in samples:
                    if v is None or not (low <= v <= high):
                        base = random.choice(passed) if passed else (low + high) / 2
                        new_val = base + random.uniform(-(high-low)*0.05, (high-low)*0.05)
                        new_samples.append(max(low, min(high, new_val)))
                    else: new_samples.append(v)
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    cell.value = math.floor(new_samples[i] * 100000) / 100000
                    cell.number_format = '0.00000'
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- UI ---
st.title("Excel Auto-Filler v30 (Sequential B1-B80)")
if st.button("🔄 Reset"):
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

all_f = st.file_uploader("All dimension", type=["xlsx"])
lat_f = st.file_uploader("LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw" not in st.session_state:
        if st.button("🚀 รอบ 1: วางข้อมูลตามพิกัด"):
            st.session_state.raw = process_step_1(all_f, lat_f)
            st.rerun()
    else:
        st.success("วางข้อมูลเรียบร้อย (B1-B80 ใช้ Sequential Queue)")
        st.download_button("📥 โหลดไฟล์เช็คผล", st.session_state.raw, "check_v30.xlsx")
        if st.button("🪄 สุ่มให้ผ่าน Spec ทั้งหมด"):
            st.session_state.final = process_step_2(st.session_state.raw)
            st.rerun()
    if "final" in st.session_state:
        st.download_button("📥 โหลดไฟล์ Final", st.session_state.final, "final_v30.xlsx")
