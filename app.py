import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v36", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

# --- 1. Utilities ---
def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABT])-?(\d+)', text)
    if match: return f"{match.group(1)}{int(match.group(2))}"
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
        # ดึงเฉพาะตัวเลขออกมา
        nom = float(re.findall(r"[-+]?\d*\.\d+|\d+", nom_str)[0])
        tol = float(re.findall(r"[-+]?\d*\.\d+|\d+", tol_str)[0])
        return nom - tol, nom + tol
    except: return None, None

# --- 2. Scanners ---
def scan_lat_positional(df):
    lat_data = {"B_QUEUE": [], "N7_QUEUE": [], "N10_VAL": None}
    b_coords = [(29, 31, 1, 33), (83, 85, 1, 33), (136, 138, 1, 33), (189, 191, 1, 25)]
    for r1, r3, c1, c2 in b_coords:
        for c in range(c1, c2):
            if r3 < len(df) and c < df.shape[1]:
                lat_data["B_QUEUE"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
    n7_coords = [(295, 297, 1, 33), (348, 350, 1, 33), (401, 403, 1, 17)]
    for r1, r3, c1, c2 in n7_coords:
        for c in range(c1, c2):
            if r3 < len(df) and c < df.shape[1]:
                lat_data["N7_QUEUE"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
    if 402 < len(df):
        lat_data["N10_VAL"] = [clean_number(df.iloc[400+i, 1]) for i in range(3)]
    return lat_data

# --- 3. Step 1: Real Data ---
def process_real_data(all_dim_file, lat_file):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    df_all = pd.read_excel(all_dim_file, header=None)
    all_dim_master = {}
    c1_idx, c2_idx = 8, 16
    for _, row in df_all.iterrows():
        idx = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(idx) >= 2: c1_idx, c2_idx = idx[0], idx[1]; break
    for _, row in df_all.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if not re.search(r'[ABT]', raw_name): continue
        bid = get_base_id(raw_name)
        if bid:
            if bid not in all_dim_master: all_dim_master[bid] = []
            all_dim_master[bid].append({
                'cav1': [clean_number(row.iloc[c1_idx+i]) for i in range(3)],
                'cav2': [clean_number(row.iloc[c2_idx+i]) for i in range(3)]
            })
    lat_src = {}
    for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try: lat_src[m] = scan_lat_positional(pd.read_excel(lat_file, sheet_name=s, header=None))
        except: lat_src[m] = None

    for sn in ["Cav 1", "Cav 2"]:
        if sn not in wb.sheetnames: continue
        ws = wb[sn]
        m = 'cav1' if sn == "Cav 1" else 'cav2'
        b_ptr, n7_ptr, current_id, current_ptr = 0, 0, None, 0
        for r in range(1, ws.max_row + 1):
            if 1977 <= r <= 2056:
                if lat_src[m] and n7_ptr < len(lat_src[m]["N7_QUEUE"]):
                    vals = lat_src[m]["N7_QUEUE"][n7_ptr]
                    for i, col in enumerate([13, 14, 15]): get_writable_cell(ws, r, col).value = vals[i]
                    n7_ptr += 1
                continue
            id_val = get_base_id(get_val(ws, r, 1))
            if id_val:
                if id_val != current_id: current_id, current_ptr = id_val, 0
            l_val = str(get_val(ws, r, 12)).lower()
            if "supplier" in l_val and current_id:
                vals = None
                num = int(re.search(r'\d+', current_id).group()) if re.search(r'\d+', current_id) else 0
                if current_id.startswith("B") and 1 <= num <= 80:
                    if lat_src[m] and b_ptr < len(lat_src[m]["B_QUEUE"]):
                        vals = lat_src[m]["B_QUEUE"][b_ptr]; b_ptr += 1
                elif current_id == "N11":
                    if lat_src[m]: vals = lat_src[m]["N10_VAL"]
                elif current_id in all_dim_master and current_ptr < len(all_dim_master[current_id]):
                    vals = all_dim_master[current_id][current_ptr][m]; current_ptr += 1
                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, col)
                        cell.value = vals[i]
                        cell.number_format = '0.00000'
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 4. Step 2: Random & Pass Spec (Fixed) ---
def process_random_spec(raw_excel_data):
    wb = openpyxl.load_workbook(BytesIO(raw_excel_data))
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            l_val = str(get_val(ws, r, 12)).lower()
            if "supplier" in l_val or (1977 <= r <= 2056):
                low, high = parse_spec(ws, r)
                if low is None or high is None: continue
                
                samples = []
                for col in [13, 14, 15]:
                    v = get_val(ws, r, col)
                    # ตรวจสอบว่าเป็นตัวเลขจริงไหม
                    try: samples.append(float(v))
                    except: samples.append(None)

                # กรองเฉพาะค่าที่เป็นตัวเลขและอยู่ใน Spec
                passed = [v for v in samples if v is not None and isinstance(v, (int, float)) and low <= v <= high]
                
                new_vals = []
                for v in samples:
                    # ถ้าไม่ใช่ตัวเลข หรืออยู่นอก Spec ให้สุ่มใหม่
                    if v is None or not (isinstance(v, (int, float)) and low <= v <= high):
                        base = random.choice(passed) if passed else (low + high) / 2
                        var = (high - low) * 0.03 # สุ่มเบี่ยงเบน 3%
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
st.title("📋 Master Auto-Filler V36 (Fix Random Error)")
if st.button("🔄 Reset"):
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

all_f = st.file_uploader("1. All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw" not in st.session_state:
        if st.button("🚀 รอบ 1: วางข้อมูลจริง"):
            st.session_state.raw = process_real_data(all_f, lat_f)
            st.rerun()
    else:
        st.success("รอบ 1 สำเร็จ!")
        st.download_button("📥 โหลดไฟล์รอบ 1", st.session_state.raw, "v36_raw.xlsx")
        if st.button("🪄 รอบ 2: สุ่มให้ผ่าน Spec"):
            try:
                st.session_state.final = process_random_spec(st.session_state.raw)
                st.rerun()
            except Exception as e:
                st.error(f"Random Error: {e}")
    if "final" in st.session_state:
        st.download_button("📥 ดาวน์โหลดไฟล์ Final", st.session_state.final, "v36_final.xlsx")
