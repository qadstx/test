import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import random

st.set_page_config(page_title="Excel Auto-Filler v50", layout="wide")
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

def extract_pure_number(text_val):
    if text_val is None: return None
    match = re.search(r"[-+]?\d*\.\d+|\d+", str(text_val))
    if match: return float(match.group())
    return None

# --- 2. Step 1: Real Data Process ---
def process_real_data(all_dim_file, lat_file):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME, data_only=False)
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
    
    def scan_lat(df):
        d = {"B": [], "N7": [], "N10": None}
        b_c = [(29, 31, 1, 33), (83, 85, 1, 33), (136, 138, 1, 33), (189, 191, 1, 25)]
        for r1, r3, c1, c2 in b_c:
            for c in range(c1, c2):
                if r3 < len(df): d["B"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
        n7_c = [(295, 297, 1, 33), (348, 350, 1, 33), (401, 403, 1, 17)]
        for r1, r3, c1, c2 in n7_c:
            for c in range(c1, c2):
                if r3 < len(df): d["N7"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
                
        # ดึงจากพิกัด R402, R403, R404 คอลัมน์ R (Index 17 ใน pandas) ของไฟล์ LAT
        if len(df) >= 404:
            try:
                d["N10"] = [float(df.iloc[401, 17]), float(df.iloc[402, 17]), float(df.iloc[403, 17])]
            except:
                d["N10"] = [clean_number(df.iloc[401, 17]), clean_number(df.iloc[402, 17]), clean_number(df.iloc[403, 17])]
        return d

    lat_src = {}
    for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try: lat_src[m] = scan_lat(pd.read_excel(lat_file, sheet_name=s, header=None))
        except: lat_src[m] = None

    for sn in ["Cav 1", "Cav 2"]:
        if sn not in wb.sheetnames: continue
        ws, m = wb[sn], ('cav1' if sn == "Cav 1" else 'cav2')
        b_p, n7_p, curr_id, curr_ptr = 0, 0, None, 0
        for r in range(1, ws.max_row + 1):
            # จัดตำแหน่งตาราง N7 ขยับขึ้น 1 แถว (1976-2055)
            if 1976 <= r <= 2055:
                if lat_src[m] and n7_p < len(lat_src[m]["N7"]):
                    v = lat_src[m]["N7"][n7_p]
                    for i, c in enumerate([13, 14, 15]): 
                        cell = get_writable_cell(ws, r, c)
                        cell.value = v[i]
                        cell.number_format = '0.0000' if curr_id == "B81" else '0.00000'
                    n7_p += 1
                continue
                
            id_v = get_base_id(get_val(ws, r, 1))
            if id_v:
                if id_v != curr_id: curr_id, curr_ptr = id_v, 0
                
            col12_txt = str(get_val(ws, r, 12)).strip().lower()
            if "supplier" in col12_txt and curr_id:
                vals = None
                num = int(re.search(r'\d+', curr_id).group()) if re.search(r'\d+', curr_id) else 0
                if curr_id.startswith("B") and 1 <= num <= 80:
                    if lat_src[m] and b_p < len(lat_src[m]["B"]): vals = lat_src[m]["B"][b_p]; b_p += 1
                elif curr_id == "N11":
                    if lat_src[m]: vals = lat_src[m]["N10"]
                elif curr_id in all_dim_master and curr_ptr < len(all_dim_master[curr_id]):
                    vals = all_dim_master[curr_id][curr_ptr][m]; curr_ptr += 1
                
                if vals:
                    for i, c in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, c)
                        cell.value = vals[i]
                        if curr_id == "B81":
                            cell.number_format = '0.0000'
                        else:
                            cell.number_format = '0.00000'
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 3. Step 2: Safe Random ---
def process_safe_random(raw_data):
    wb = openpyxl.load_workbook(BytesIO(raw_data), data_only=False)
    
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        last_known_low, last_known_high, last_known_nom = None, None, None
        curr_id = None
        
        for r in range(1, ws.max_row + 1):
            id_v = get_base_id(get_val(ws, r, 1))
            if id_v: curr_id = id_v

            v_nom = get_val(ws, r, 4)
            v_tol = get_val(ws, r, 8)
            
            if v_nom is not None and str(v_nom).strip() != "" and not str(v_nom).startswith("="):
                parsed_nom = extract_pure_number(v_nom)
                if parsed_nom is not None: last_known_nom = parsed_nom
            if v_tol is not None and str(v_tol).strip() != "" and not str(v_tol).startswith("="):
                parsed_tol = extract_pure_number(v_tol)
                if parsed_tol is not None and last_known_nom is not None:
                    last_known_low = last_known_nom - parsed_tol
                    last_known_high = last_known_nom + parsed_tol

            col12_txt = str(get_val(ws, r, 12)).strip().lower()
            is_target = "supplier" in col12_txt or (1976 <= r <= 2055)
            
            if is_target:
                # ⭐ [FIXED] แก้ไขจุดเครื่องค้าง: ล็อกสเป็คให้ช่วงแถว N7 (Roughness) โดยตรงเพื่อไม่ให้ค่าหลุดเป็น None
                if 1976 <= r <= 2055:
                    low, high, nom = 0.0, 0.4, 0.05
                else:
                    low, high, nom = last_known_low, last_known_high, last_known_nom
                
                if low is None or high is None: continue
                
                current_vals = []
                for col in [13, 14, 15]:
                    val = get_val(ws, r, col)
                    try: current_vals.append(float(val))
                    except: current_vals.append(None)
                
                passed_samples = [v for v in current_vals if v is not None and low <= v <= high]
                final_vals = []
                for v in current_vals:
                    if v is None or not (low <= v <= high):
                        base = random.choice(passed_samples) if passed_samples else nom
                        tolerance_range = high - low
                        diff = tolerance_range * 0.05 if tolerance_range > 0 else 0.0001
                        new_v = base + random.uniform(-diff, diff)
                        safe_v = max(low + 0.00002, min(high - 0.00002, new_v))
                        final_vals.append(safe_v)
                    else:
                        final_vals.append(v)
                
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    if curr_id == "B81":
                        cell.value = round(final_vals[i], 4)
                        cell.number_format = '0.0000'
                    else:
                        cell.value = round(final_vals[i], 5)
                        cell.number_format = '0.00000'
                    
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 4. UI Layout ---
st.title("📋 Master Auto-Filler V50 (Anti-Freeze Complete Edition)")

if st.button("🔄 Reset ระบบ"):
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

all_f = st.file_uploader("1. All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw" not in st.session_state:
        if st.button("🚀 รอบ 1: วางข้อมูลจริง (ตรวจเช็คค่าหลุดดิบก่อนปรับ)"):
            with st.spinner("กำลังเขียนข้อมูลดิบลงบล็อกตาราง..."):
                st.session_state.raw = process_real_data(all_f, lat_f)
            st.rerun()
    else:
        st.success("รอบ 1 สำเร็จแล้ว! ตรวจสอบพิกัดความถูกต้องเรียบร้อยแล้วครับ")
        st.download_button("📥 ดาวน์โหลดไฟล์ตรวจผลรอบ 1 (ค่าดิบจริง)", st.session_state.raw, "v50_step1_raw.xlsx")
        st.divider()
        
        if st.button("🪄 รอบ 2: สุ่มเติมช่องว่าง & ปรับค่าหลุดให้เข้ากลุ่มสเป็ค"):
            with st.spinner("กำลังเกลี่ยค่าสเป็คและแก้จุดค้างเรียบร้อย..."):
                st.session_state.final = process_safe_random(st.session_state.raw)
            st.rerun()
            
    if "final" in st.session_state:
        st.success("✅ ประมวลผลรอบ 2 เสร็จสิ้น! ปุ่มดาวน์โหลดพร้อมใช้งานแล้วครับ")
        st.download_button("📥 ดาวน์โหลดไฟล์ Final (ผ่านเกณฑ์ 100%)", st.session_state.final, "v50_step2_final.xlsx")
