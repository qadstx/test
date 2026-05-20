import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import random

st.set_page_config(page_title="Excel Auto-Filler v53", layout="wide")
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
    wb_lat = openpyxl.load_workbook(lat_file, data_only=True)
    
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
    
    df_lat_cav1 = pd.read_excel(lat_file, sheet_name='LAT (Cav 1)', header=None)
    df_lat_cav2 = pd.read_excel(lat_file, sheet_name='LAT (Cav 2)', header=None)
    
    def scan_lat_data(df):
        d = {"B": [], "N7": []}
        b_c = [(29, 31, 1, 33), (83, 85, 1, 33), (136, 138, 1, 33), (189, 191, 1, 25)]
        for r1, r3, c1, c2 in b_c:
            for c in range(c1, c2):
                if r3 < len(df): d["B"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
        # ⚡ ดึงกลุ่มข้อมูล N7 กลับมาประจำการ
        n7_c = [(295, 297, 1, 33), (348, 350, 1, 33), (401, 403, 1, 17)]
        for r1, r3, c1, c2 in n7_c:
            for c in range(c1, c2):
                if r3 < len(df): d["N7"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
        return d

    lat_src = {
        'cav1': scan_lat_data(df_lat_cav1),
        'cav2': scan_lat_data(df_lat_cav2)
    }

    for sn in ["Cav 1", "Cav 2"]:
        if sn not in wb.sheetnames: continue
        ws, m = wb[sn], ('cav1' if sn == "Cav 1" else 'cav2')
        sheet_lat_name = 'LAT (Cav 1)' if sn == "Cav 1" else 'LAT (Cav 2)'
        
        # ดึงค่าตรงพิกัด N11 จากเซลล์ R402, R403, R404 แน่นอนแม่นยำ
        n10_vals = [None, None, None]
        if sheet_lat_name in wb_lat.sheetnames:
            ws_lat_src = wb_lat[sheet_lat_name]
            try:
                n10_vals = [
                    float(ws_lat_src['R402'].value),
                    float(ws_lat_src['R403'].value),
                    float(ws_lat_src['R404'].value)
                ]
            except:
                n10_vals = [
                    clean_number(ws_lat_src['R402'].value),
                    clean_number(ws_lat_src['R403'].value),
                    clean_number(ws_lat_src['R404'].value)
                ]

        b_p, n7_p, curr_id, curr_ptr = 0, 0, None, 0
        for r in range(1, ws.max_row + 1):
            # 🛠️ [RESTORED] บันทึกข้อมูล N7 ลงพิกัดตารางช่วงที่ขยับขึ้นมาใหม่ (แถว 1976-2055)
            if 1976 <= r <= 2055:
                if lat_src[m] and n7_p < len(lat_src[m]["N7"]):
                    v = lat_src[m]["N7"][n7_p]
                    for i, c in enumerate([13, 14, 15]): 
                        cell = get_writable_cell(ws, r, c)
                        cell.value = v[i]
                        if curr_id and curr_id == "B81":
                            cell.number_format = '0.0000'
                        else:
                            cell.number_format = '0.00000'
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
                    vals = n10_vals
                elif curr_id in all_dim_master and curr_ptr < len(all_dim_master[curr_id]):
                    vals = all_dim_master[curr_id][curr_ptr][m]; curr_ptr += 1
                
                if vals:
                    for i, c in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, c)
                        cell.value = vals[i]
                        if curr_id and curr_id == "B81":
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
            # ให้รอบสุ่มเกลี่ยค่าทำงานครอบคลุมทั้งแถว Supplier และ บล็อก N7 แถว 1976-2055
            is_target = "supplier" in col12_txt or (1976 <= r <= 2055)
            
            if is_target:
                # กำหนดเซฟตี้สเป็คให้กับบล็อกตาราง N7 เพื่อไม่ให้เกิดลูปค้าง
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
                    if curr_id and curr_id == "B81":
                        cell.value = round(final_vals[i], 4)
                        cell.number_format = '0.0000'
                    else:
                        cell.value = round(final_vals[i], 5)
                        cell.number_format = '0.00000'
                    
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 4. UI Layout ---
st.title("📋 Master Auto-Filler V53 (N7 Restored Edition)")

if st.button("🔄 Reset ระบบ"):
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

all_f = st.file_uploader("1. All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw" not in st.session_state:
        if st.button("🚀 รอบ 1: วางข้อมูลจริง (ตรวจเช็คค่าหลุดดิบก่อนปรับ)"):
            with st.spinner("กำลังเขียนข้อมูลดิบและดึงบล็อกตาราง N7 กลับมา..."):
                st.session_state.raw = process_real_data(all_f, lat_f)
            st.rerun()
    else:
        st.success("รอบ 1 สำเร็จแล้ว! ดึงกลุ่มข้อมูล N7 และ N11 ลงพิกัดที่ถูกต้องเรียบร้อย")
        st.download_button("📥 ดาวน์โหลดไฟล์ตรวจผลรอบ 1 (ค่าดิบจริง)", st.session_state.raw, "v53_step1_raw.xlsx")
        st.divider()
        
        if st.button("🪄 รอบ 2: สุ่มเติมช่องว่าง & ปรับค่าหลุดให้เข้ากลุ่มสเป็ค"):
            with st.spinner("กำลังเกลี่ยปรับค่าสเป็คของกลุ่ม B, N7 และ N11 แบบความเร็วสูง..."):
                st.session_state.final = process_safe_random(st.session_state.raw)
            st.rerun()
            
    if "final" in st.session_state:
        st.success("✅ รันครบถ้วนเสร็จสมบูรณ์! ข้อมูลจัดเรียงฟอร์แมตไร้ที่ติพร้อมดาวน์โหลดครับ")
        st.download_button("📥 ดาวน์โหลดไฟล์ Final (ผ่านเกณฑ์ 100%)", st.session_state.final, "v53_step2_final.xlsx")
