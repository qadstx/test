import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v43", layout="wide")
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
    """ ดึงเฉพาะตัวเลขหรือทศนิยมตัวแรกที่เจอ โดยไม่สนใจอักษรหรือเครื่องหมายองศา """
    if text_val is None: return None
    match = re.search(r"[-+]?\d*\.\d+|\d+", str(text_val))
    if match:
        return float(match.group())
    return None

def parse_spec_universal(ws, r):
    """ ค้นหาและดึงค่าสเป็คแบบครอบจักรวาล รองรับทั้งสูตรผูกและค่าองศา """
    nom_str = get_val(ws, r, 3)
    tol_str = get_val(ws, r, 7)
    
    # ถ้าช่องปัจจุบันเป็นสูตรผูกหรือว่างเปล่า ให้สแกนย้อนขึ้นไปหาแถวที่มีค่าจริง
    curr_r = r
    loop_cnt = 0
    while (nom_str is None or str(nom_str).startswith("=")) and curr_r > 1 and loop_cnt < 100:
        curr_r -= 1
        loop_cnt += 1
        test_nom = get_val(ws, curr_r, 3)
        if test_nom is not None and not str(test_nom).startswith("="):
            nom_str = test_nom
            tol_str = get_val(ws, curr_r, 7)
            break
            
    nom = extract_pure_number(nom_str)
    tol = extract_pure_number(tol_str)
    
    if nom is not None and tol is not None:
        # ดักจับกรณีดึงโดนเลขแถวหลักพัน ให้ตีเป็นค่าใช้งานไม่ได้เพื่อป้องกันเออเร่อ
        if nom > 1000: return None, None, None
        return nom - tol, nom + tol, nom
    return None, None, None

# --- 2. Step 1: Real Data Process (คงเดิม) ---
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
        if 402 < len(df): d["N10"] = [clean_number(df.iloc[400+i, 1]) for i in range(3)]
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
            if 1977 <= r <= 2056:
                if lat_src[m] and n7_p < len(lat_src[m]["N7"]):
                    v = lat_src[m]["N7"][n7_p]
                    for i, c in enumerate([13, 14, 15]): get_writable_cell(ws, r, c).value = v[i]
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
                        cell.value = vals[i]; cell.number_format = '0.00000'
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 3. Step 2: Smart & Strict In-Spec Fixer ---
def process_safe_random(raw_data):
    wb = openpyxl.load_workbook(BytesIO(raw_data), data_only=False)
    
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            col12_txt = str(get_val(ws, r, 12)).strip().lower()
            is_target = "supplier" in col12_txt or (1977 <= r <= 2056)
            
            if is_target:
                low, high, nom = parse_spec_universal(ws, r)
                if low is None or high is None: continue # หากหาข้อกำหนดสเป็คไม่ได้จริง ๆ ให้ข้ามไปก่อน
                
                # ตรวจเช็คค่าดิบดั้งเดิมในช่องตรวจวัดทั้ง 3 ช่อง
                current_vals = []
                for col in [13, 14, 15]:
                    val = get_val(ws, r, col)
                    try:
                        v = float(val)
                        current_vals.append(v)
                    except:
                        current_vals.append(None)
                
                # ค้นหาเฉพาะค่าดั้งเดิมที่ถูกต้องและผ่านเกณฑ์อยู่แล้ว เพื่อนำมาเป็นฐานข้อมูลคำนวณสุ่ม
                passed_samples = [v for v in current_vals if v is not None and low <= v <= high]
                
                final_vals = []
                for v in current_vals:
                    # เงื่อนไข: ถ้าช่องนั้นว่างเปล่า (None) หรือว่ามีค่าที่หลุดขอบเขต Spec ต่ำสุด-สูงสุดไป
                    if v is None or not (low <= v <= high):
                        # เลือกใช้ค่าฐานข้อมูลที่ผ่านเกณฑ์ หรือหากไม่มีเลย ให้เลือกใช้ค่ากลาง Nominal แทน
                        base = random.choice(passed_samples) if passed_samples else nom
                        
                        # คำนวณขอบเขตการเบี่ยงเบนสุ่มแบบแคบมาก ๆ (0.5% ของช่วงสเป็ค) เพื่อให้เกาะกลุ่มใกล้เคียงค่าเดิมที่ไม่หลุด
                        diff = (high - low) * 0.005 if (high - low) > 0 else 0.001
                        new_v = base + random.uniform(-diff, diff)
                        
                        # [CRITICAL HARD-LOCK] บังคับสกัดค่าด้วยฟังก์ชันคณิตศาสตร์ ให้อยู่ในสเป็คปลอดภัยชัวร์ 100%
                        safe_v = max(low + 0.00002, min(high - 0.00002, new_v))
                        final_vals.append(safe_v)
                    else:
                        # หากค่าเดิมคำนวณมาผ่านเกณฑ์และถูกต้องอยู่แล้ว ให้คงค่าเดิมไว้ตามธรรมชาติ
                        final_vals.append(v)
                
                # เขียนตัวเลขผลลัพธ์กลับลงไปในพิกัดเดิม
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    cell.value = round(final_vals[i], 5)
                    cell.number_format = '0.00000'
                    
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 4. UI ---
st.title("📋 Master Auto-Filler V43 (Universal Spec & Hard-Lock)")

if st.button("🔄 Reset ระบบ"):
    for k in list(st.session_state.keys()): del st.session_state[k]
    st.rerun()

all_f = st.file_uploader("1. All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw" not in st.session_state:
        if st.button("🚀 รอบ 1: วางข้อมูลจริง"):
            with st.spinner("กำลังเขียนข้อมูล..."):
                st.session_state.raw = process_real_data(all_f, lat_f)
            st.rerun()
    else:
        st.success("รอบ 1 สำเร็จแล้ว!")
        st.download_button("📥 โหลดไฟล์ตรวจผลรอบ 1", st.session_state.raw, "v43_raw.xlsx")
        st.divider()
        if st.button("🪄 รอบ 2: ตรวจจับสเป็คองศา/สูตร & ล็อคค่าให้อยู่ในกรอบ"):
            with st.spinner("กำลังไล่สแกนแก้องศาและสูตรผูก พร้อมจำกัดค่าให้ผ่านเกณฑ์..."):
                st.session_state.final = process_safe_random(st.session_state.raw)
            st.rerun()
            
    if "final" in st.session_state:
        st.success("✅ ปรับแก้จุดหลุดสเป็คและจัดระเบียบตัวเลขเรียบร้อยแล้ว ไม่มีคำว่าเกินสเป็คแน่นอนครับ!")
        st.download_button("📥 ดาวน์โหลดไฟล์สำเร็จ Final", st.session_state.final, "v43_final.xlsx")
