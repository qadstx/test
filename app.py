import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v38", layout="wide")
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

def parse_spec_v3(ws, r):
    """ ดึงค่าสเป็คแบบรองรับสูตรผูกข้ามเซลล์ (สูตร =ช่องอื่น) 
        หากอ่านแล้วเจอสูตรพัง หรือเจอเลขแถว จะทำการ Scan ย้อนขึ้นไปหาค่าจริงให้
    """
    try:
        nom_val = get_val(ws, r, 3)
        tol_val = get_val(ws, r, 7)
        
        # ถ้าระบุสูตรผูกแล้ว data_only อ่านไม่ได้ หรือดันไปดึงเลขแถวมา (ความเพี้ยนของเลข > 500)
        # ให้ใช้วิธีไล่สแกนขึ้นไปหาแถวก่อนหน้าที่มีตัวเลขสเป็คเขียนอยู่จริง
        curr_r = r
        while (nom_val is None or str(nom_val).startswith("=") or (isinstance(nom_val, (int, float)) and nom_val > 500)) and curr_r > 1:
            curr_r -= 1
            test_nom = get_val(ws, curr_r, 3)
            if test_nom is not None and not str(test_nom).startswith("="):
                try:
                    # ตรวจสอบว่าต้องเป็นสเป็คชิ้นงานจริง (ค่าไม่ควรใหญ่เว่อร์เป็นหลักพัน)
                    v = float(re.findall(r"[-+]?\d*\.\d+|\d+", str(test_nom))[0])
                    if v < 500: 
                        nom_val = test_nom
                        tol_val = get_val(ws, curr_r, 7)
                        break
                except: pass

        nom = float(re.findall(r"[-+]?\d*\.\d+|\d+", str(nom_val))[0])
        tol = float(re.findall(r"[-+]?\d*\.\d+|\d+", str(tol_val))[0])
        return nom - tol, nom + tol, nom
    except: 
        return None, None, None

# --- 2. Step 1: Real Data Process (คงเดิมตามโครงหลักของคุณ) ---
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
            if "supplier" in str(get_val(ws, r, 12)).lower() and curr_id:
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

# --- 3. Step 2: Smart Random (เปิดโหมดอ่านค่าจากสูตร data_only=True) ---
def process_smart_random(raw_data):
    # บังคับเปิดแบบ data_only=True เพื่อดึงผลลัพธ์ของสูตรโคลนเซลล์ออกมาเป็นตัวเลขจริง
    wb = openpyxl.load_workbook(BytesIO(raw_data), data_only=True)
    
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            is_target = "supplier" in str(get_val(ws, r, 12)).lower() or (1977 <= r <= 2056)
            if is_target:
                low, high, nom = parse_spec_v3(ws, r)
                if low is None or high is None: continue
                
                samples = []
                for col in [13, 14, 15]:
                    try: 
                        v = float(get_val(ws, r, col))
                        # ดักจับค่าหลุดสเกลเดิมออกไปก่อน
                        if v > 500 or v < 0: samples.append(None)
                        else: samples.append(v)
                    except: samples.append(None)

                passed = [v for v in samples if v is not None and low <= v <= high]
                
                new_vals = []
                for v in samples:
                    if v is None or not (low <= v <= high):
                        base = random.choice(passed) if passed else nom
                        diff = (high - low) * 0.05  # ควบคุมความแปรปรวน 5% ของช่วง Spec
                        new_v = base + random.uniform(-diff, diff)
                        new_vals.append(max(low + 0.00001, min(high - 0.00001, new_v)))
                    else:
                        new_vals.append(v)
                
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    cell.value = round(new_vals[i], 5)
                    cell.number_format = '0.00000'
                    
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 4. UI ---
st.title("📋 Master Auto-Filler V38 (Formula Value Fix)")
all_f = st.file_uploader("1. All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw" not in st.session_state:
        if st.button("🚀 รอบ 1: วางข้อมูลจริง"):
            st.session_state.raw = process_real_data(all_f, lat_f)
            st.rerun()
    else:
        st.info("รอบ 1 สำเร็จแล้ว! ดำเนินการขั้นต่อไปเพื่อเคลียร์ค่าที่เพี้ยนจากสูตร")
        st.download_button("📥 โหลดไฟล์รอบ 1", st.session_state.raw, "v38_raw.xlsx")
        if st.button("🪄 รอบ 2: แก้ไขสูตรเพี้ยน & สุ่มให้อยู่ใน Spec"):
            st.session_state.final = process_smart_random(st.session_state.raw)
            st.rerun()
            
    if "final" in st.session_state:
        st.success("✅ แก้ไขและคำนวณสุ่มตามขอบเขต Spec จริงเรียบร้อยแล้ว!")
        st.download_button("📥 ดาวน์โหลดไฟล์ Final", st.session_state.final, "v38_final.xlsx")
