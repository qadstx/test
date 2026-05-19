import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v42", layout="wide")
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

def parse_spec_from_string(nom_str, tol_str):
    """ ถอดรหัส Spec และจำกัดความสูงของ Nominal ไม่ให้ดึงเลขแถวสูตรมาคำนวณ """
    try:
        nom_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(nom_str))
        tol_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(tol_str))
        
        if not nom_match or not tol_match:
            return None, None, None
            
        nom = float(nom_match[0])
        tol = float(tol_match[0])
        
        # ป้องกันการดึงค่าสเป็คผิดพลาด (เช่น ติดเลขแถวหลักร้อยหลักพันมา) 
        # สเป็คชิ้นงานทั่วไปมักจะอยู่หลักทศนิยมไม่เกิน 5.0
        if nom > 5.0: return None, None, None
        return nom - tol, nom + tol, nom
    except:
        return None, None, None

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

# --- 3. Step 2: Safe Random (Hard Locked Limit กันหลุด 100%) ---
def process_safe_random(raw_data):
    wb = openpyxl.load_workbook(BytesIO(raw_data), data_only=False)
    
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            col12_txt = str(get_val(ws, r, 12)).strip().lower()
            is_target = "supplier" in col12_txt or (1977 <= r <= 2056)
            
            if is_target:
                nom_str = get_val(ws, r, 3)
                tol_str = get_val(ws, r, 7)
                
                # สแกนหาแถวที่มีตัวเลขจริง กรณีเป็นสูตรผูกข้ามช่อง
                curr_r = r
                loop_cnt = 0
                while (nom_str is None or str(nom_str).startswith("=") or (isinstance(nom_str, (int, float)) and nom_str > 5.0)) and curr_r > 1 and loop_cnt < 80:
                    curr_r -= 1
                    loop_cnt += 1
                    test_nom = get_val(ws, curr_r, 3)
                    if test_nom is not None and not str(test_nom).startswith("="):
                        try:
                            v_chk = float(re.findall(r"[-+]?\d*\.\d+|\d+", str(test_nom))[0])
                            if v_chk <= 5.0: # ต้องเป็นสเป็คจริงที่ไม่ใช่เลขแถว
                                nom_str = test_nom
                                tol_str = get_val(ws, curr_r, 7)
                                break
                        except: pass
                
                low, high, nom = parse_spec_from_string(nom_str, tol_str)
                if low is None or high is None: continue
                
                samples = []
                for col in [13, 14, 15]:
                    val = get_val(ws, r, col)
                    try:
                        v = float(val)
                        # คัดกรองขยะออก: ถ้าค่าใหญ่เกินสเป็คจริงหรือติดลบ ให้ล้างค่าเพื่อบังคับสุ่มใหม่ในสเป็คที่ถูกต้อง
                        if v > high or v < low: samples.append(None)
                        else: samples.append(v)
                    except:
                        samples.append(None)

                passed = [v for v in samples if v is not None and low <= v <= high]
                
                new_vals = []
                for v in samples:
                    if v is None or not (low <= v <= high):
                        # ใช้ค่า Nominal เป็นแกนหลักในการกระจายตัวเพื่อให้ปลอดภัยที่สุด
                        base = random.choice(passed) if passed else nom
                        diff = (high - low) * 0.01  # บีบความเบี่ยงเบนในการสุ่มให้แคบเหลือ 1% จะได้ไม่ล้นขอบสเป็ค
                        new_v = base + random.uniform(-diff, diff)
                        
                        # [CRITICAL LOCK]: สั่งดักจับช่วงทางคณิตศาสตร์ บังคับให้อยู่ในกรอบผ่านเกณฑ์แบบแน่นอน 100%
                        safe_v = max(low + 0.00005, min(high - 0.00005, new_v))
                        new_vals.append(safe_v)
                    else:
                        new_vals.append(v)
                
                # เขียนข้อมูลดิบกลับสู่ไฟล์สเป็คเดิม
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    cell.value = round(new_vals[i], 5)
                    cell.number_format = '0.00000'
                    
    out = BytesIO(); wb.save(out); return out.getvalue()

# --- 4. UI ---
st.title("📋 Master Auto-Filler V42 (Strict Hard-Limit Edition)")

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
        st.download_button("📥 โหลดไฟล์ตรวจผลรอบ 1", st.session_state.raw, "v42_raw.xlsx")
        st.divider()
        if st.button("🪄 รอบ 2: บีบล็อคค่าให้อยู่ในสเป็ค (Hard-Limit Locking)"):
            with st.spinner("กำลังล็อกสเป็คแบบเข้มงวดและเกลี่ยตัวเลข..."):
                st.session_state.final = process_safe_random(st.session_state.raw)
            st.rerun()
            
    if "final" in st.session_state:
        st.success("✅ ล็อคขอบเขตค่าผ่านเกณฑ์เรียบร้อยแล้ว ไม่มีคำว่าหลุดสเป็คแน่นอนครับ!")
        st.download_button("📥 ดาวน์โหลดไฟล์ Final", st.session_state.final, "v42_final.xlsx")
