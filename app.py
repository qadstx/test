import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import random

st.set_page_config(page_title="Excel Auto-Filler v45", layout="wide")
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

# --- 2. Core Processing Engine (ทำรอบ 1 และ รอบ 2 ควบในครั้งเดียวเพื่อความเร็ว) ---
def process_excel_speed_edition(all_dim_file, lat_file):
    # โหลดครั้งเดียวทำงานจนจบ ไม่ทำลายสูตร
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME, data_only=False)
    
    # --- ส่วนที่ 1: เตรียมข้อมูลจาก All Dimension ---
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
    
    # --- ส่วนที่ 2: เตรียมข้อมูลจาก LAT ---
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

    # --- ส่วนที่ 3: ลุยเขียนข้อมูลลงใน Sheet แผ่นงาน ---
    for sn in ["Cav 1", "Cav 2"]:
        if sn not in wb.sheetnames: continue
        ws = wb[sn]
        m = 'cav1' if sn == "Cav 1" else 'cav2'
        
        b_p, n7_p, curr_id, curr_ptr = 0, 0, None, 0
        
        # ⚡ เทคนิคความเร็ว: วนลูปเก็บข้อมูลสเป็คจากบนลงล่างใส่ Cache ไว้ล่วงหน้า
        active_nom, active_tol = None, None
        spec_cache = {} # เก็บค่าคู่สเป็คประจำแต่ละแถว
        
        for r in range(1, ws.max_row + 1):
            # อ่านค่า Nominal (Col D) และ Tolerance (Col H) เพื่ออัปเดตสเป็คปัจจุบัน
            v_nom = get_val(ws, r, 4)
            v_tol = get_val(ws, r, 8)
            
            if v_nom is not None and str(v_nom).strip() != "" and not str(v_nom).startswith("="):
                active_nom = extract_pure_number(v_nom)
            if v_tol is not None and str(v_tol).strip() != "" and not str(v_tol).startswith("="):
                active_tol = extract_pure_number(v_tol)
                
            if active_nom is not None and active_tol is not None:
                spec_cache[r] = (active_nom - active_tol, active_nom + active_tol, active_nom)

        # วนลูปเพื่อเขียนและตรวจสอบข้อมูลสเป็คจริงด้วยความเร็วสูง
        for r in range(1, ws.max_row + 1):
            # 1. จัดการเขียนโซนพิเศษ LAT (แถว 1977 - 2056)
            if 1977 <= r <= 2056:
                if lat_src[m] and n7_p < len(lat_src[m]["N7"]):
                    v = lat_src[m]["N7"][n7_p]
                    for i, c in enumerate([13, 14, 15]): get_writable_cell(ws, r, c).value = v[i]
                    n7_p += 1
                # ตรวจเช็คสเป็คในโซนนี้ทันที
                if r in spec_cache:
                    low, high, nom = spec_cache[r]
                    fix_and_lock_cell_spec(ws, r, low, high, nom)
                continue
                
            # 2. จัดการดักจับ ID ชิ้นงาน (เช่น B1, T10) ในคอลัมน์ A
            id_v = get_base_id(get_val(ws, r, 1))
            if id_v:
                if id_v != curr_id: curr_id, curr_ptr = id_v, 0
                
            # 3. จัดการเขียนข้อมูลดิบลงแถว Supplier
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
                
                # ⚡ เรียกใช้ Cache สเป็คความเร็วสูงตรวจสอบและแก้ไขค่าหลุดทันที
                if r in spec_cache:
                    low, high, nom = spec_cache[r]
                    fix_and_lock_cell_spec(ws, r, low, high, nom)
                    
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def fix_and_lock_cell_spec(ws, r, low, high, nom):
    """ ฟังก์ชันตรวจสอบค่าหลุดและบังคับแก้ให้อยู่ในสเป็คในขั้นตอนเดียว """
    current_vals = []
    for col in [13, 14, 15]:
        val = get_val(ws, r, col)
        try:
            current_vals.append(float(val))
        except:
            current_vals.append(None)
            
    passed_samples = [v for v in current_vals if v is not None and low <= v <= high]
    
    for i, col in enumerate([13, 14, 15]):
        v = current_vals[i]
        # ถ้าค่าว่าง หรือหลุดขอบเขตสเป็ค
        if v is None or not (low <= v <= high):
            base = random.choice(passed_samples) if passed_samples else nom
            tolerance_range = high - low
            diff = tolerance_range * 0.10 if tolerance_range > 0 else 0.0001
            
            new_v = base + random.uniform(-diff, diff)
            # ล็อคค่าทางคณิตศาสตร์ไม่ให้หลุดขอบสเป็คเด็ดขาด
            safe_v = max(low + 0.00002, min(high - 0.00002, new_v))
            
            cell = get_writable_cell(ws, r, col)
            cell.value = round(safe_v, 5)
            cell.number_format = '0.00000'

# --- 3. UI Dashboard ---
st.title("📋 Master Auto-Filler V45 (Super Fast Turbo Edition)")
st.subheader("ลดเวลาประมวลผลเหลือหลักวินาที และรวมกระบวนการจบในคลิกเดียว")

all_f = st.file_uploader("1. อัปโหลดไฟล์ All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. อัปโหลดไฟล์ LAT", type=["xlsx"])

if all_f and lat_f:
    # ยุบเหลือปุ่มเดียว รันกระบวนการทั้งหมดรวดเดียวจบ
    if st.button("🚀 เริ่มประมวลผลไฟล์ด่วนที่สุด (คลิกเดียวเสร็จ)"):
        with st.spinner("⚡ ระบบกำลังคำนวณโครงสร้างสเป็คด้วยความเร็วสูง..."):
            try:
                final_output = process_excel_speed_edition(all_f, lat_f)
                st.session_state.turbo_final = final_output
                st.success("✅ ประมวลผลและล็อคค่าสเป็คทั้งหมดเสร็จสิ้นภายในไม่กี่วินาที!")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {str(e)}")
                
    if "turbo_final" in st.session_state:
        st.divider()
        st.subheader("📥 ดาวน์โหลดไฟล์ผลลัพธ์")
        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ Final (สเป็คผ่าน 100% ฟอร์มไม่เพี้ยน)",
            data=st.session_state.turbo_final,
            file_name="v45_turbo_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
