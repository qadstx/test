import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math

st.set_page_config(page_title="Excel Auto-Filler v33", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

# --- 1. การทำความสะอาดข้อมูล (เหมือน v14) ---
def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    # รองรับทั้ง B-41, B41, A-1, T-01
    match = re.search(r'([ABT])-?(\d+)', text)
    if match:
        return f"{match.group(1)}{int(match.group(2))}" # ใช้ int เพื่อล้าง 0 ข้างหน้า เช่น T-01 -> T1
    return text

def clean_number(value):
    try:
        if value is None or pd.isna(value): return None
        return abs(float(value)) # ต้องเป็นค่าบวกเท่านั้นตาม v14
    except:
        return value

def get_writable_cell(ws, r, c):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

# --- 2. ฟังก์ชันกวาดข้อมูลจาก LAT (เน้นพิกัด) ---
def scan_lat_positional(df):
    lat_data = {"B_QUEUE": [], "N7_QUEUE": [], "N10_VAL": None}
    
    # กวาด B1-B80 (Row 30, 84, 137, 190 ใน Excel)
    b_coords = [(29, 31, 1, 33), (83, 85, 1, 33), (136, 138, 1, 33), (189, 191, 1, 25)]
    for r_start, r_end, c_start, c_end in b_coords:
        for c in range(c_start, c_end):
            if r_end < len(df) and c < df.shape[1]:
                samples = [clean_number(df.iloc[r_start+i, c]) for i in range(3)]
                lat_data["B_QUEUE"].append(samples)

    # กวาด N7 (พิกัดเฉพาะ)
    n7_coords = [(295, 297, 1, 33), (348, 350, 1, 33), (401, 403, 1, 17)]
    for r_start, r_end, c_start, c_end in n7_coords:
        for c in range(c_start, c_end):
            if r_end < len(df) and c < df.shape[1]:
                lat_data["N7_QUEUE"].append([clean_number(df.iloc[r_start+i, c]) for i in range(3)])

    # กวาด N10 (Row 401-403 Col N)
    if 402 < len(df):
        lat_data["N10_VAL"] = [clean_number(df.iloc[400+i, 13]) for i in range(3)]
        
    return lat_data

# --- 3. ฟังก์ชันหลัก (Process) ---
def process_data(all_dim_file, lat_file):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    df_all = pd.read_excel(all_dim_file, header=None)

    # --- ส่วนที่ 1: เก็บข้อมูลจาก All Dimension (Logic v14 เป๊ะๆ) ---
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

    # --- ส่วนที่ 2: เตรียมข้อมูลจาก LAT ---
    lat_sources = {}
    for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try:
            df_lat = pd.read_excel(lat_f, sheet_name=s, header=None)
            lat_sources[m] = scan_lat_positional(df_lat)
        except: lat_sources[m] = None

    # --- ส่วนที่ 3: หยอดข้อมูลลง Template ---
    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        b_ptr = 0
        n7_ptr = 0
        current_id = None
        current_ptr = 0

        for r in range(1, ws.max_row + 1):
            # A. กรณี N7 (พิกัดแถว 1977-2056)
            if 1977 <= r <= 2056:
                if lat_sources[mode] and n7_ptr < len(lat_sources[mode]["N7_QUEUE"]):
                    vals = lat_sources[mode]["N7_QUEUE"][n7_ptr]
                    for i, col in enumerate([13, 14, 15]):
                        get_writable_cell(ws, r, col).value = vals[i]
                    n7_ptr += 1
                continue

            # B. กรณีปกติ (เช็ค ID จากคอลัมน์ A)
            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                new_id = get_base_id(cell_a)
                if new_id != current_id:
                    current_id = new_id
                    current_ptr = 0
            
            # เช็คคำว่า Supplier ในคอลัมน์ L (12)
            l_val = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in l_val and current_id:
                vals = None
                
                # แยกเงื่อนไขการหยิบข้อมูล
                match_num = re.search(r'\d+', current_id)
                num = int(match_num.group()) if match_num else 0

                # 1. ถ้าเป็น B1 - B80 ให้หยิบจาก Queue ของ LAT
                if current_id.startswith("B") and 1 <= num <= 80:
                    if lat_sources[mode] and b_ptr < len(lat_sources[mode]["B_QUEUE"]):
                        vals = lat_sources[mode]["B_QUEUE"][b_ptr]
                        b_ptr += 1
                
                # 2. ถ้าเป็น N11 ให้หยิบ N10 จาก LAT
                elif current_id == "N11":
                    if lat_sources[mode]:
                        vals = lat_sources[mode]["N10_VAL"]

                # 3. ที่เหลือทั้งหมด (A, T, B81+) ให้หยิบจาก All Dimension (Logic v14)
                else:
                    if current_id in all_dim_master:
                        data_list = all_dim_master[current_id]
                        if current_ptr < len(data_list):
                            vals = data_list[current_ptr][mode]
                            current_ptr += 1

                # หยอดข้อมูลลง M, N, O (13, 14, 15)
                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        target = get_writable_cell(ws, r, col)
                        target.value = vals[i]
                        target.number_format = '0.00000'

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- 4. UI ---
st.title("📋 Master Auto-Filler V33 (Full Extended Version)")
st.info("รวม Logic v14 (All Dimension) + การกวาดพิกัด LAT (B1-B80, N7, N11)")

all_f = st.file_uploader("1. ไฟล์ All Dimension (สำหรับ A, T, B81 ขึ้นไป)", type=["xlsx"])
lat_f = st.file_uploader("2. ไฟล์ LAT (สำหรับ B1-B80, N7, N11)", type=["xlsx"])

if all_f and lat_f:
    if st.button("🚀 ประมวลผลข้อมูลทั้งหมด"):
        try:
            result = process_data(all_f, lat_f)
            st.success("✅ รวมข้อมูลเสร็จสมบูรณ์! ตรวจสอบ A, T, B1-B80, B81+ ได้เลย")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="V33_Master_Result.xlsx")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
