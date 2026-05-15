import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v32", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

# --- Utility Functions ---
def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABTN])(\d+)', text)
    if match: return f"{match.group(1)}{match.group(2)}"
    return text

def clean_number(value):
    """แปลงค่าให้เป็นตัวเลขบวก 5 ตำแหน่ง (ตาม Logic v14 เดิมของคุณ)"""
    try:
        if value is None or pd.isna(value): return None
        num = abs(float(value))
        return math.floor(num * 100000) / 100000
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

# --- Data Scanners ---
def scan_lat_sequential(df):
    """กวาดข้อมูล B1-B80, N7, N10 ตามพิกัด"""
    data = {"B_QUEUE": [], "N7_QUEUE": [], "N10_VAL": None}
    
    # B1-B80 Horizontal Scan
    b_ranges = [(29, 31, 1, 33), (83, 85, 1, 33), (136, 138, 1, 33), (189, 191, 1, 25)]
    for r1, r3, c1, c2 in b_ranges:
        for c in range(c1, c2):
            if r3 < len(df) and c < df.shape[1]:
                data["B_QUEUE"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])

    # N7 Scan
    n7_ranges = [(295, 297, 1, 33), (348, 350, 1, 33), (401, 403, 1, 17)]
    for r1, r3, c1, c2 in n7_ranges:
        for c in range(c1, c2):
            if r3 < len(df) and c < df.shape[1]:
                data["N7_QUEUE"].append([clean_number(df.iloc[r1+i, c]) for i in range(3)])

    # N10 for N11 (Col N = 13)
    if 403 < len(df):
        data["N10_VAL"] = [clean_number(df.iloc[400+i, 13]) for i in range(3)]
        
    return data

# --- Core Processor ---
def process_master(all_dim_file, lat_file):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    df_all = pd.read_excel(all_dim_file, header=None)
    
    # 1. เตรียมข้อมูลจาก All Dimension (สำหรับ A, T, B81+)
    all_dim_master = {}
    c1_idx, c2_idx = 8, 16 
    for _, row in df_all.iterrows():
        bid = get_base_id(row.iloc[0])
        if not bid: continue
        if bid not in all_dim_master: all_dim_master[bid] = []
        all_dim_master[bid].append({
            'cav1': [clean_number(row.iloc[c1_idx+i]) for i in range(3)],
            'cav2': [clean_number(row.iloc[c2_idx+i]) for i in range(3)]
        })

    # 2. เตรียมข้อมูลจาก LAT (B1-B80, N7, N11)
    lat_data = {}
    for mode, sheet in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try:
            df_lat = pd.read_excel(lat_file, sheet_name=sheet, header=None)
            lat_data[mode] = scan_lat_sequential(df_lat)
        except: lat_data[mode] = None

    # 3. หยอดข้อมูลลงใน Template
    for sn in ["Cav 1", "Cav 2"]:
        if sn not in wb.sheetnames: continue
        ws = wb[sn]
        m = 'cav1' if sn == "Cav 1" else 'cav2'
        
        b_idx, n7_idx, curr_id, curr_ptr = 0, 0, None, 0

        for r in range(1, ws.max_row + 1):
            # โซน N7
            if 1977 <= r <= 2056:
                if lat_data[m] and n7_idx < len(lat_data[m]["N7_QUEUE"]):
                    for i, col in enumerate([13, 14, 15]):
                        get_writable_cell(ws, r, col).value = lat_data[m]["N7_QUEUE"][n7_idx][i]
                    n7_idx += 1
                continue

            # โซนปกติ (A, B, T, N)
            id_val = get_base_id(get_val(ws, r, 1))
            if id_val:
                if id_val != curr_id:
                    curr_id, curr_ptr = id_val, 0
            
            l_val = str(get_val(ws, r, 12)).lower()
            if "supplier" in l_val and curr_id:
                vals = None
                
                # ตรวจสอบว่าเป็น B1-B80 หรือไม่
                match = re.search(r'\d+', curr_id)
                num = int(match.group()) if match else 0
                
                if curr_id.startswith("B") and 1 <= num <= 80:
                    # ใช้ Sequential Queue จาก LAT
                    if lat_data[m] and b_idx < len(lat_data[m]["B_QUEUE"]):
                        vals = lat_data[m]["B_QUEUE"][b_idx]
                        b_idx += 1
                elif curr_id == "N11":
                    # ใช้ N10 จาก LAT
                    if lat_data[m]: vals = lat_data[m]["N10_VAL"]
                else:
                    # ส่วนที่เหลือ (A, T, B81+) ดึงจาก All Dimension เหมือนเดิม
                    if curr_id in all_dim_master and curr_ptr < len(all_dim_master[curr_id]):
                        vals = all_dim_master[curr_id][curr_ptr][m]
                        curr_ptr += 1
                
                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, col)
                        cell.value = vals[i]
                        cell.number_format = '0.00000'

    out = BytesIO(); wb.save(out)
    return out.getvalue()

# --- UI ---
st.title("📋 Master Auto-Filler V32")
st.markdown("""
**ระบบรวม 3 พลัง:**
1. **B1-B80:** ดึงตามพิกัด LAT แนวนอน (ครบ i/ii)
2. **N7/N11:** ดึงตามพิกัด LAT เฉพาะจุด
3. **อื่นๆ (A, T, B81+):** ดึงจากไฟล์ All Dimension ตามชื่อ Item
""")

all_f = st.file_uploader("1. อัปโหลดไฟล์ All Dimension", type=["xlsx"])
lat_f = st.file_uploader("2. อัปโหลดไฟล์ LAT", type=["xlsx"])

if all_f and lat_f:
    if st.button("🚀 เริ่มการประมวลผลทั้งหมด"):
        try:
            result = process_master(all_f, lat_f)
            st.success("✅ รวมข้อมูลทุกส่วนเรียบร้อยแล้ว!")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="Master_Combined_Result.xlsx")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
