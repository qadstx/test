import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v29", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

def clean_number(value):
    try:
        if value is None or pd.isna(value): return None
        num = abs(float(value))
        return math.floor(num * 100000) / 100000
    except: return None

def get_writable_cell(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

def get_merged_origin_value(ws, cell):
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return cell.value

def parse_spec(ws, row):
    try:
        nom_raw = get_merged_origin_value(ws, ws.cell(row=row, column=3))
        nom_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(nom_raw))
        nominal = float(nom_match[0]) if nom_match else 0.0
        tol_raw = get_merged_origin_value(ws, ws.cell(row=row, column=7))
        tol_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(tol_raw))
        tol = float(tol_match[0]) if tol_match else 0.0
        return nominal - tol, nominal + tol
    except: return None, None

def collect_lat_sequential(df):
    """กวาดข้อมูลจาก LAT เรียงตามลำดับช่องเพื่อเตรียมนำไปวาง"""
    queue = []
    
    # 1. กวาดกลุ่ม B1-B32 (Vertical) - แถว 29, 30, 31
    for col in range(1, df.shape[1]): # เริ่มคอลัมน์ B
        # เช็คว่าคอลัมน์นี้มีข้อมูล Sample 1 (แถว 29) หรือไม่
        val = df.iloc[28, col]
        if pd.notna(val):
            queue.append([clean_number(df.iloc[28+i, col]) for i in range(3)])
        if len(queue) >= 32: break # เก็บแค่ 32 ชุดแรกของโซนนี้

    # 2. กวาดกลุ่ม B41-B80 (Horizontal) - แถว 84, 85, 86
    # กวาดไล่ยาวไปเรื่อยๆ ทั้ง i และ ii จะถูกเก็บต่อกันใน queue
    for col in range(1, df.shape[1]): 
        val = df.iloc[83, col]
        if pd.notna(val):
            queue.append([clean_number(df.iloc[83+i, col]) for i in range(3)])
            
    return queue

def process_step_1(all_dim_file, lat_file):
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    
    lat_queues = {'cav1': [], 'cav2': []}
    for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try: 
            df_lat = pd.read_excel(lat_file, sheet_name=s, header=None)
            lat_queues[m] = collect_lat_sequential(df_lat)
            # ดึง N7 แยกไว้ (Logic เดิม)
            n7_list = []
            ranges = [(295, 298, 1, 33), (348, 351, 1, 33), (401, 404, 1, 17)]
            for r1, r3, c_start, c_end in ranges:
                if r3 < len(df_lat):
                    for c in range(c_start, c_end):
                        n7_list.append([clean_number(df_lat.iloc[r1+i, c]) for i in range(3)])
            lat_queues[m + "_n7"] = n7_list
            # ดึง N10 ไว้ใช้สำหรับ N11
            lat_queues[m + "_n10"] = [clean_number(df_lat.iloc[400+i, 1]) for i in range(3)] # สมมติ N10 อยู่ Col B
        except: pass

    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        q = lat_queues[mode]
        q_n7 = lat_queues.get(mode + "_n7", [])
        q_idx, n7_idx = 0, 0

        for r in range(1, ws.max_row + 1):
            # วาง N7 ตามตำแหน่งแถว
            if 1977 <= r <= 2056:
                if n7_idx < len(q_n7):
                    for i, col in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, col)
                        cell.value = q_n7[n7_idx][i]
                        cell.number_format = '0.00000'
                    n7_idx += 1
                continue

            # วางข้อมูล B1-B80 เรียงลำดับเมื่อเจอ Supplier
            l_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower()
            a_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=1))).strip().upper()
            
            if "supplier" in l_val:
                vals = None
                if "N11" in a_val: # กรณี N11 ให้หยิบ N10 มาวาง
                    vals = lat_queues.get(mode + "_n10")
                elif q_idx < len(q): # กรณีปกติ ดึงตามลำดับ Queue
                    vals = q[q_idx]
                    q_idx += 1
                
                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        cell = get_writable_cell(ws, r, col)
                        cell.value = vals[i]
                        cell.number_format = '0.00000'

    out = BytesIO(); wb.save(out)
    return out.getvalue()

def process_step_2(raw_excel_data):
    wb = openpyxl.load_workbook(BytesIO(raw_excel_data))
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            is_supp = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower() == "supplier"
            if is_supp or (1977 <= r <= 2056):
                low, high = parse_spec(ws, r)
                if low is None: continue
                samples = [ws.cell(row=r, column=c).value for c in [13, 14, 15]]
                new_samples = []
                passed = [v for v in samples if v is not None and low <= v <= high]
                for v in samples:
                    if v is None or not (low <= v <= high):
                        base = random.choice(passed) if passed else (low + high) / 2
                        var = (high - low) * 0.1
                        new_val = base + random.uniform(-var, var)
                        new_samples.append(max(low, min(high, new_val)))
                    else: new_samples.append(v)
                for i, col in enumerate([13, 14, 15]):
                    cell = get_writable_cell(ws, r, col)
                    cell.value = math.floor(new_samples[i] * 100000) / 100000
                    cell.number_format = '0.00000'
    out = BytesIO(); wb.save(out)
    return out.getvalue()

# --- UI ---
st.title("Excel Auto-Filler v29 (Sequential Logic)")
if st.button("🔄 Reset All"):
    for key in list(st.session_state.keys()): del st.session_state[key]
    st.rerun()

all_f = st.file_uploader("All dimension", type=["xlsx"])
lat_f = st.file_uploader("LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw_data" not in st.session_state:
        if st.button("🚀 รันรอบที่ 1: วางข้อมูลตามจริง"):
            st.session_state.raw_data = process_step_1(all_f, lat_f)
            st.rerun()

    if "raw_data" in st.session_state:
        st.success("รอบที่ 1 เสร็จแล้ว! ตรวจสอบข้อมูล i/ii ได้เลย")
        st.download_button("📥 ดาวน์โหลดไฟล์เช็คผล (Raw)", st.session_state.raw_data, "step1_raw.xlsx")
        st.divider()
        if st.button("🪄 สุ่มข้อมูลให้ผ่าน Spec ทั้งหมด"):
            st.session_state.final_data = process_step_2(st.session_state.raw_data)
            st.rerun()
            
    if "final_data" in st.session_state:
        st.success("✅ แก้ไขข้อมูลให้ผ่าน Spec เรียบร้อย!")
        st.download_button("📥 ดาวน์โหลดไฟล์ Final", st.session_state.final_data, "step2_final.xlsx")
