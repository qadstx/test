import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math

st.set_page_config(page_title="Excel Auto-Filler v24", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABTN])-?(\d+)', text)
    if match: return f"{match.group(1)}{match.group(2)}"
    return text

def clean_number(value):
    try:
        if value is None or pd.isna(value): return None
        num = abs(float(value))
        return math.floor(num * 100000) / 100000
    except: return value

def get_merged_origin_value(ws, cell):
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return cell.value

def scan_lat_data(df):
    lat_map = {}
    
    # --- 1. อ่าน B1-B80 และ N10 ตามปกติ (ใช้ Header Scan) ---
    sections = [
        {'name_row': 11, 'data_start': 29},   # B1-B32
        {'name_row': 65, 'data_start': 83},   # B33-B52
        {'name_row': 119, 'data_start': 136}, # B53-B68
        {'name_row': 172, 'data_start': 189}, # B69-B80
        {'name_row': 384, 'data_start': 401}  # N10
    ]
    for sec in sections:
        if sec['name_row'] >= len(df): continue
        names = df.iloc[sec['name_row']]
        for col_idx, raw_name in enumerate(names):
            base_id = get_base_id(raw_name)
            if base_id and base_id != "N7": # แยก N7 ไปทำข้างล่าง
                d1 = clean_number(df.iloc[sec['data_start'], col_idx])
                d2 = clean_number(df.iloc[sec['data_start']+1, col_idx])
                d3 = clean_number(df.iloc[sec['data_start']+2, col_idx])
                if base_id not in lat_map: lat_map[base_id] = []
                lat_map[base_id].append([d1, d2, d3])

    # --- 2. อ่าน N7 แบบล็อคช่วงเซลล์ (ตามคำสั่ง) ---
    n7_all_data = []
    # ช่วงที่ 1: แถว 296-298 (B-AG -> index 1-32)
    # หมายเหตุ: index ใน pandas เริ่มจาก 0 ดังนั้นแถว 296 คือ index 295
    r1, r2, r3 = 295, 296, 297
    if r3 < len(df):
        for c in range(1, 33): # Column B(1) ถึง AG(32)
            if c < df.shape[1]:
                n7_all_data.append([clean_number(df.iloc[r1, c]), clean_number(df.iloc[r2, c]), clean_number(df.iloc[r3, c])])

    # ช่วงที่ 2: แถว 349-351 (B-AG -> index 1-32)
    r1, r2, r3 = 348, 349, 350
    if r3 < len(df):
        for c in range(1, 33):
            if c < df.shape[1]:
                n7_all_data.append([clean_number(df.iloc[r1, c]), clean_number(df.iloc[r2, c]), clean_number(df.iloc[r3, c])])

    # ช่วงที่ 3: แถว 402-404 (B-Q -> index 1-16)
    r1, r2, r3 = 401, 402, 403
    if r3 < len(df):
        for c in range(1, 17): # Column B(1) ถึง Q(16)
            if c < df.shape[1]:
                n7_all_data.append([clean_number(df.iloc[r1, c]), clean_number(df.iloc[r2, c]), clean_number(df.iloc[r3, c])])
    
    lat_map["N7"] = n7_all_data
    return lat_map

def process_data(template_path, all_dim_file, lat_file):
    wb = openpyxl.load_workbook(template_path)
    df_all = pd.read_excel(all_dim_file, header=None)
    
    # หาคอลัมน์ Cavity ใน All Dimension
    cav1_idx, cav2_idx = None, None
    for _, row in df_all.iterrows():
        indices = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(indices) >= 2:
            cav1_idx, cav2_idx = indices[0], indices[1]
            break
    if cav1_idx is None: cav1_idx, cav2_idx = 8, 16

    all_dim_master = {}
    for _, row in df_all.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if not re.search(r'[ABTN]', raw_name): continue
        bid = get_base_id(raw_name)
        if bid not in all_dim_master: all_dim_master[bid] = []
        all_dim_master[bid].append({
            'cav1': [clean_number(row.iloc[cav1_idx]), clean_number(row.iloc[cav1_idx+1]), clean_number(row.iloc[cav1_idx+2])],
            'cav2': [clean_number(row.iloc[cav2_idx]), clean_number(row.iloc[cav2_idx+1]), clean_number(row.iloc[cav2_idx+2])]
        })

    lat_source = {'cav1': {}, 'cav2': {}}
    if lat_file:
        for m_k, s_n in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
            try:
                df_lat = pd.read_excel(lat_file, sheet_name=s_n, header=None)
                lat_source[m_k] = scan_lat_data(df_lat)
            except: continue

    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        current_id, current_ptr = None, 0
        n7_ptr = 0
        n7_queue = lat_source[mode].get("N7", [])

        for r in range(1, ws.max_row + 1):
            # --- พิเศษ N7 (1977-2056) ---
            if 1977 <= r <= 2056:
                if n7_ptr < len(n7_queue):
                    vals = n7_queue[n7_ptr]
                    for i, col in enumerate([13, 14, 15]):
                        cell = ws.cell(row=r, column=col)
                        if not isinstance(cell, MergedCell):
                            cell.value = vals[i]
                            cell.number_format = '0.00000'
                    n7_ptr += 1
                continue

            # --- Logic ปกติ ---
            val_a = get_merged_origin_value(ws, ws.cell(row=r, column=1))
            if val_a is not None:
                new_id = get_base_id(val_a)
                if new_id != current_id:
                    current_id, current_ptr = new_id, 0
            
            l_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower()
            if "supplier" in l_val and current_id:
                target_id = "N10" if current_id == "N11" else current_id
                is_b_lat = target_id.startswith('B') and any(target_id == f"B{i}" for i in range(1, 81))
                
                vals = None
                if (is_b_lat or (target_id.startswith('N') and target_id != 'N7')) and target_id in lat_source[mode]:
                    data_list = lat_source[mode][target_id]
                    if current_ptr < len(data_list):
                        vals = data_list[current_ptr]
                        current_ptr += 1
                elif target_id in all_dim_master:
                    if current_ptr < len(all_dim_master[target_id]):
                        vals = all_dim_master[target_id][current_ptr][mode]
                        current_ptr += 1

                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        cell = ws.cell(row=r, column=col)
                        if not isinstance(cell, MergedCell):
                            cell.value = vals[i]
                            cell.number_format = '0.00000'

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# UI
st.header("📋 Auto-Filler v24 (N7 Exact Range Mapping)")
st.info("N7 Mapping: 296-298(B-AG), 349-351(B-AG), 402-404(B-Q)")

all_f = st.file_uploader("📂 ไฟล์ All dimension", type=["xlsx"])
lat_f = st.file_uploader("📂 ไฟล์ LAT", type=["xlsx"])

if all_f and lat_f:
    if st.button("🚀 รันการประมวลผล V24"):
        try:
            res = process_data(TEMPLATE_FILENAME, all_f, lat_f)
            st.success("✅ ประมวลผลสำเร็จตามช่วงเซลล์ที่กำหนด!")
            st.download_button("📥 ดาวน์โหลด", res, file_name="V24_N7_Exact_Result.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
