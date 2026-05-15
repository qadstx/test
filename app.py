import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v18", layout="wide")
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
        return abs(float(value))
    except: return value

def get_merged_origin_value(ws, cell):
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return cell.value

def scan_lat_data(df):
    lat_map = {}
    sections = [
        {'name_row': 11, 'data_start': 29},   # B1-B32
        {'name_row': 65, 'data_start': 83},   # B33-B52
        {'name_row': 119, 'data_start': 136}, # B53-B68
        {'name_row': 172, '0': 189},          # B69-B80
        {'name_row': 278, 'data_start': 295}, # N7 ชุด 1
        {'name_row': 329, 'data_start': 348}, # N7 ชุด 2
        {'name_row': 384, 'data_start': 401}  # N7/N10
    ]
    for sec in sections:
        if sec['name_row'] >= len(df): continue
        names = df.iloc[sec['name_row']]
        for col_idx, raw_name in enumerate(names):
            base_id = get_base_id(raw_name)
            if base_id and (base_id.startswith('B') or base_id.startswith('N')):
                data_points = []
                for i in range(3):
                    row_idx = sec.get('data_start', 189) + i # แก้ไขพิกัดชุด 4
                    if row_idx < len(df):
                        data_points.append(clean_number(df.iloc[row_idx, col_idx]))
                lat_map[base_id] = data_points
    return lat_map

def process_data(template_path, all_dim_file, lat_file):
    wb = openpyxl.load_workbook(template_path)
    
    # 1. โหลดข้อมูล All Dimension (Logic เดิม)
    df_all = pd.read_excel(all_dim_file, header=None)
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

    # 2. โหลดข้อมูล LAT
    lat_source = {'cav1': {}, 'cav2': {}}
    if lat_file:
        for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
            try:
                df_lat = pd.read_excel(lat_file, sheet_name=s, header=None)
                lat_source[m] = scan_lat_data(df_lat)
            except: continue

    # 3. หยอดลง Template
    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        current_id, current_ptr = None, 0

        for r in range(1, ws.max_row + 1):
            val_a = get_merged_origin_value(ws, ws.cell(row=r, column=1))
            if val_a is not None:
                new_id = get_base_id(val_a)
                if new_id != current_id:
                    current_id, current_ptr = new_id, 0
            
            l_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower()
            if "supplier" in l_val and current_id:
                # แยกเช็คช่วง LAT (B1-B82, N7, N10)
                is_b_lat = current_id.startswith('B') and any(str(current_id) == f"B{i}" for i in range(1, 83))
                is_n_lat = current_id.startswith('N')
                
                vals = None
                if (is_b_lat or is_n_lat) and current_id in lat_source[mode]:
                    vals = lat_source[mode][current_id]
                elif current_id in all_dim_master:
                    # ใช้ Logic เดิมที่เคยดีอยู่แล้วสำหรับไอเทมอื่นๆ (เช่น B181)
                    if current_ptr < len(all_dim_master[current_id]):
                        vals = all_dim_master[current_id][current_ptr][mode]
                        current_ptr += 1

                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        cell = ws.cell(row=r, column=col)
                        if not isinstance(cell, MergedCell):
                            cell.value = vals[i]

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# UI (เหมือนเดิม)
st.header("📋 Auto-Filler v18 (Original Logic Protected)")
all_file = st.file_uploader("📂 กรุณาอัพโหลดไฟล์ All dimension", type=["xlsx"])
lat_file = st.file_uploader("📂 กรุณาอัพไฟล์ LAT", type=["xlsx"])

if all_file and lat_file:
    if st.button("🚀 รันการประมวลผล"):
        try:
            res = process_data(TEMPLATE_FILENAME, all_file, lat_file)
            st.success("✅ ประมวลผลสำเร็จ! (B181 และไอเทมอื่นใช้ Logic เดิมเรียบร้อย)")
            st.download_button("📥 ดาวน์โหลด", res, file_name="V18_Final_Result.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
