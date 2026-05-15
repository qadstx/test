import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math

st.set_page_config(page_title="Excel Auto-Filler v20", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABTN])-?(\d+)', text)
    if match: return f"{match.group(1)}{match.group(2)}"
    return text

def clean_number(value):
    """ตัดทศนิยมทิ้งที่หลักที่ 5 (ไม่ปัดเศษ) และทำให้เป็นบวก"""
    try:
        if value is None or pd.isna(value): return None
        num = abs(float(value))
        # ใช้เทคนิคคูณ 100000 แล้วตัดเศษทิ้งก่อนหารกลับ เพื่อไม่ให้เกิดการปัดเศษ
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
    sections = [
        {'name_row': 11, 'data_start': 29},   # B1-B32
        {'name_row': 65, 'data_start': 83},   # B33-B52
        {'name_row': 119, 'data_start': 136}, # B53-B68
        {'name_row': 172, 'data_start': 189}, # B69-B80
        {'name_row': 278, 'data_start': 295}, # N7 (1)
        {'name_row': 329, 'data_start': 348}, # N7 (2)
        {'name_row': 384, 'data_start': 401}  # N7/N10
    ]
    for sec in sections:
        if sec['name_row'] >= len(df): continue
        names = df.iloc[sec['name_row']]
        for col_idx, raw_name in enumerate(names):
            base_id = get_base_id(raw_name)
            if base_id and (base_id.startswith('B') or base_id.startswith('N')):
                d1 = clean_number(df.iloc[sec['data_start'], col_idx])
                d2 = clean_number(df.iloc[sec['data_start']+1, col_idx])
                d3 = clean_number(df.iloc[sec['data_start']+2, col_idx])
                if base_id not in lat_map: lat_map[base_id] = []
                lat_map[base_id].append([d1, d2, d3])
    return lat_map

def process_data(template_path, all_dim_file, lat_file):
    wb = openpyxl.load_workbook(template_path)
    
    # 1. จัดการ All Dimension (Logic V14 เดิม)
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

    # 2. จัดการ LAT
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
        
        # เตรียมข้อมูล N7 ให้เป็นสายยาว 1, 2, 3, 1, 2, 3...
        n7_stream = []
        if 'N7' in lat_source[mode]:
            for sets in lat_source[mode]['N7']:
                n7_stream.extend(sets)
        n7_idx = 0

        for r in range(1, ws.max_row + 1):
            # --- พิเศษ N7 (แถว 1977-2056) หยอด Sample 1,2,3 ต่อเนื่อง ---
            if 1977 <= r <= 2056:
                if n7_idx < len(n7_stream):
                    cell = ws.cell(row=r, column=13)
                    if not isinstance(cell, MergedCell):
                        cell.value = n7_stream[n7_idx]
                        cell.number_format = '0.00000' # บังคับทศนิยม 5 หลัก
                        n7_idx += 1
                continue

            # --- Logic ปกติ ---
            val_a = get_merged_origin_value(ws, ws.cell(row=r, column=1))
            if val_a is not None:
                new_id = get_base_id(val_a)
                if new_id != current_id:
                    current_id, current_ptr = new_id, 0
            
            l_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower()
            if "supplier" in l_val and current_id:
                # ข้อยกเว้น N10 -> N11
                target_id = "N10" if current_id == "N11" else current_id
                
                # เช็คขอบเขต B1-B80
                is_b_lat = target_id.startswith('B') and any(target_id == f"B{i}" for i in range(1, 81))
                
                vals = None
                # ถ้าเป็น B1-B80 หรือกลุ่ม N (ยกเว้น N7 ที่ทำไปแล้ว) ให้ดู LAT
                if (is_b_lat or target_id.startswith('N')) and target_id in lat_source[mode]:
                    data_list = lat_source[mode][target_id]
                    if current_ptr < len(data_list):
                        vals = data_list[current_ptr]
                        current_ptr += 1
                # ถ้าเป็น B81 ขึ้นไป หรือ A, T ให้ดู All Dimension (Logic เดิม)
                elif target_id in all_dim_master:
                    if current_ptr < len(all_dim_master[target_id]):
                        vals = all_dim_master[target_id][current_ptr][mode]
                        current_ptr += 1

                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        cell = ws.cell(row=r, column=col)
                        if not isinstance(cell, MergedCell):
                            cell.value = vals[i]
                            cell.number_format = '0.00000' # บังคับโชว์ทศนิยม 5 หลัก

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# UI
st.header("📋 Auto-Filler v20 (Decimal & Logic Final Fix)")
st.warning("Logic: B1-B80 ใช้ LAT | B81+ และอื่นๆ ใช้ All Dimension (3 Samples ครบ)")

all_f = st.file_uploader("📂 ไฟล์ All dimension", type=["xlsx"])
lat_f = st.file_uploader("📂 ไฟล์ LAT", type=["xlsx"])

if all_f and lat_f:
    if st.button("🚀 ประมวลผล"):
        try:
            res = process_data(TEMPLATE_FILENAME, all_f, lat_f)
            st.success("✅ สำเร็จ! ทศนิยมถูกตัดเหลือ 5 หลัก และ Logic B81 กลับมาทำงานแล้ว")
            st.download_button("📥 ดาวน์โหลด", res, file_name="V20_Final_Standard.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
