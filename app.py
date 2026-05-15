import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re

st.set_page_config(page_title="Excel Auto-Filler v19", layout="wide")
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
        {'name_row': 172, 'data_start': 189}, # B69-B80
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
                    row_idx = sec['data_start'] + i
                    if row_idx < len(df):
                        data_points.append(clean_number(df.iloc[row_idx, col_idx]))
                # เก็บข้อมูลไว้ใน List เพื่อรองรับการมีหลายชุด (โดยเฉพาะ N7)
                if base_id not in lat_map: lat_map[base_id] = []
                lat_map[base_id].append(data_points)
    return lat_map

def process_data(template_path, all_dim_file, lat_file):
    wb = openpyxl.load_workbook(template_path)
    
    # 1. โหลด All Dimension
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

    # 2. โหลด LAT
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
        n7_global_ptr = 0 # ตัวนับพิเศษสำหรับ N7

        for r in range(1, ws.max_row + 1):
            # --- ข้อยกเว้น N7: ช่วงแถว 1977-2056 กรอกต่อเนื่องไม่ต้องเช็ค Supplier ---
            if 1977 <= r <= 2056:
                if 'N7' in lat_source[mode]:
                    # รวมข้อมูล N7 ทั้งหมดที่มี (Flatten list)
                    all_n7_data = [item for sublist in lat_source[mode]['N7'] for item in sublist]
                    if n7_global_ptr < len(all_n7_data):
                        # หยอดลงคอลัมน์ M (13)
                        cell = ws.cell(row=r, column=13)
                        if not isinstance(cell, MergedCell):
                            cell.value = all_n7_data[n7_global_ptr]
                            n7_global_ptr += 1
                continue

            # --- Logic ปกติสำหรับแถวอื่นๆ ---
            val_a = get_merged_origin_value(ws, ws.cell(row=r, column=1))
            if val_a is not None:
                new_id = get_base_id(val_a)
                if new_id != current_id:
                    current_id, current_ptr = new_id, 0
            
            l_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower()
            if "supplier" in l_val and current_id:
                # ข้อยกเว้น N10: ถ้าเจอ N11 ในฟอร์ม ให้ไปดึงข้อมูล N10 มาใส่
                target_id = "N10" if current_id == "N11" else current_id
                
                is_lat_range = target_id.startswith('B') and any(target_id == f"B{i}" for i in range(1, 83))
                is_n_lat = target_id.startswith('N')
                
                vals = None
                if (is_lat_range or is_n_lat) and target_id in lat_source[mode]:
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

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# UI
st.header("📋 Auto-Filler v19 (N7 Continuous & N11 Redirection)")
all_f = st.file_uploader("📂 อัพโหลดไฟล์ All dimension", type=["xlsx"])
lat_f = st.file_uploader("📂 อัพไฟล์ LAT", type=["xlsx"])

if all_f and lat_f:
    if st.button("🚀 รันการประมวลผล"):
        try:
            res = process_data(TEMPLATE_FILENAME, all_f, lat_f)
            st.success("✅ ประมวลผลสำเร็จ! N7 กรอกต่อเนื่องแถว 1977+ และ N11 ดึงจาก N10 เรียบร้อย")
            st.download_button("📥 ดาวน์โหลด", res, file_name="V19_Special_N_Result.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
