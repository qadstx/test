import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v16", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    # ปรับให้รองรับ A, B, T, N และดึงตัวเลขชุดแรกมาจับคู่
    match = re.search(r'([ABTN])-?(\d+)', text)
    if match: return f"{match.group(1)}{match.group(2)}"
    return text

def clean_number(value):
    try:
        if value is None or pd.isna(value): return None
        return abs(float(value))
    except: return value

def scan_lat_data(df):
    """สแกนหา B1-B82 และ N7, N10 ในไฟล์ LAT แบบระบุพิกัดแถว"""
    lat_map = {}
    # เพิ่มพิกัดชุด N7 และ N10 เข้าไป (index ใน pandas เริ่มจาก 0 ดังนั้นต้องลบจากเลขแถวใน Excel ไป 1)
    sections = [
        {'name_row': 11, 'data_start': 29},   # B12 -> B30 (ชุด 1)
        {'name_row': 65, 'data_start': 83},   # B66 -> B84 (ชุด 2)
        {'name_row': 119, 'data_start': 136}, # B120 -> B137 (ชุด 3)
        {'name_row': 172, 'data_start': 189}, # B173 -> B190 (ชุด 4)
        {'name_row': 278, 'data_start': 295}, # แถว 279 -> 296 (N7 ชุด 1)
        {'name_row': 329, 'data_start': 348}, # แถว 330 -> 349 (N7 ชุด 2)
        {'name_row': 384, 'data_start': 401}  # แถว 385 -> 402 (N7/N10 ชุดสุดท้าย)
    ]
    
    for sec in sections:
        if sec['name_row'] >= len(df): continue
        names = df.iloc[sec['name_row']]
        for col_idx, raw_name in enumerate(names):
            base_id = get_base_id(raw_name)
            # รองรับทั้ง Item B และ Item N
            if base_id and (base_id.startswith('B') or base_id.startswith('N')):
                data_points = []
                for i in range(3):
                    row_idx = sec['data_start'] + i
                    if row_idx < len(df):
                        data_points.append(clean_number(df.iloc[row_idx, col_idx]))
                lat_map[base_id] = data_points
    return lat_map

def process_data(template_path, all_dim_file, lat_file):
    wb = openpyxl.load_workbook(template_path)
    
    # 1. จัดการไฟล์ All Dimension
    df_all = pd.read_excel(all_dim_file, header=None)
    cav1_idx, cav2_idx = None, None
    for _, row in df_all.iterrows():
        indices = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(indices) >= 2:
            cav1_idx, cav2_idx = indices[0], indices[1]
            break
    if cav1_idx is None: cav1_idx, cav2_idx = 8, 16

    source_master = {}
    for _, row in df_all.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if not re.search(r'[ABTN]', raw_name): continue
        bid = get_base_id(raw_name)
        if bid not in source_master: source_master[bid] = []
        source_master[bid].append({
            'cav1': [clean_number(row.iloc[cav1_idx]), clean_number(row.iloc[cav1_idx+1]), clean_number(row.iloc[cav1_idx+2])],
            'cav2': [clean_number(row.iloc[cav2_idx]), clean_number(row.iloc[cav2_idx+1]), clean_number(row.iloc[cav2_idx+2])]
        })

    # 2. จัดการไฟล์ LAT
    lat_source = {'cav1': {}, 'cav2': {}}
    if lat_file:
        # โหลดทั้ง 2 Sheet ของไฟล์ LAT
        df_lat1 = pd.read_excel(lat_file, sheet_name="LAT (Cav 1)", header=None)
        df_lat2 = pd.read_excel(lat_file, sheet_name="LAT (Cav 2)", header=None)
        lat_source['cav1'] = scan_lat_data(df_lat1)
        lat_source['cav2'] = scan_lat_data(df_lat2)

    # 3. หยอดลง Template
    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        current_id, current_ptr = None, 0

        for r in range(1, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                new_id = get_base_id(cell_a)
                if new_id != current_id:
                    current_id, current_ptr = new_id, 0
            
            l_val = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in l_val and current_id:
                # ลำดับความสำคัญ: ถ้ามีในไฟล์ LAT ให้ใช้ LAT ก่อน (ครอบคลุม B1-B82 และ N7, N10)
                if current_id in lat_source[mode]:
                    vals = lat_source[mode][current_id]
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                # ถ้าไม่มีใน LAT ให้ไปดูใน All Dimension (A, T และ B อื่นๆ)
                elif current_id in source_master:
                    data_list = source_master[current_id]
                    if current_ptr < len(data_list):
                        vals = data_list[current_ptr][mode]
                        ws.cell(row=r, column=13).value = vals[0]
                        ws.cell(row=r, column=14).value = vals[1]
                        ws.cell(row=r, column=15).value = vals[2]
                        current_ptr += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Filler v16 (LAT N-Series Support)")

col1, col2 = st.columns(2)
with col1:
    all_dim_file = st.file_uploader("📂 กรุณาอัพโหลดไฟล์ All dimension", type=["xlsx"])
with col2:
    lat_file = st.file_uploader("📂 กรุณาอัพไฟล์ LAT (B1-B82, N7, N10)", type=["xlsx"])

if all_dim_file and lat_file:
    if st.button("🚀 เริ่มการประมวลผล"):
        try:
            result = process_data(TEMPLATE_FILENAME, all_dim_file, lat_file)
            st.success("✅ ประมวลผลรหัส B และ N เรียบร้อยแล้ว!")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="LAT_N_Series_Result.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
