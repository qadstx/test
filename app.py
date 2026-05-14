import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v10", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().strip()
    match = re.search(r'[AB]\d+', text)
    return match.group(0).lower() if match else text.lower()

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    # อ่านไฟล์ข้อมูลแบบดิบๆ มาสแกนหาพิกัด
    df_raw = pd.read_excel(data_file, header=None)
    
    # 1. ค้นหาพิกัดคอลัมน์ '1' ของ Cav 1 และ Cav 2
    cav1_col_idx = None
    cav2_col_idx = None
    header_row_idx = None

    for r_idx, row in df_raw.iterrows():
        row_values = [str(v).strip() for v in row]
        if "Item" in row_values:
            header_row_idx = r_idx
        # หาจุดที่มีเลข 1 สองตัว (ปกติ Cav 1 มาก่อน Cav 2)
        indices = [i for i, v in enumerate(row_values) if v == '1']
        if len(indices) >= 2:
            cav1_col_idx = indices[0]
            cav2_col_idx = indices[1]
            break
    
    if cav1_col_idx is None:
        # Fallback: ถ้าหาเลข 1 ไม่เจอ ให้ใช้ค่า Standard (G=6, N=13 หรือใกล้เคียง)
        cav1_col_idx, cav2_col_idx = 8, 16 

    # 2. จัดระเบียบข้อมูล Source
    # เริ่มอ่านข้อมูลจริงใต้แถว Header
    data_start_row = (header_row_idx + 2) if header_row_idx is not None else 11
    df_data = pd.read_excel(data_file, skiprows=data_start_row, header=None)
    
    source_master = {}
    for _, row in df_data.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if not raw_name or raw_name == "nan": continue
        
        base_id = get_base_id(raw_name)
        if base_id not in source_master:
            source_master[base_id] = []

        # ดึงข้อมูล Cav 1 และ Cav 2 (เอามา 3 ช่อง: 1, 2, 3)
        try:
            c1 = [row.iloc[cav1_col_idx], row.iloc[cav1_col_idx+1], row.iloc[cav1_col_idx+2]]
            c2 = [row.iloc[cav2_col_idx], row.iloc[cav2_col_idx+1], row.iloc[cav2_col_idx+2]]
            
            source_master[base_id].append({
                'name': raw_name,
                'cav1': c1,
                'cav2': c2
            })
        except:
            continue

    # เรียงลำดับ A1-1, A1-2...
    for bid in source_master:
        source_master[bid].sort(key=lambda x: natural_sort_key(x['name']))

    # 3. กรอกลง Template
    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_id = None
        fill_count = 0

        for r in range(8, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                current_id = get_base_id(cell_a)
                fill_count = 0
            
            corr = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in corr and current_id in source_master:
                data_list = source_master[current_id]
                if fill_count < len(data_list):
                    vals = data_list[fill_count][mode]
                    # กรอกข้อมูลลงช่อง M, N, O (13, 14, 15)
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                    fill_count += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Filler V10 (Coordinate Scanner)")
uploaded_file = st.file_uploader("อัพโหลดไฟล์ Source Data", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 รันการประมวลผล"):
        try:
            result = process_data(TEMPLATE_FILENAME, uploaded_file)
            st.success("✅ ประมวลผลสำเร็จ!")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="V10_Final_Result.xlsx")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.info("คำแนะนำ: ตรวจสอบว่าในไฟล์ข้อมูลมีแถวที่มีเลข '1' อย่างน้อย 2 จุด (สำหรับ Cav 1/2) หรือไม่")
