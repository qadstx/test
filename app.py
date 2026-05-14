import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v9", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    # อ่านแบบไม่ skiprows ก่อนเพื่อหาหัวตารางที่แท้จริง
    df_raw = pd.read_excel(data_file)
    
    # ค้นหาแถวที่มีคำว่า 'Item' เพื่อเป็น Header
    header_row_idx = 0
    for i, row in df_raw.iterrows():
        if "Item" in str(row.iloc[0]):
            header_row_idx = i
            break
    
    df_source = pd.read_excel(data_file, skiprows=header_row_idx + 1)
    # ล้างชื่อคอลัมน์ให้ไม่มีช่องว่าง
    df_source.columns = [str(c).strip() for c in df_source.columns]

    # ระบุคอลัมน์สำหรับ Cav 1 และ Cav 2 (อ้างอิงจากรูป: 1, 2, 3 คือข้อมูลที่เราต้องการ)
    # เราจะหา Index ของคอลัมน์ที่มีชื่อว่า '1', '2', '3'
    cols = df_source.columns.tolist()
    # หาตำแหน่ง '1' ตัวแรก (Cav 1) และ '1' ตัวที่สอง (Cav 2)
    one_indices = [i for i, x in enumerate(cols) if x == '1']
    
    source_master = {}
    for _, row in df_source.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if pd.isna(raw_name) or raw_name == "nan": continue
        
        # คลีน ID: B186-1 -> b186
        base_id = re.sub(r'[- ].*$', '', raw_name).lower() # ตัดตั้งแต่ - หรือช่องว่างทิ้ง
        base_id = re.sub(r'^[ab]_', '', base_id) # ตัด prefix a_ หรือ b_
        
        if base_id not in source_master:
            source_master[base_id] = []

        # ดึงข้อมูล Cav 1 (จากเลข 1, 2, 3 ชุดแรก)
        c1_idx = one_indices[0]
        c1_data = [row.iloc[c1_idx], row.iloc[c1_idx+1], row.iloc[c1_idx+2]]
        
        # ดึงข้อมูล Cav 2 (จากเลข 1, 2, 3 ชุดที่สอง)
        c2_idx = one_indices[1] if len(one_indices) > 1 else c1_idx
        c2_data = [row.iloc[c2_idx], row.iloc[c2_idx+1], row.iloc[c2_idx+2]]
        
        source_master[base_id].append({
            'name': raw_name,
            'cav1': c1_data,
            'cav2': c2_data
        })

    # เรียงลำดับ B186-1, B186-2... ให้ถูกต้อง
    for bid in source_master:
        source_master[bid].sort(key=lambda x: natural_sort_key(x['name']))

    # --- เริ่มกรอกลงฟอร์ม ---
    target_sheets = ["Cav 1", "Cav 2"]
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_id_in_form = None
        fill_count = 0

        for r in range(8, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                # คลีนชื่อในฟอร์มให้เหลือแค่รหัสหลัก (เช่น B186)
                current_id_in_form = str(cell_a).strip().lower()
                fill_count = 0
            
            # เช็คคำว่า Supplier ในคอลัมน์ L (12)
            corr_val = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in corr_val and current_id_in_form in source_master:
                data_list = source_master[current_id_in_form]
                if fill_count < len(data_list):
                    vals = data_list[fill_count][mode]
                    # กรอกลงคอลัมน์ M, N, O (13, 14, 15)
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                    fill_count += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- Streamlit UI ---
st.header("📋 Excel Auto-Filler V9 (Hard-Fixed Columns)")
st.warning("เวอร์ชันนี้จะกวาดหาคอลัมน์ '1', '2', '3' อัตโนมัติ เพื่อป้องกันเลขไม่ตรง")

uploaded_file = st.file_uploader("อัพโหลดไฟล์ Source Data", type=["xlsx"])
if uploaded_file:
    if st.button("🚀 รันการประมวลผล (Final Fix)"):
        try:
            result = process_data(TEMPLATE_FILENAME, uploaded_file)
            st.success("✅ กรอกข้อมูลเรียบร้อย! โปรดตรวจสอบเลข B186-1 อีกครั้ง")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="Final_Fix_Result.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
