import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v11", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def clean_for_match(s):
    """ล้างชื่อให้เหลือแค่รหัสหลักแบบเรียบง่ายที่สุด เช่น B186-1 -> B186"""
    if pd.isna(s): return ""
    text = str(s).upper().strip()
    # หาตัวอักษร A หรือ B แล้วตามด้วยตัวเลข (เอาแค่ชุดแรก)
    match = re.search(r'[AB]\d+', text)
    return match.group(0) if match else text

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    
    # อ่านไฟล์ Source แบบดิบๆ ไม่ต้องสน Header
    df_raw = pd.read_excel(data_file, header=None)
    
    # ค้นหาพิกัดคอลัมน์จากแถวที่มีเลข '1' ปรากฏ (Cav 1 และ Cav 2)
    cav1_idx, cav2_idx = None, None
    for _, row in df_raw.iterrows():
        indices = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(indices) >= 2:
            cav1_idx, cav2_idx = indices[0], indices[1]
            break
    
    # ถ้าหาพิกัดไม่เจอ ให้ใช้ค่า Default (G=8, N=16 ในระบบ index)
    if cav1_idx is None: cav1_idx, cav2_idx = 8, 16

    # สร้าง Dictionary เก็บข้อมูล "ตามลำดับที่ปรากฏในไฟล์"
    # โครงสร้าง: { 'B186': [ [data1], [data2], [data3] ] }
    source_master = {}
    for _, row in df_raw.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if not re.search(r'[AB]\d+', raw_name): continue # ข้ามถ้าไม่ใช่บรรทัด Item
        
        base_id = clean_for_match(raw_name)
        if base_id not in source_master:
            source_master[base_id] = []
            
        # เก็บข้อมูล 3 ช่อง (Sample 1, 2, 3)
        c1 = [row.iloc[cav1_idx], row.iloc[cav1_idx+1], row.iloc[cav1_idx+2]]
        c2 = [row.iloc[cav2_idx], row.iloc[cav2_idx+1], row.iloc[cav2_idx+2]]
        
        source_master[base_id].append({'cav1': c1, 'cav2': c2})

    # --- เริ่มหยอดลง Template ---
    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_id = None
        current_pointer = 0 # ตัวชี้ว่าตอนนี้ใช้ Data ลำดับที่เท่าไหร่ของ Item นั้น

        for r in range(1, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                new_id = clean_for_match(cell_a)
                if new_id != current_id: # ถ้าเปลี่ยน Item (เช่น จาก B185 เป็น B186)
                    current_id = new_id
                    current_pointer = 0 # รีเซ็ตมาเริ่มใช้ข้อมูลตัวแรกใหม่
            
            # ตรวจสอบว่าเป็นช่อง Supplier (คอลัมน์ L/12)
            l_val = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in l_val and current_id in source_master:
                data_list = source_master[current_id]
                if current_pointer < len(data_list):
                    vals = data_list[current_pointer][mode]
                    # หยอดลง M, N, O (13, 14, 15)
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                    current_pointer += 1 # หยอดเสร็จ ขยับไปใช้ข้อมูลตัวถัดไปในไฟล์ Source

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Filler v11 (Sequential Fill)")
st.info("Logic: ใครอยู่บนในไฟล์ข้อมูล จะถูกหยอดในช่อง Supplier แรกของ Item นั้นๆ ทันที")

uploaded_file = st.file_uploader("อัพโหลดไฟล์ Source Data", type=["xlsx"])
if uploaded_file:
    if st.button("🚀 รันการหยอดข้อมูล"):
        try:
            result = process_data(TEMPLATE_FILENAME, uploaded_file)
            st.success("✅ หยอดตามลำดับบรรทัดเรียบร้อย!")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="Sequential_Result.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
