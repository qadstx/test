import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler Final", layout="wide")

TEMPLATE_FILENAME = "form_template.xlsx" 

def get_base_id(s):
    """ดึงรหัสหลักออกมา เช่น 'a_B177 (L)-1' -> 'b177'"""
    if pd.isna(s): return ""
    text = str(s).upper()
    # หาตัว B ตามด้วยตัวเลข
    match = re.search(r'B\d+', text)
    return match.group(0).lower() if match else text.lower()

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    df_source = pd.read_excel(data_file, skiprows=10)
    
    # เก็บข้อมูลเป็น List แยกตาม Base ID เพื่อให้ดึงตามลำดับได้
    # { 'b177': [[g,h,i], [g,h,i], ...], 'b178': [...] }
    source_master = {}
    for _, row in df_source.iterrows():
        raw_name = row.iloc[0]
        if pd.isna(raw_name): continue
        
        base_id = get_base_id(raw_name)
        if base_id not in source_master:
            source_master[base_id] = []
            
        source_master[base_id].append({
            'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],
            'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]]
        })

    target_sheets = ["Cav 1", "Cav 2"]
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_id = None
        fill_count = 0

        for r in range(8, ws.max_row + 1):
            # 1. เช็คว่าขึ้น Item ใหม่หรือยัง
            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                current_id = get_base_id(cell_a)
                fill_count = 0 # รีเซ็ตตัวนับข้อมูลเมื่อขึ้น Item ใหม่
            
            # 2. เช็คว่าเป็นช่องสีเหลือง (Supplier) หรือไม่
            correlation = str(ws.cell(row=r, column=12).value).strip().lower()
            
            if "supplier" in correlation and current_id in source_master:
                data_list = source_master[current_id]
                
                # ถ้ายังมีข้อมูลใน List ให้ดึงมาใส่ตามลำดับ 1, 2, 3...
                if fill_count < len(data_list):
                    vals = data_list[fill_count]
                    ws.cell(row=r, column=13).value = vals[mode][0]
                    ws.cell(row=r, column=14).value = vals[mode][1]
                    ws.cell(row=r, column=15).value = vals[mode][2]
                    fill_count += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 โปรแกรมกรอกข้อมูลฟอร์ม (เวอร์ชันแก้ทางเดินข้อมูล)")

if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}` บน GitHub")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูลล่าสุด", type=["xlsx"])
    if uploaded_file:
        if st.button("🚀 ประมวลผล (Strict Matching)"):
            try:
                result_xlsx = process_data(TEMPLATE_FILENAME, uploaded_file)
                st.success("✅ ดำเนินการเสร็จสิ้น! ข้อมูลถูกกรอกตามลำดับในแต่ละ Block แล้ว")
                st.download_button(label="📥 ดาวน์โหลดไฟล์", data=result_xlsx, file_name="Complete_Fill_Result.xlsx")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
