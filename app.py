import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v8", layout="wide")

TEMPLATE_FILENAME = "form_template.xlsx" 

def get_base_id(s):
    """แยก ID หลัก เช่น 'a_B177 (L)-1' -> 'b177'"""
    if pd.isna(s): return ""
    text = str(s).upper().strip()
    match = re.search(r'[AB]\d+', text)
    return match.group(0).lower() if match else text.lower()

def natural_sort_key(s):
    """ฟังก์ชันช่วยเรียงลำดับให้ 1-2 มาก่อน 1-10"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    df_source = pd.read_excel(data_file, skiprows=10)
    
    # 1. จัดระเบียบข้อมูลและทำ Natural Sort
    source_master = {}
    for _, row in df_source.iterrows():
        raw_name = row.iloc[0]
        if pd.isna(raw_name): continue
        
        base_id = get_base_id(raw_name)
        if base_id not in source_master:
            source_master[base_id] = []
            
        source_master[base_id].append({
            'name': str(raw_name),
            'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],
            'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]]
        })

    # สั่งเรียงลำดับแต่ละกลุ่ม Item ตามตัวเลข (ป้องกัน A1-10 มาก่อน A1-2)
    for bid in source_master:
        source_master[bid].sort(key=lambda x: natural_sort_key(x['name']))

    target_sheets = ["Cav 1", "Cav 2"]
    for sheet_name in target_sheets:
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
            
            correlation = str(ws.cell(row=r, column=12).value).strip().lower()
            
            if "supplier" in correlation and current_id in source_master:
                data_list = source_master[current_id]
                
                # กรอกเฉพาะเท่าที่มีช่องว่าง (เช่น มี 13 เอาแค่ 10 อันแรกที่เรียงแล้ว)
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
st.header("📋 Auto-Fill v8 (Natural Sort)")
st.info("แก้ไขการเรียงลำดับ A1-1 ถึง A1-10 ให้ตรงตามตำแหน่งช่องในฟอร์ม")

if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}`")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source .xlsx)", type=["xlsx"])
    if uploaded_file:
        if st.button("🚀 รันการประมวลผล"):
            try:
                result_xlsx = process_data(TEMPLATE_FILENAME, uploaded_file)
                st.success("✅ เรียงลำดับตัวเลขและกรอกข้อมูลเรียบร้อย!")
                st.download_button(label="📥 ดาวน์โหลดไฟล์", data=result_xlsx, file_name="Natural_Sorted_Result.xlsx")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
