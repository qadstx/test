import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v3", layout="wide")

TEMPLATE_FILENAME = "form_template.xlsx" 

def clean_string(s):
    if pd.isna(s): return ""
    # ดึงเฉพาะตัวอักษรและตัวเลข และทำเป็นตัวพิมพ์เล็ก
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).lower()

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    
    # อ่านไฟล์ข้อมูลแถวที่ 11 เป็นต้นไป
    df_source = pd.read_excel(data_file, skiprows=10)
    
    # ดึงข้อมูลจากไฟล์ Source เก็บไว้เป็นรายการทั้งหมด
    # โครงสร้าง: [ {'id': 'b84a', 'cav1': [...], 'cav2': [...]}, ... ]
    all_source_data = []
    for _, row in df_source.iterrows():
        item_raw = row.iloc[0]
        if pd.isna(item_raw): continue
        
        all_source_data.append({
            'clean_id': clean_string(item_raw),
            'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],   # G, H, I
            'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]] # L, M, N
        })

    target_sheets = ["Cav 1", "Cav 2"]
    
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_item_key = None
        matched_data_for_item = []
        item_usage_idx = 0

        # ไล่ตั้งแต่แถวที่ 8 ลงไป
        for r in range(8, ws.max_row + 1):
            cell_a_value = ws.cell(row=r, column=1).value
            
            # เมื่อเจอชื่อ Item ใหม่ในคอลัมน์ A (ช่องที่ Merge ไว้)
            if cell_a_value is not None:
                current_item_key = clean_string(cell_a_value)
                # ค้นหาข้อมูลใน all_source_data ที่ "ขึ้นต้นด้วย" current_item_key
                # เช่น ถ้าฟอร์มคือ 'b84' จะแมปกับ 'b84a', 'b84b' ในไฟล์ข้อมูล
                matched_data_for_item = [
                    d[mode] for d in all_source_data 
                    if d['clean_id'].startswith(current_item_key)
                ]
                item_usage_idx = 0 # รีเซ็ตตัวนับการใช้ข้อมูลสำหรับ Item นี้

            # ตรวจสอบคอลัมน์ L ว่าเป็น "Supplier" หรือไม่ (ช่องสีเหลือง)
            correlation_val = str(ws.cell(row=r, column=12).value).strip().lower()
            
            if "supplier" in correlation_val and matched_data_for_item:
                # ถ้ายังมีข้อมูลที่แมปไว้เหลืออยู่
                if item_usage_idx < len(matched_data_for_item):
                    vals = matched_data_for_item[item_usage_idx]
                    
                    # กรอกลงช่อง M, N, O
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                    
                    item_usage_idx += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ส่วนเดิม ---
st.header("📋 Auto-Fill v3 (Fixed B84/B85 Mapping)")

if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}`")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source .xlsx)", type=["xlsx"])
    if uploaded_file:
        if st.button("🚀 เริ่มกรอกข้อมูล"):
            try:
                result_xlsx = process_data(TEMPLATE_FILENAME, uploaded_file)
                st.success("✅ แก้ไขการแมปชื่อ B84, B85 และกรอกข้อมูลเรียบร้อย!")
                st.download_button(label="📥 ดาวน์โหลดไฟล์", data=result_xlsx, file_name="Fixed_Mapping_Result.xlsx")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
