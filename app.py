import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v5", layout="wide")

TEMPLATE_FILENAME = "form_template.xlsx" 

def intensive_clean(s):
    if pd.isna(s): return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).lower()

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    
    # อ่านไฟล์ข้อมูล (Source) เริ่มที่แถว 11
    df_source = pd.read_excel(data_file, skiprows=10)
    
    # ดึงข้อมูลทั้งหมดมาเก็บใน List และทำความสะอาด ID
    all_source_data = []
    for _, row in df_source.iterrows():
        item_raw = row.iloc[0]
        if pd.isna(item_raw): continue
        
        all_source_data.append({
            'raw_id': str(item_raw).strip(),
            'clean_id': intensive_clean(item_raw),
            'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],   # G, H, I
            'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]] # L, M, N
        })

    # เรียงลำดับข้อมูลใน Source ตามชื่อ Item (สำคัญมาก เพื่อให้กรอกเรียงแถวตรงกัน)
    all_source_data.sort(key=lambda x: x['raw_id'])

    target_sheets = ["Cav 1", "Cav 2"]
    
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_item_key = None
        matched_data_list = []
        usage_idx = 0

        for r in range(8, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1)
            
            # เมื่อเจอหัวข้อ Item ใหม่ในฟอร์ม (เช่น B82)
            if cell_a.value is not None:
                current_item_key = intensive_clean(cell_a.value)
                
                # ค้นหาข้อมูลที่ "ขึ้นต้นด้วย" รหัสนี้เท่านั้น 
                # เช่น ถ้าฟอร์มคือ 'b82' ต้องเจอ 'b82-2', 'b82-3'
                # แต่จะไม่เอา 'b8' หรือ 'b820'
                matched_data_list = [
                    d[mode] for d in all_source_data 
                    if d['clean_id'].startswith(current_item_key)
                ]
                usage_idx = 0 

            # ตรวจสอบแถว Supplier (ช่องสีเหลือง)
            correlation = str(ws.cell(row=r, column=12).value).strip().lower()
            
            if "supplier" in correlation and matched_data_list:
                if usage_idx < len(matched_data_list):
                    vals = matched_data_list[usage_idx]
                    
                    # กรอกข้อมูล (คอลัมน์ M, N, O)
                    # ใช้รอบคอบ: ถ้าค่าเป็น None ให้เว้นว่างไว้
                    ws.cell(row=r, column=13).value = vals[0] if vals[0] != 0 else ""
                    ws.cell(row=r, column=14).value = vals[1] if vals[1] != 0 else ""
                    ws.cell(row=r, column=15).value = vals[2] if vals[2] != 0 else ""
                    
                    usage_idx += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Fill v5 (Sorted & Strict Matching)")

if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}` ใน GitHub")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source .xlsx)", type=["xlsx"])
    if uploaded_file:
        if st.button("🚀 เริ่มกรอกข้อมูล (ลำดับใหม่)"):
            try:
                result_xlsx = process_data(TEMPLATE_FILENAME, uploaded_file)
                st.success("✅ แก้ไขลำดับการกรอกข้อมูลให้ตรงตามไฟล์ต้นฉบับแล้ว!")
                st.download_button(label="📥 ดาวน์โหลดไฟล์", data=result_xlsx, file_name="Sorted_Match_Result.xlsx")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
