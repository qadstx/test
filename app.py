import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v4", layout="wide")

TEMPLATE_FILENAME = "form_template.xlsx" 

def intensive_clean(s):
    """ลบทุกอย่างที่ไม่ใช่ตัวเลขและตัวอักษร และตัดเว้นวรรคทิ้งทั้งหมด"""
    if pd.isna(s): return ""
    text = str(s).strip()
    # ลบอักขระพิเศษและช่องว่างทั้งหมด
    cleaned = re.sub(r'[^a-zA-Z0-9]', '', text)
    return cleaned.lower()

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    
    # อ่านไฟล์ข้อมูลแถวที่ 11
    df_source = pd.read_excel(data_file, skiprows=10)
    
    all_source_data = []
    # เก็บ Log ไว้เช็ค
    source_items_found = []

    for _, row in df_source.iterrows():
        item_raw = row.iloc[0]
        if pd.isna(item_raw): continue
        
        clean_id = intensive_clean(item_raw)
        source_items_found.append(clean_id)
        
        all_source_data.append({
            'clean_id': clean_id,
            'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],   # G, H, I
            'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]] # L, M, N
        })

    target_sheets = ["Cav 1", "Cav 2"]
    log_messages = []

    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_item_key = None
        matched_data_for_item = []
        item_usage_idx = 0

        for r in range(8, ws.max_row + 1):
            cell_a_value = ws.cell(row=r, column=1).value
            
            if cell_a_value is not None:
                current_item_key = intensive_clean(cell_a_value)
                # ค้นหาข้อมูล: ใช้ regex เช็คว่าชื่อใน source มีชื่อใน template อยู่ข้างในหรือไม่
                # เช่น ถ้า template คือ 'b86' ให้เจอ 'b86', 'b86a', 'itemb86'
                matched_data_for_item = [
                    d[mode] for d in all_source_data 
                    if current_item_key in d['clean_id']
                ]
                item_usage_idx = 0
                
                if matched_data_for_item:
                    log_messages.append(f"Sheet {sheet_name}: Found {len(matched_data_for_item)} matches for Item {cell_a_value}")

            # เช็คคำว่า Supplier ในคอลัมน์ L
            correlation_val = str(ws.cell(row=r, column=12).value).strip().lower()
            
            if "supplier" in correlation_val and matched_data_for_item:
                if item_usage_idx < len(matched_data_for_item):
                    vals = matched_data_for_item[item_usage_idx]
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                    item_usage_idx += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue(), log_messages

# --- UI ---
st.header("📋 Auto-Fill v4 (Intensive Mapping)")

if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}`")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source .xlsx)", type=["xlsx"])
    if uploaded_file:
        if st.button("🚀 เริ่มกรอกข้อมูล"):
            try:
                result_xlsx, logs = process_data(TEMPLATE_FILENAME, uploaded_file)
                st.success("✅ ประมวลผลเสร็จสิ้น!")
                
                with st.expander("ดูรายการ Item ที่ระบบตรวจเจอ"):
                    for msg in logs[:20]: # แสดงแค่ 20 อันแรก
                        st.write(msg)
                    if len(logs) > 20: st.write("...")

                st.download_button(label="📥 ดาวน์โหลดไฟล์", data=result_xlsx, file_name="Full_Fill_Result.xlsx")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
