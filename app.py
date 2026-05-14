import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v2", layout="wide")

TEMPLATE_FILENAME = "form_template.xlsx" 

def clean_string(s):
    if pd.isna(s): return ""
    # ดึงเฉพาะตัวอักษรและตัวเลข เช่น a_B-41 -> b41
    cleaned = re.sub(r'[^a-zA-Z0-9]', '', str(s))
    return cleaned.lower()

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    
    # อ่านไฟล์ข้อมูลแถวที่ 11 เป็นต้นไป
    df_source = pd.read_excel(data_file, skiprows=10)
    
    # เก็บข้อมูลเป็น list ของแถว เพื่อรองรับกรณี 1 Item มีหลายค่า (เช่น B81-1, B81-2)
    # โครงสร้าง: { 'b81': [ [g11,h11,i11], [g12,h12,i12], ... ] }
    data_map = {}
    for _, row in df_source.iterrows():
        item_raw = row.iloc[0]
        if pd.isna(item_raw): continue
        
        # ตัดตัวเลขพ่วงท้ายออกเพื่อให้ match กับหัวข้อหลัก (เช่น B81-1 -> b81)
        # ถ้าไฟล์อัพโหลดไม่มี -1 ต่อท้ายอยู่แล้ว logic นี้จะไม่กระทบครับ
        base_item_id = clean_string(str(item_raw).split('-')[0])
        
        if base_item_id not in data_map:
            data_map[base_item_id] = []
        
        data_map[base_item_id].append({
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
        item_usage_counter = {} # เอาไว้ใช้นับว่า Item นี้ถูกกรอกไปกี่แถวแล้ว (สำหรับกรณี B81)

        # ไล่ตั้งแต่แถวที่ 8 ลงไป
        for r in range(8, ws.max_row + 1):
            # 1. อัปเดต Item ID ปัจจุบัน (จากคอลัมน์ A ที่ Merge อยู่)
            cell_a_value = ws.cell(row=r, column=1).value
            if cell_a_value is not None:
                current_item_key = clean_string(cell_a_value)
                if current_item_key not in item_usage_counter:
                    item_usage_counter[current_item_key] = 0

            # 2. ตรวจสอบคอลัมน์ L (Correlation) ต้องเป็น "Supplier" เท่านั้นถึงจะกรอก
            correlation_val = str(ws.cell(row=r, column=12).value).strip().lower()
            
            if "supplier" in correlation_val and current_item_key in data_map:
                # ดึงข้อมูลลำดับที่ต้องการกรอก
                idx = item_usage_counter[current_item_key]
                
                # ตรวจสอบว่ายังมีข้อมูลใน list ให้กรอกอยู่ไหม (กัน Error กรณีแถวในฟอร์มเยอะกว่าข้อมูลที่มี)
                if idx < len(data_map[current_item_key]):
                    source_vals = data_map[current_item_key][idx][mode]
                    
                    # กรอกข้อมูลในคอลัมน์ M, N, O (13, 14, 15) ซึ่งเป็นช่องสีเหลือง
                    ws.cell(row=r, column=13).value = source_vals[0]
                    ws.cell(row=r, column=14).value = source_vals[1]
                    ws.cell(row=r, column=15).value = source_vals[2]
                    
                    # เพิ่มตัวนับ เพื่อให้แถว Supplier ถัดไปของ Item เดิม ไปดึงข้อมูลลำดับถัดไปมาใส่
                    item_usage_counter[current_item_key] += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Fill เฉพาะช่อง Supplier (สีเหลือง)")

if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}`")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source .xlsx)", type=["xlsx"])

    if uploaded_file:
        if st.button("🚀 เริ่มกรอกข้อมูล"):
            try:
                result_xlsx = process_data(TEMPLATE_FILENAME, uploaded_file)
                st.success("✅ กรอกข้อมูลเฉพาะช่อง Supplier เรียบร้อยแล้ว!")
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์",
                    data=result_xlsx,
                    file_name="Filled_Form_SupplierOnly.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
