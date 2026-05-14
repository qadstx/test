import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v6", layout="wide")

TEMPLATE_FILENAME = "form_template.xlsx" 

def intensive_clean(s):
    if pd.isna(s): return ""
    # ลบอักขระพิเศษและทำเป็นตัวพิมพ์เล็ก
    return re.sub(r'[^a-zA-Z0-9]', '', str(s)).lower()

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    
    # 1. เตรียมข้อมูลจากไฟล์ Source (ขวา)
    df_source = pd.read_excel(data_file, skiprows=10)
    source_dict = {} # เก็บข้อมูลแยกตาม ID
    
    for _, row in df_source.iterrows():
        item_raw = row.iloc[0]
        if pd.isna(item_raw): continue
        
        clean_id = intensive_clean(item_raw)
        # เราจะพยายามเดา ID หลัก เช่น b82-2 -> b82 เพื่อให้ match กับฟอร์ม
        # ใช้ regex ตัดเลขพ่วงท้ายออกเพื่อให้ได้รหัสหลัก
        base_id = re.sub(r'\d+$', '', clean_id) if '-' in str(item_raw) else clean_id
        # แต่ถ้าเป็นพวก B81, B82 เฉยๆ ก็ใช้ค่านั้นเลย
        
        if clean_id not in source_dict:
            source_dict[clean_id] = {
                'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],
                'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]]
            }

    # สร้าง List ของ IDs ที่เรียงลำดับแล้ว
    sorted_source_keys = sorted(source_dict.keys())

    # 2. จัดการไฟล์ Template (ซ้าย)
    target_sheets = ["Cav 1", "Cav 2"]
    
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_item_in_form = None
        # เก็บข้อมูลที่ Match ได้สำหรับ Item ปัจจุบัน
        matched_items_data = []
        idx_to_fill = 0

        # วนลูปแถวใน Template
        for r in range(8, ws.max_row + 1):
            cell_a_val = ws.cell(row=r, column=1).value
            
            # ถ้าเจอชื่อ Item ใหม่ใน Column A
            if cell_a_val is not None:
                current_item_in_form = intensive_clean(cell_a_val)
                # ค้นหาข้อมูลจาก Source ที่เกี่ยวข้องกับรหัสนี้
                # เช่น ถ้าฟอร์มคือ b82 ให้เอา b82-2, b82-3 มาทั้งหมด
                matched_items_data = [
                    source_dict[k][mode] for k in sorted_source_keys 
                    if k.startswith(current_item_in_form)
                ]
                idx_to_fill = 0 # รีเซ็ตทุกครั้งที่ขึ้น Item ใหม่ **ป้องกันข้อมูลไหล**

            # ตรวจสอบว่าต้องกรอกไหม (ช่องสีเหลือง Supplier)
            correlation = str(ws.cell(row=r, column=12).value).strip().lower()
            
            if "supplier" in correlation:
                if matched_items_data and idx_to_fill < len(matched_items_data):
                    data_row = matched_items_data[idx_to_fill]
                    
                    # กรอกข้อมูล M, N, O
                    ws.cell(row=r, column=13).value = data_row[0]
                    ws.cell(row=r, column=14).value = data_row[1]
                    ws.cell(row=r, column=15).value = data_row[2]
                    
                    idx_to_fill += 1
                else:
                    # ถ้าไม่มีข้อมูลให้กรอก หรือข้อมูลหมดแล้ว ให้ปล่อยว่าง (สำคัญ!)
                    # วิธีนี้จะป้องกันไม่ให้ข้อมูลเก่าไหลลงไปแถวล่างๆ
                    pass

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Fill v6 (Anti-Drift Edition)")
st.info("แก้ไขปัญหาข้อมูลไหลลงไปกองด้านล่าง และการ Reset ตำแหน่ง Item")

if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}`")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source .xlsx)", type=["xlsx"])
    if uploaded_file:
        if st.button("🚀 รันการกรอกแบบแม่นยำตำแหน่ง"):
            try:
                result_xlsx = process_data(TEMPLATE_FILENAME, uploaded_file)
                st.success("✅ แก้ไขปัญหาการเยื้องของข้อมูลเรียบร้อยแล้ว!")
                st.download_button(label="📥 ดาวน์โหลดไฟล์ที่แก้ไขแล้ว", data=result_xlsx, file_name="Final_Corrected_Form.xlsx")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
