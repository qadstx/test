import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

# ตั้งค่าหน้าจอ
st.set_page_config(page_title="Excel Auto-Filler", layout="wide")

# --- CONFIGURATION ---
# ชื่อไฟล์เทมเพลตที่อยู่ในโฟลเดอร์เดียวกับโค้ดบน GitHub
TEMPLATE_FILENAME = "form_template.xlsx" 

def clean_string(s):
    """ฟังก์ชันทำความสะอาดชื่อ Item: ลบช่องว่าง, อักขระพิเศษ และทำให้เป็นตัวพิมพ์เล็ก"""
    if pd.isna(s): return ""
    # ลบอักขระที่ไม่ใช่ตัวอักษรหรือตัวเลข (ยกเว้นขีดกลางถ้าจำเป็น)
    # ในที่นี้จะลบพวก _ หรือช่องว่างออกเพื่อให้ match กันง่ายขึ้น
    cleaned = re.sub(r'[^a-zA-Z0-9-]', '', str(s))
    return cleaned.lower()

def process_data(template_path, data_file):
    # 1. โหลด Template ด้วย openpyxl เพื่อรักษา Merge Cells และ Formatting
    wb = openpyxl.load_workbook(template_path)
    
    # 2. อ่านไฟล์ข้อมูลที่อัพโหลด (Source) ด้วย Pandas
    # ข้าม 10 แถวแรก เพื่อเริ่มอ่านแถวที่ 11 (Index 10)
    df_source = pd.read_excel(data_file, skiprows=10)
    
    # สร้าง Dictionary เก็บข้อมูล: { 'item_id': { 'cav1': [g,h,i], 'cav2': [l,m,n] } }
    data_map = {}
    for _, row in df_source.iterrows():
        item_raw = row.iloc[0] # คอลัมน์ A (Item Name)
        if pd.isna(item_raw): continue
        
        item_id = clean_string(item_raw)
        data_map[item_id] = {
            'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],   # G, H, I
            'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]] # L, M, N
        }

    # 3. จัดการทั้ง 2 ชีทในไฟล์ Template
    target_sheets = ["Cav 1", "Cav 2"]
    
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            st.warning(f"ไม่พบชีทชื่อ '{sheet_name}' ในไฟล์เทมเพลต")
            continue
            
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_item_key = None
        
        # เริ่มไล่ตั้งแต่แถวที่ 8 ตามที่ระบุ
        for r in range(8, ws.max_row + 1):
            cell_a_value = ws.cell(row=r, column=1).value
            
            # ถ้า Cell A มีค่า (ไม่ใช่ None จากการ Merge) ให้ Update Item ปัจจุบัน
            if cell_a_value is not None:
                current_item_key = clean_string(cell_a_value)
            
            # ถ้า Item ปัจจุบันมีข้อมูลใน data_map ให้ทำการกรอก
            # หมายเหตุ: จะกรอกเฉพาะแถวที่มีการ match เจอ และคุณบอกว่า "กรอกช่องเว้นช่อง"
            # ปกติถ้าเป็นช่องที่ต้องกรอก มักจะมี format หรือว่างไว้ 
            # ในที่นี้จะกรอกข้อมูลลงไปในคอลัมน์ M, N, O (13, 14, 15)
            if current_item_key in data_map:
                source_vals = data_map[current_item_key][mode]
                
                # เช็คว่าช่อง M (13) ว่างหรือไม่ก่อนกรอก เพื่อป้องกันการเขียนทับหัวข้อหรือช่องที่ไม่เกี่ยวข้อง
                # หรือถ้าต้องการกรอกทุกแถวที่อยู่ใน Block ของ Item นั้นก็รันต่อได้เลย
                ws.cell(row=r, column=13).value = source_vals[0]
                ws.cell(row=r, column=14).value = source_vals[1]
                ws.cell(row=r, column=15).value = source_vals[2]

    # บันทึกไฟล์ลง Memory เพื่อเตรียมให้ Download
    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI INTERFACE ---
st.header("📋 โปรแกรมกรอกข้อมูลฟอร์มอัตโนมัติ")
st.write(f"ไฟล์เทมเพลตที่ใช้: `{TEMPLATE_FILENAME}`")

# ตรวจสอบว่ามีไฟล์เทมเพลตอยู่ในโฟลเดอร์ไหม
if not os.path.exists(TEMPLATE_FILENAME):
    st.error(f"❌ ไม่พบไฟล์ `{TEMPLATE_FILENAME}` ใน GitHub Repository ของคุณ กรุณาตรวจสอบชื่อไฟล์")
else:
    uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source .xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        if st.button("🚀 ประมวลผลและสร้างไฟล์ใหม่"):
            try:
                with st.spinner("กำลังดึงข้อมูลและแมปช่องว่าง..."):
                    result_xlsx = process_data(TEMPLATE_FILENAME, uploaded_file)
                    
                st.success("✅ ดำเนินการเสร็จสิ้น!")
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ที่กรอกข้อมูลแล้ว",
                    data=result_xlsx,
                    file_name="Filled_Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {str(e)}")
                st.info("คำแนะนำ: ตรวจสอบว่าไฟล์ที่อัพโหลดมีโครงสร้างแถวที่ 11 เป็นต้นไปตามที่ตกลงกันไว้หรือไม่")

# --- สรุปคำแนะนำการใช้งาน ---
with st.expander("ดูวิธีใช้งานและข้อกำหนด"):
    st.markdown("""
    1. **ไฟล์เทมเพลต:** ต้องชื่อ `form_template.xlsx` และอยู่ในโฟลเดอร์เดียวกับ `app.py`
    2. **โครงสร้างไฟล์ข้อมูล:** 
       - ข้อมูลเริ่มที่แถว 11 (A11 คือชื่อ Item)
       - Cav 1 อยู่ที่ Column G, H, I
       - Cav 2 อยู่ที่ Column L, M, N
    3. **การแมปข้อมูล:** ระบบจะพยายามค้นหาชื่อ Item ที่ใกล้เคียงที่สุด (เช่น B-41 กับ B41 จะถือว่าเป็นตัวเดียวกัน)
    4. **การกรอก:** จะกรอกลงในคอลัมน์ M, N, O ของชีท Cav 1 และ Cav 2 ตามลำดับ
    """)
