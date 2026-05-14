import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v13", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def get_base_id(s):
    """
    ทำให้ทุกรูปแบบกลายเป็นมาตรฐานเดียวกัน:
    'a_B-41' -> 'B41'
    'b_B61'  -> 'B61'
    'T-1'    -> 'T1'
    'B186-1' -> 'B186'
    """
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    # หาตัวอักษร A, B หรือ T แล้วตามด้วยตัวเลข (ข้ามขีดกลางถ้ามี)
    # เช่น เจอ B ตามด้วย - แล้วตามด้วย 41 ให้ดึงแค่ B41
    match = re.search(r'([ABT])-?(\d+)', text)
    if match:
        return f"{match.group(1)}{match.group(2)}"
    return text

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    df_raw = pd.read_excel(data_file, header=None)
    
    # 1. ค้นหาพิกัด Cav 1 (เลข 1 ตัวแรก) และ Cav 2 (เลข 1 ตัวที่สอง)
    cav1_idx, cav2_idx = None, None
    for _, row in df_raw.iterrows():
        indices = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(indices) >= 2:
            cav1_idx, cav2_idx = indices[0], indices[1]
            break
    if cav1_idx is None: cav1_idx, cav2_idx = 8, 16 # ค่า Default ถ้าหาไม่เจอ

    # 2. จัดเก็บข้อมูลจากไฟล์ Source เรียงตามบรรทัด
    source_master = {}
    for _, row in df_raw.iterrows():
        raw_name = str(row.iloc[0]).strip()
        # กรองเฉพาะแถวที่มีรหัสหลัก A, B, T
        if not re.search(r'[ABT]', raw_name): continue
        
        base_id = get_base_id(raw_name)
        if not base_id: continue

        if base_id not in source_master:
            source_master[base_id] = []
            
        c1 = [row.iloc[cav1_idx], row.iloc[cav1_idx+1], row.iloc[cav1_idx+2]]
        c2 = [row.iloc[cav2_idx], row.iloc[cav2_idx+1], row.iloc[cav2_idx+2]]
        source_master[base_id].append({'cav1': c1, 'cav2': c2})

    # 3. หยอดลง Template
    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        current_id = None
        current_ptr = 0

        for r in range(1, ws.max_row + 1):
            cell_a = ws.cell(row=r, column=1).value
            if cell_a is not None:
                new_id = get_base_id(cell_a)
                if new_id != current_id:
                    current_id = new_id
                    current_ptr = 0 # เริ่มใช้ข้อมูลตัวแรกสุดของ Item ใหม่
            
            # ตรวจสอบช่อง Supplier (Column L = 12)
            l_val = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in l_val and current_id in source_master:
                data_list = source_master[current_id]
                if current_ptr < len(data_list):
                    vals = data_list[current_ptr][mode]
                    # หยอดลง M, N, O (13, 14, 15)
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                    current_ptr += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Filler v13 (B-41 & B61 Support)")
st.info("แก้ไขการอ่านรหัสที่มีขีดกลาง (B-xx) และ Prefix (a_, b_) ให้จับคู่กับฟอร์มได้แม่นยำ")

uploaded_file = st.file_uploader("อัพโหลดไฟล์ข้อมูล (Source)", type=["xlsx"])
if uploaded_file:
    if st.button("🚀 ประมวลผลข้อมูล"):
        try:
            result = process_data(TEMPLATE_FILENAME, uploaded_file)
            st.success("✅ อ่านรหัส B41-B80 และหยอดข้อมูลเรียบร้อย!")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="Complete_Fill_V13.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
