import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import re
import os

st.set_page_config(page_title="Excel Auto-Filler v14", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx" 

def get_base_id(s):
    """ล้างรหัสให้ตรงกัน: 'a_B-41' -> 'B41', 'T-1' -> 'T1'"""
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABT])-?(\d+)', text)
    if match:
        return f"{match.group(1)}{match.group(2)}"
    return text

def clean_number(value):
    """แปลงค่าให้เป็นตัวเลขบวกเท่านั้น และตัดเครื่องหมายลบทิ้ง"""
    try:
        if value is None or pd.isna(value):
            return None
        # แปลงเป็น float แล้วใช้ abs() เพื่อเอาเครื่องหมายลบออก
        num = float(value)
        return abs(num)
    except (ValueError, TypeError):
        # ถ้าไม่ใช่ตัวเลข (เช่น เป็นข้อความ) ให้ส่งกลับค่าเดิม
        return value

def process_data(template_path, data_file):
    wb = openpyxl.load_workbook(template_path)
    df_raw = pd.read_excel(data_file, header=None)
    
    # 1. หาพิกัดคอลัมน์ Cav 1 และ Cav 2
    cav1_idx, cav2_idx = None, None
    for _, row in df_raw.iterrows():
        indices = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(indices) >= 2:
            cav1_idx, cav2_idx = indices[0], indices[1]
            break
    if cav1_idx is None: cav1_idx, cav2_idx = 8, 16

    # 2. เก็บข้อมูลจาก Source (หยิบตามบรรทัดที่เจอ)
    source_master = {}
    for _, row in df_raw.iterrows():
        raw_name = str(row.iloc[0]).strip()
        if not re.search(r'[ABT]', raw_name): continue
        
        base_id = get_base_id(raw_name)
        if not base_id: continue

        if base_id not in source_master:
            source_master[base_id] = []
            
        # ดึงข้อมูลและทำให้เป็นค่าบวกทั้งหมด
        c1 = [clean_number(row.iloc[cav1_idx]), 
              clean_number(row.iloc[cav1_idx+1]), 
              clean_number(row.iloc[cav1_idx+2])]
        c2 = [clean_number(row.iloc[cav2_idx]), 
              clean_number(row.iloc[cav2_idx+1]), 
              clean_number(row.iloc[cav2_idx+2])]
        
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
                    current_ptr = 0
            
            l_val = str(ws.cell(row=r, column=12).value).strip().lower()
            if "supplier" in l_val and current_id in source_master:
                data_list = source_master[current_id]
                if current_ptr < len(data_list):
                    vals = data_list[current_ptr][mode]
                    # หยอดค่าบวกลงช่อง M, N, O
                    ws.cell(row=r, column=13).value = vals[0]
                    ws.cell(row=r, column=14).value = vals[1]
                    ws.cell(row=r, column=15).value = vals[2]
                    current_ptr += 1

    out_stream = BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

# --- UI ---
st.header("📋 Auto-Filler v14 (Absolute Value Fix)")
st.info("Logic: หยอดตามลำดับ + แปลงค่าติดลบให้เป็นบวกทั้งหมด (ตัดเครื่องหมาย - ทิ้ง)")

uploaded_file = st.file_uploader("อัปโหลดไฟล์ Source Data", type=["xlsx"])
if uploaded_file:
    if st.button("🚀 รันการประมวลผล (Positive Numbers Only)"):
        try:
            result = process_data(TEMPLATE_FILENAME, uploaded_file)
            st.success("✅ ประมวลผลสำเร็จ! ค่าติดลบถูกแปลงเป็นค่าบวกเรียบร้อยแล้ว")
            st.download_button("📥 ดาวน์โหลดไฟล์", result, file_name="Positive_Value_Result.xlsx")
        except Exception as e:
            st.error(f"Error: {e}")
