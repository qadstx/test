import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO

st.set_page_config(page_title="Data Auto-Filler", layout="wide")

st.title("📊 Auto-Fill Form System (Cav 1 & Cav 2)")
st.info("อัพโหลดไฟล์ข้อมูลเพื่อนำไปเติมในฟอร์ม Template ที่เตรียมไว้บน GitHub")

# --- Helper Functions ---
def clean_item_name(name):
    """ฟังก์ชันดึงส่วนสำคัญของชื่อ Item เช่น a_B-41 -> B41"""
    if pd.isna(name): return ""
    name = str(name)
    # ลบอักขระพิเศษเหลือแค่ตัวอักษร ตัวเลข และเครื่องหมายลบ (ถ้ามี)
    # หรือปรับเปลี่ยนตาม Logic ที่ต้องการเพื่อให้ match กันได้
    # ในที่นี้จะลองดึง pattern ที่เป็น B ตามด้วยตัวเลข
    match = re.search(r'[Bb]\d+([-]\d+)?', name)
    return match.group(0).upper() if match else name.strip()

def process_excel(template_file, data_file):
    # โหลด Template ด้วย openpyxl เพื่อรักษา Format และ Merge Cells
    wb = openpyxl.load_workbook(template_file)
    
    # อ่านไฟล์ข้อมูล (Source) ด้วย Pandas
    # คอลัมน์ A11 (Index 0), G,H,I (Cav1), L,M,N (Cav2)
    # หมายเหตุ: ใน Pandas index เริ่มที่ 0 ดังนั้น A11 คือแถวที่ 10
    df_data = pd.read_excel(data_file, skiprows=10) # ข้ามไปที่บรรทัด 11
    
    # สมมติชื่อคอลัมน์ตามตำแหน่งที่แจ้ง (หรือใช้ index)
    # A=0, G=6, H=7, I=8, L=11, M=12, N=13
    data_map = {}
    for _, row in df_data.iterrows():
        raw_item = str(row.iloc[0])
        clean_key = clean_item_name(raw_item)
        data_map[clean_key] = {
            'cav1': [row.iloc[6], row.iloc[7], row.iloc[8]],
            'cav2': [row.iloc[11], row.iloc[12], row.iloc[13]]
        }

    # รายชื่อชีทที่ต้องจัดการ
    target_sheets = ["Cav 1", "Cav 2"]
    
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        data_key = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        
        # เริ่มไล่จากแถวที่ 8 (A8) ลงไปตามที่แจ้ง
        # เราจะเช็คคอลัมน์ A เพื่อหา Item No.
        current_item_name = None
        
        for row_idx in range(8, ws.max_row + 1):
            cell_a = ws.cell(row=row_idx, column=1)
            
            # ถ้าเป็นช่อง Merge ค่าจะอยู่ที่ Cell แรก, ช่องถัดไปจะเป็น None
            if cell_a.value:
                current_item_name = clean_item_name(cell_a.value)
            
            # ถ้าเจอ Item ที่ตรงกับในไฟล์อัพโหลด
            if current_item_name in data_map:
                vals = data_map[current_item_name][data_key]
                
                # กรอกข้อมูลที่ M, N, O (คอลัมน์ 13, 14, 15)
                # เช็คเงื่อนไข "ช่องเว้นช่อง" 
                # ในที่นี้สมมติว่าถ้าเป็นแถวที่จะกรอก จะมี Format บางอย่าง 
                # หรือกรอกทุกแถวที่อยู่ในขอบเขตของ Item นั้นแต่ตรวจสอบช่องว่าง
                ws.cell(row=row_idx, column=13).value = vals[0]
                ws.cell(row=row_idx, column=14).value = vals[1]
                ws.cell(row=row_idx, column=15).value = vals[2]

    # บันทึกลง Memory
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- UI ส่วนการจัดการไฟล์ ---
# ในการใช้งานจริง Template ควรจะดึงจาก GitHub โดยตรง (ต้องใช้ raw link)
template_url = "https://raw.githubusercontent.com/YOUR_USER/YOUR_REPO/main/template.xlsx"

uploaded_data = st.file_uploader("1. อัพโหลดไฟล์ข้อมูล (Source File)", type=["xlsx"])
uploaded_template = st.file_uploader("2. หรืออัพโหลดไฟล์ฟอร์ม (ถ้าไม่ใช้ตัวบน GitHub)", type=["xlsx"])

if st.button("เริ่มทำการ Fill ข้อมูล"):
    if uploaded_data and (uploaded_template or template_url):
        with st.spinner("กำลังประมวลผล..."):
            try:
                # ใช้ไฟล์ที่อัพโหลด หรือถ้าไม่มีให้ไปดึงจาก GitHub/ที่ระบุ
                temp_file = uploaded_template if uploaded_template else uploaded_template # ปรับแก้ตามสะดวก
                
                final_file = process_excel(temp_file, uploaded_data)
                
                st.success("เสร็จเรียบร้อย!")
                st.download_button(
                    label="ดาวน์โหลดไฟล์ที่เติมข้อมูลแล้ว",
                    data=final_file,
                    file_name="Filled_Form.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")
    else:
        st.warning("กรุณาอัพโหลดไฟล์ให้ครบถ้วน")
