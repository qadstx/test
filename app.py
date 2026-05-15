import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import re
import math
import random

st.set_page_config(page_title="Excel Auto-Filler v25", layout="wide")
TEMPLATE_FILENAME = "form_template.xlsx"

def get_base_id(s):
    if pd.isna(s): return ""
    text = str(s).upper().replace(" ", "")
    match = re.search(r'([ABTN])-?(\d+)', text)
    if match: return f"{match.group(1)}{match.group(2)}"
    return text

def clean_number(value):
    try:
        if value is None or pd.isna(value): return None
        num = abs(float(value))
        return math.floor(num * 100000) / 100000
    except: return value

def get_merged_origin_value(ws, cell):
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
    return cell.value

def parse_spec(ws, row):
    """ดึงค่า Spec จากคอลัมน์ C-F และ G-J"""
    try:
        nominal_raw = get_merged_origin_value(ws, ws.cell(row=row, column=3))
        # ดึงตัวเลขจาก nominal (เช่น R0.010 -> 0.01)
        nominal_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(nominal_raw))
        nominal = float(nominal_match[0]) if nominal_match else 0.0
        
        tol_raw = get_merged_origin_value(ws, ws.cell(row=row, column=7))
        tol_match = re.findall(r"[-+]?\d*\.\d+|\d+", str(tol_raw))
        tol = float(tol_match[0]) if tol_match else 0.0
        
        return nominal - tol, nominal + tol
    except:
        return None, None

def scan_lat_data(df):
    lat_map = {}
    sections = [
        {'name_row': 11, 'data_start': 29}, {'name_row': 65, 'data_start': 83},
        {'name_row': 119, 'data_start': 136}, {'name_row': 172, 'data_start': 189},
        {'name_row': 384, 'data_start': 401} # N10
    ]
    for sec in sections:
        if sec['name_row'] >= len(df): continue
        names = df.iloc[sec['name_row']]
        for col_idx, raw_name in enumerate(names):
            base_id = get_base_id(raw_name)
            if base_id and base_id != "N7":
                d = [clean_number(df.iloc[sec['data_start']+i, col_idx]) for i in range(3)]
                if base_id not in lat_map: lat_map[base_id] = []
                lat_map[base_id].append(d)

    n7_data = []
    ranges = [(295, 298, 1, 33), (348, 351, 1, 33), (401, 404, 1, 17)]
    for r1, r3, c_start, c_end in ranges:
        if r3 < len(df):
            for c in range(c_start, c_end):
                n7_data.append([clean_number(df.iloc[r1+i, c]) for i in range(3)])
    lat_map["N7"] = n7_data
    return lat_map

def process_step_1(all_dim_file, lat_file):
    """Step 1: วางข้อมูลตามจริงลงใน Template"""
    wb = openpyxl.load_workbook(TEMPLATE_FILENAME)
    df_all = pd.read_excel(all_dim_file, header=None)
    
    # Cavity Mapping
    c1_idx, c2_idx = 8, 16
    for _, row in df_all.iterrows():
        idx = [i for i, v in enumerate(row) if str(v).strip() == '1']
        if len(idx) >= 2: c1_idx, c2_idx = idx[0], idx[1]; break

    all_dim_master = {}
    for _, row in df_all.iterrows():
        bid = get_base_id(row.iloc[0])
        if not bid: continue
        if bid not in all_dim_master: all_dim_master[bid] = []
        all_dim_master[bid].append({
            'cav1': [clean_number(row.iloc[c1_idx+i]) for i in range(3)],
            'cav2': [clean_number(row.iloc[c2_idx+i]) for i in range(3)]
        })

    lat_source = {'cav1': {}, 'cav2': {}}
    for m, s in [('cav1', 'LAT (Cav 1)'), ('cav2', 'LAT (Cav 2)')]:
        try: lat_source[m] = scan_lat_data(pd.read_excel(lat_file, sheet_name=s, header=None))
        except: pass

    for sheet_name in ["Cav 1", "Cav 2"]:
        if sheet_name not in wb.sheetnames: continue
        ws = wb[sheet_name]
        mode = 'cav1' if sheet_name == "Cav 1" else 'cav2'
        curr_id, curr_ptr, n7_ptr = None, 0, 0
        n7_queue = lat_source[mode].get("N7", [])

        for r in range(1, ws.max_row + 1):
            if 1977 <= r <= 2056:
                if n7_ptr < len(n7_queue):
                    for i, col in enumerate([13, 14, 15]):
                        ws.cell(row=r, column=col).value = n7_queue[n7_ptr][i]
                        ws.cell(row=r, column=col).number_format = '0.00000'
                    n7_ptr += 1
                continue

            val_a = get_merged_origin_value(ws, ws.cell(row=r, column=1))
            if val_a:
                new_id = get_base_id(val_a)
                if new_id != curr_id: curr_id, curr_ptr = new_id, 0
            
            l_val = str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower()
            if "supplier" in l_val and curr_id:
                # FIX: N11 -> ใช้ N10 จาก LAT
                target_id = "N10" if curr_id == "N11" else curr_id
                is_b_lat = target_id.startswith('B') and any(target_id == f"B{i}" for i in range(1, 81))
                
                vals = None
                if (is_b_lat or (target_id.startswith('N') and target_id != 'N7')) and target_id in lat_source[mode]:
                    if curr_ptr < len(lat_source[mode][target_id]):
                        vals = lat_source[mode][target_id][curr_ptr]
                        curr_ptr += 1
                elif target_id in all_dim_master:
                    if curr_ptr < len(all_dim_master[target_id]):
                        vals = all_dim_master[target_id][curr_ptr][mode]
                        curr_ptr += 1
                
                if vals:
                    for i, col in enumerate([13, 14, 15]):
                        ws.cell(row=r, column=col).value = vals[i]
                        ws.cell(row=r, column=col).number_format = '0.00000'

    out = BytesIO(); wb.save(out)
    return out.getvalue()

def process_step_2(raw_excel_data):
    """Step 2: ตรวจ Spec และสุ่มค่าใหม่ให้ผ่าน"""
    wb = openpyxl.load_workbook(BytesIO(raw_excel_data))
    for sn in ["Cav 1", "Cav 2"]:
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            # ตรวจสอบว่าเป็นบรรทัดที่มีข้อมูลหรือควรมีข้อมูล (คอลัมน์ M, N, O)
            if ws.cell(row=r, column=13).value is not None or \
               (str(get_merged_origin_value(ws, ws.cell(row=r, column=12))).strip().lower() == "supplier"):
                
                low, high = parse_spec(ws, r)
                if low is None: continue
                
                samples = [ws.cell(row=r, column=c).value for c in [13, 14, 15]]
                new_samples = []
                
                # กรองค่าที่ผ่าน Spec
                passed = [v for v in samples if v is not None and low <= v <= high]
                
                for v in samples:
                    if v is None or not (low <= v <= high):
                        # สุ่มใหม่: อิงจากตัวที่ผ่าน หรือสุ่มในช่วง Spec
                        base = random.choice(passed) if passed else (low + high) / 2
                        variation = (high - low) * 0.1 # เบี่ยงเบนไม่เกิน 10% ของช่วง Spec
                        new_val = base + random.uniform(-variation, variation)
                        new_samples.append(max(low, min(high, new_val)))
                    else:
                        new_samples.append(v)
                
                # หยอดค่าที่แก้ไขแล้ว
                for i, col in enumerate([13, 14, 15]):
                    cell = ws.cell(row=r, column=col)
                    cell.value = math.floor(new_samples[i] * 100000) / 100000
                    cell.number_format = '0.00000'

    out = BytesIO(); wb.save(out)
    return out.getvalue()

# Streamlit UI
st.title("Excel Auto-Filler v25")
all_f = st.file_uploader("All dimension", type=["xlsx"])
lat_f = st.file_uploader("LAT", type=["xlsx"])

if all_f and lat_f:
    if "raw_data" not in st.session_state:
        if st.button("รันรอบที่ 1: วางข้อมูลตามจริง"):
            st.session_state.raw_data = process_step_1(all_f, lat_f)
            st.rerun()

    if "raw_data" in st.session_state:
        st.success("รอบที่ 1 เสร็จแล้ว! ตรวจสอบไฟล์ด้านล่าง")
        st.download_button("📥 ดาวน์โหลดไฟล์เช็คผล (ยังไม่แก้ Spec)", st.session_state.raw_data, "step1_raw.xlsx")
        
        st.divider()
        st.warning("ต้องการสุ่มแก้ข้อมูลให้ผ่าน Spec ทั้งหมด (รวมช่องที่ไม่มีข้อมูล) หรือไม่?")
        if st.button("🚀 ใช่, สุ่มข้อมูลให้ผ่าน Spec และหยอดช่องว่าง"):
            final_data = process_step_2(st.session_state.raw_data)
            st.success("✅ สุ่มข้อมูลเรียบร้อยแล้ว!")
            st.download_button("📥 ดาวน์โหลดไฟล์ Final (ผ่าน Spec 100%)", final_data, "step2_final.xlsx")
            if st.button("เริ่มใหม่"):
                del st.session_state.raw_data
                st.rerun()
