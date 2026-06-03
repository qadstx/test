import streamlit as st
import pandas as pd
import io
import os
import glob
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

st.set_page_config(layout="wide")
st.title("🧪 Washing Date Auto-Fetch (Test App)")
st.write("แอปทดสอบดึงข้อมูลอัตโนมัติจากเว็บบริษัท โดยอัปโหลดเพียงไฟล์เดียว")

# ==========================================
# 🌐 FUNCTION: SELENIUM BOT
# ==========================================
def fetch_file2_from_web(start_lot, end_lot):
    # ตั้งค่าโฟลเดอร์สำหรับดาวน์โหลดชั่วคราว
    download_dir = os.path.join(os.getcwd(), "temp_downloads")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
        
    # เคลียร์ไฟล์เก่าออกก่อนรันรอบใหม่
    for f in glob.glob(os.path.join(download_dir, "*")):
        try: os.remove(f)
        except: pass

    chrome_options = webdriver.ChromeOptions()
    # 💡 เพื่อทดสอบในรอบแรก ผมขอเปิดหน้าจอเบราว์เซอร์ไว้ก่อนนะครับ คุณจะได้เห็นตอนบอทมันทำงาน
    # ถ้าใช้งานจริงแล้วอยากให้มันซ่อน ค่อยเอาเครื่องหมาย # ออกจากบรรทัดด้านล่างครับ
    # chrome_options.add_argument("--headless") 
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 30) # เพิ่มเวลารอเผื่อเน็ตหรือเว็บช้า
    actions = ActionChains(driver)

    try:
        # 1. เข้าหน้าเว็บและทำการล็อกอิน
        st.info("🌐 กำลังเปิดหน้าเว็บ...")
        driver.get("http://sms53.nokpct.com:88/MPCTSystem/View/Login/Default.aspx")
        
        st.info("🔑 กำลังกรอกข้อมูล Login...")
        # 💡 กรอก Username
        user_input = wait.until(EC.presence_of_element_located((By.ID, "txtSearchUser")))
        user_input.send_keys("2M260005")
        
        # 💡 กรอก Password
        pass_input = driver.find_element(By.ID, "txtSearchPass")
        pass_input.send_keys("56891")
        
        # 💡 กดปุ่ม Login
        login_btn = driver.find_element(By.ID, "imgBtnLogin")
        login_btn.click()
        time.sleep(3) 

        # 2. เลื่อนเมาส์ชี้เมนูย่อยลึก 4 ชั้น
        st.info("🖱️ กำลังนำทางไปเมนู Seagate Traceability...")
        
        menu_product = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(@class, 'rmLink')]//span[text()='Product']")))
        actions.move_to_element(menu_product).perform()
        time.sleep(0.5)
        
        menu_hdd = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(@class, 'rmLink')]//span[text()='HDD']")))
        actions.move_to_element(menu_hdd).perform()
        time.sleep(0.5)
        
        menu_seagate = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(@class, 'rmLink')]//span[text()='Seagate']")))
        actions.move_to_element(menu_seagate).perform()
        time.sleep(0.5)
        
        menu_trace = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(@class, 'rmLink')]//span[text()='Seagate Traceability']")))
        menu_trace.click()

        # 3. กรอกข้อมูลค้นหาในหน้าเงื่อนไข
        st.info("✍️ กำลังตั้งค่าเงื่อนไขการค้นหา...")
        
        # เลือก MPN
        mpn_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_cbMPN_Input")))
        mpn_input.click()
        time.sleep(0.3)
        mpn_input.send_keys(Keys.CONTROL + "a")
        mpn_input.send_keys(Keys.DELETE)
        mpn_input.send_keys("FHNGZ8427V00") 
        time.sleep(0.3)
        mpn_input.send_keys(Keys.ENTER)
        time.sleep(1)

        # กรอก Lot No From
        input_from = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtSearchLotNoFrom")))
        input_from.click()
        input_from.send_keys(Keys.CONTROL + "a")
        input_from.send_keys(Keys.DELETE)
        input_from.send_keys(start_lot)
        time.sleep(0.3)

        # กรอก Lot No To
        input_to = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_txtSearchLotNoTo")))
        input_to.click()
        input_to.send_keys(Keys.CONTROL + "a")
        input_to.send_keys(Keys.DELETE)
        input_to.send_keys(end_lot)
        time.sleep(0.5)

        # 4. กดปุ่ม Execute
        st.info("⚡ กำลังประมวลผลดึงข้อมูล (Execute)...")
        btn_execute = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'rtbWrap')]//span[text()='Execute']")))
        btn_execute.click()
        time.sleep(6) # รอเว็บโหลดข้อมูลเข้าตารางสักครู่

        # 5. กดปุ่ม Excel เพื่อดาวน์โหลด
        st.info("📥 กำลังดาวน์โหลดไฟล์ Excel จากหน้าเว็บ...")
        btn_excel = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'rtbWrap')]//span[text()='Excel']")))
        btn_excel.click()
        time.sleep(5) # รอไฟล์ดาวน์โหลดเสร็จสิ้น
        
        # 6. เช็กไฟล์ที่โหลดมาล่าสุด
        files = glob.glob(os.path.join(download_dir, "*"))
        if not files:
            st.error("❌ บอทดาวน์โหลดไฟล์ไม่สำเร็จ")
            return None
            
        latest_file = max(files, key=os.path.getctime)
        st.success(f"📂 ได้รับไฟล์เรียบร้อย: {os.path.basename(latest_file)}")
        
        if latest_file.endswith(('.xls', '.xlsx')):
            return pd.read_excel(latest_file)
        elif latest_file.endswith('.csv'):
            return pd.read_csv(latest_file)
            
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในสคริปต์บอท: {e}")
        return None
    finally:
        driver.quit()

# ==========================================
# 📊 APP STREAMLIT UI & READ FILE 1
# ==========================================
file1 = st.file_uploader("📂 Step 1: อัปโหลดไฟล์ที่ 1 (Lot/Serial)", type=["xls", "xlsx", "csv"])

if file1 is not None:
    # อ่านไฟล์แรก (อิงตาม Logic ตำแหน่งแถวและคอลัมน์เดิมของคุณ)
    try:
        df1_raw = pd.read_excel(file1, engine="openpyxl", header=None)
    except:
        df1_raw = pd.read_excel(file1, engine="xlrd", header=None)
        
    col = 5
    start_row = 16
    data = df1_raw.iloc[start_row:, col]

    lot_list = []
    for val in data:
        if pd.isna(val): break
        val_str = str(val).strip()
        if val_str == "": break
        lot_list.append(val_str)
        
    df1 = pd.DataFrame({"Lot": lot_list})
    
    if not df1.empty:
        start_lot = df1["Lot"].iloc[0]
        end_lot = df1["Lot"].iloc[-1]
        
        # แสดงข้อมูล Lot ที่ตรวจจับได้จากไฟล์แรก
        col1, col2, col3 = st.columns(3)
        col1.metric("จำนวน Lot ทั้งหมดในไฟล์", len(df1))
        col2.metric("Lot ตัวแรก (Start)", start_lot)
        col3.metric("Lot ตัวสุดท้าย (End)", end_lot)
        
        st.write("---")
        st.subheader("📋 ข้อมูลจากไฟล์แรกที่คุณอัปโหลด")
        st.dataframe(df1, height=200)
        
        # ปุ่มกดเพื่อเริ่มการทำงานของบอท
        if st.button("🚀 Step 2: สั่งบอทเปิดเว็บไปดาวน์โหลดไฟล์อัตโนมัติ", type="primary"):
            with st.spinner("🤖 บอทกำลังทำงาน ห้ามปิดหน้าต่างนี้จนกว่าจะเสร็จสิ้น..."):
                web_data = fetch_file2_from_web(start_lot, end_lot)
                
            if web_data is not None:
                st.success("🎉 ดึงข้อมูลสำเร็จ! ด้านล่างนี้คือข้อมูลจำลองจาก File 2 ที่บอทไปดึงมาให้")
                st.dataframe(web_data)
    else:
        st.error("❌ ไม่พบข้อมูล Lot ในไฟล์แรกที่อัปโหลด กรุณาตรวจสอบตำแหน่งแถว/คอลัมน์")
