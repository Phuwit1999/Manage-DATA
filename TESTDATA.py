import streamlit as st
import pandas as pd

st.title("ระบบจัดการข้อมูลใบแจ้งหนี้")

# --- ต้องมีบรรทัดนี้ก่อนเพื่อสร้างตัวแปร uploaded_file ---
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ", type=['xlsx'])

# --- จากนั้นค่อยเช็คเงื่อนไข ---
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # 1. จัดการเลขที่ใบแจ้งหนี้ (ตัด .0 ออก)
    # แปลงเป็น numeric -> เติม 0 แทนค่าว่าง -> แปลงเป็น int -> แปลงเป็น str
    df['เลขที่ใบแจ้งหนี้'] = pd.to_numeric(df['เลขที่ใบแจ้งหนี้'], errors='coerce').fillna(0).astype(int).astype(str)

    # 2. ตั้งค่าการยุบรวมข้อมูล (Aggregating)
    agg_dict = {
        'เลขที่ใบแจ้งหนี้': lambda x: ', '.join(x),
        'รายการ': lambda x: ' | '.join(x.astype(str)),
        'Email ผู้แทน': 'first',
        'Email บัญชี': 'first'
    }

    # เช็คคอลัมน์ยอดเงิน (ถ้ามีให้รวมยอด)
    if 'รวมทั้งสิ้น' in df.columns:
        df['รวมทั้งสิ้น'] = pd.to_numeric(df['รวมทั้งสิ้น'], errors='coerce').fillna(0)
        agg_dict['รวมทั้งสิ้น'] = 'sum'

    # 3. รวมข้อมูลตามชื่อบริษัท
    df_grouped = df.groupby('บริษัท').agg(agg_dict).reset_index()

    # 4. รวม "รายการ" กับ "เลขที่ใบแจ้งหนี้" เข้าด้วยกันในช่องเดียว
    df_grouped['สรุปรายการและเลขที่ใบแจ้งหนี้'] = (
        "รายการ: " + df_grouped['รายการ'] + 
        " [เลขที่: " + df_grouped['เลขที่ใบแจ้งหนี้'] + "]"
    )

    st.success("ประมวลผลข้อมูลเรียบร้อย!")
    st.dataframe(df_grouped[['บริษัท', 'สรุปรายการและเลขที่ใบแจ้งหนี้', 'รวมทั้งสิ้น']], use_container_width=True)
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มทำงาน")
