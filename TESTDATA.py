import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("ระบบสรุปข้อมูลใบแจ้งหนี้")

# --- บรรทัดที่ 10: ต้องสร้างตัวแปร uploaded_file ก่อนจะไปเช็ค if ---
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ", type=['xlsx'])

if uploaded_file is not None:
    # อ่านไฟล์
    df = pd.read_excel(uploaded_file)
    # ล้างช่องว่างที่หัวตาราง
    df.columns = df.columns.str.strip()

    # 1. จัดการข้อมูลเบื้องต้น
    # ตัด .0 ออกจากเลขที่ใบแจ้งหนี้
    df['เลขที่ใบแจ้งหนี้'] = pd.to_numeric(df['เลขที่ใบแจ้งหนี้'], errors='coerce').fillna(0).astype(int).astype(str)
    
    # แปลงยอดเงินให้เป็นตัวเลข (เพื่อความปลอดภัยในการคำนวณ)
    money_cols = ['ก่อนVat', 'Vat', 'รวมทั้งสิ้น']
    for col in money_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 2. สร้างฟังก์ชันจัดรูปแบบ "เลขที่ | รายการ | ก่อนVat | Vat | รวม"
    def format_row(row):
        # ดึงค่าจากคอลัมน์ (ถ้าไม่มีให้เป็น 0)
        before_vat = row.get('ก่อนVat', 0)
        vat_val = row.get('Vat', 0)
        total = row.get('รวมทั้งสิ้น', 0)
        item_text = row.get('รายการ', '-')
        inv_no = row['เลขที่ใบแจ้งหนี้']
        
        return f"{inv_no} | {item_text} | ก่อนVat {before_vat:,.2f} | Vat {vat_val:,.2f} | รวม {total:,.2f}"

    # นำรูปแบบที่สร้างไว้ไปใส่ในคอลัมน์ใหม่
    df['formatted_text'] = df.apply(format_row, axis=1)

    # 3. ยุบรวมตามบริษัท และใส่เลขลำดับ 1. 2. 3.
    def combine_items(series):
        # เอาแต่ละรายการที่ format แล้วมาใส่เลขลำดับ 1. 2. ... แล้วขึ้นบรรทัดใหม่
        return "\n".join([f"{i+1}. {val}" for i, val in enumerate(series)])

    # ทำการ Groupby
    df_grouped = df.groupby('บริษัท').agg({
        'formatted_text': combine_items,
        'รวมทั้งสิ้น': 'sum',
        'Email ผู้แทน': 'first',
        'Email บัญชี': 'first'
    }).reset_index()

    # เปลี่ยนชื่อคอลัมน์ให้ดูง่าย
    df_grouped = df_grouped.rename(columns={'formatted_text': 'รายละเอียดรายการสรุป'})

    # 4. แสดงผล
    st.success("รวมข้อมูลสำเร็จ!")
    st.dataframe(df_grouped, use_container_width=True)
    
    # แสดงตัวอย่างแบบข้อความ (สำหรับก๊อปปี้ไปส่ง Email)
    if st.checkbox("แสดงตัวอย่างข้อความสำหรับ Email"):
        for index, row in df_grouped.iterrows():
            st.text(f"บริษัท: {row['บริษัท']}")
            st.code(row['รายละเอียดรายการสรุป'])
            st.write(f"ยอดรวมสุทธิ: {row['รวมทั้งสิ้น']:,.2f}")
            st.divider()

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการทำงาน")
