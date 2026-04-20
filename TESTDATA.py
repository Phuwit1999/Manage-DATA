import streamlit as st
import pandas as pd

# ... (ส่วนการ Upload และอ่านไฟล์เหมือนเดิม) ...

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # 1. เตรียมข้อมูลตัวเลขให้พร้อม (จัดการจุดทศนิยมและค่าว่าง)
    df['เลขที่ใบแจ้งหนี้'] = pd.to_numeric(df['เลขที่ใบแจ้งหนี้'], errors='coerce').fillna(0).astype(int).astype(str)
    
    # จัดการคอลัมน์เงินอื่นๆ (ถ้ามี) ให้แสดงทศนิยม 2 ตำแหน่งเป็นข้อความ
    cols_to_fix = ['ก่อนVat', 'Vat', 'รวมทั้งสิ้น'] # ปรับชื่อให้ตรงกับไฟล์ของคุณ
    for col in cols_to_fix:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 2. สร้าง "รูปแบบบรรทัด" สำหรับแต่ละรายการก่อนจะ Group
    # เราจะใช้ฟังก์ชันเพื่อสร้างรูปแบบ: เลขที่ | รายการ | ก่อนVat | Vat | รวม
    def format_row(row):
        return f"{row['เลขที่ใบแจ้งหนี้']} | {row['รายการ']} | ก่อนVat {row.get('ก่อนVat', 0):,.2f} | Vat {row.get('Vat', 0):,.2f} | รวม {row.get('รวมทั้งสิ้น', 0):,.2f}"

    df['formatted_item'] = df.apply(format_row, axis=1)

    # 3. จัดกลุ่ม (Groupby) และใส่ลำดับ 1. 2. 3.
    def combine_with_index(series):
        # นำรายการที่ format แล้วมาใส่ลำดับเลขข้างหน้า 1. 2. ...
        return "\n".join([f"{i+1}. {val}" for i, val in enumerate(series)])

    df_grouped = df.groupby('บริษัท').agg({
        'formatted_item': combine_with_index,
        'รวมทั้งสิ้น': 'sum',
        'Email ผู้แทน': 'first',
        'Email บัญชี': 'first'
    }).reset_index()

    # 4. เปลี่ยนชื่อคอลัมน์ให้สื่อความหมาย
    df_grouped = df_grouped.rename(columns={'formatted_item': 'รายละเอียดรายการ'})

    st.subheader("ข้อมูลที่จัดรูปแบบแล้ว")
    # แสดงผลโดยใช้ st.text_area หรือ st.dataframe
    st.dataframe(df_grouped[['บริษัท', 'รายละเอียดรายการ', 'รวมทั้งสิ้น']], use_container_width=True)
