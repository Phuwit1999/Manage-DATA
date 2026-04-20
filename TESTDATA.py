import streamlit as st
import pandas as pd

# ... (ส่วนการโหลดไฟล์เหมือนเดิม) ...

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # 1. จัดการเลขที่ใบแจ้งหนี้: แปลงเป็นตัวเลข -> เอาทศนิยมออก -> แปลงเป็นข้อความ
    # ใช้ errors='coerce' เพื่อป้องกันกรณีเจอข้อความที่ไม่ใช่ตัวเลข
    df['เลขที่ใบแจ้งหนี้'] = pd.to_numeric(df['เลขที่ใบแจ้งหนี้'], errors='coerce').fillna(0).astype(int).astype(str)

    # 2. จัดกลุ่มและรวมข้อมูล
    agg_dict = {
        'เลขที่ใบแจ้งหนี้': lambda x: ', '.join(x),
        'รายการ': lambda x: ' | '.join(x.astype(str)),
        'Email ผู้แทน': 'first',
        'Email บัญชี': 'first'
    }

    # (รวมยอดเงินถ้ามีคอลัมน์นั้น)
    if 'รวมทั้งสิ้น' in df.columns:
        df['รวมทั้งสิ้น'] = pd.to_numeric(df['รวมทั้งสิ้น'], errors='coerce').fillna(0)
        agg_dict['รวมทั้งสิ้น'] = 'sum'

    # ทำการยุบรวมตามชื่อบริษัท
    df_grouped = df.groupby('บริษัท').agg(agg_dict).reset_index()

    # 3. สร้างคอลัมน์ใหม่: รวม "รายการ" และ "เลขที่ใบแจ้งหนี้" เข้าด้วยกัน
    df_grouped['ข้อมูลการแจ้งหนี้'] = (
        "รายการ: " + df_grouped['รายการ'] + 
        " (เลขที่ใบแจ้งหนี้: " + df_grouped['เลขที่ใบแจ้งหนี้'] + ")"
    )

    st.subheader("ผลลัพธ์: รายการและเลขที่ใบแจ้งหนี้รวมกันแล้ว")
    # แสดงตารางโดยเน้นคอลัมน์ที่รวมกันแล้ว
    st.dataframe(df_grouped[['บริษัท', 'ข้อมูลการแจ้งหนี้', 'รวมทั้งสิ้น']], use_container_width=True)
