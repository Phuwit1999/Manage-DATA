import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("ระบบรวมข้อมูลบริษัทเพื่อส่ง Email")

# 1. ส่วนของการ Upload ไฟล์
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ", type=['xlsx'])

if uploaded_file is not None:
    # อ่านไฟล์ Excel
    df = pd.read_excel(uploaded_file)
    
    # ล้างช่องว่างในชื่อคอลัมน์ (ป้องกัน Error จากการเคาะวรรคเกิน)
    df.columns = df.columns.str.strip()
    
    st.subheader("ข้อมูลต้นฉบับ")
    st.dataframe(df, use_container_width=True)

    # 2. จัดกลุ่มข้อมูล (Groupby)
    # หมายเหตุ: ชื่อคอลัมน์ต้องตรงกับในไฟล์ของคุณเป๊ะๆ
    try:
        # เตรียม Dictionary สำหรับ Aggregate
        # คอลัมน์ไหนเป็นข้อความให้ใช้ join, คอลัมน์ไหนเป็นตัวเลขให้ใช้ sum
        agg_dict = {
            'รายการ': lambda x: ' | '.join(x.astype(str)),
            'เลขที่ใบแจ้งหนี้': lambda x: ', '.join(x.astype(str)),
            'Email ผู้แทน': 'first',
            'Email บัญชี': 'first'
        }

        # ตรวจสอบคอลัมน์ตัวเลข (ถ้ามีให้รวมยอด)
        if 'รวมทั้งสิ้น' in df.columns:
            df['รวมทั้งสิ้น'] = pd.to_numeric(df['รวมทั้งสิ้น'], errors='coerce').fillna(0)
            agg_dict['รวมทั้งสิ้น'] = 'sum'
            
        if 'ค่าเบี้ยปรับ (1.5%)/เดือน ขั้นต่ำ 200บาท' in df.columns:
            fine_col = 'ค่าเบี้ยปรับ (1.5%)/เดือน ขั้นต่ำ 200บาท'
            df[fine_col] = pd.to_numeric(df[fine_col], errors='coerce').fillna(0)
            agg_dict[fine_col] = 'sum'

        # ทำการยุบรวมข้อมูลตามชื่อบริษัท
        df_grouped = df.groupby('บริษัท').agg(agg_dict).reset_index()

        st.divider()
        st.subheader("ข้อมูลที่ยุบรวมแล้ว (1 บริษัท : 1 แถว)")
        st.dataframe(df_grouped, use_container_width=True)
        
        # สามารถกด Download ผลลัพธ์เป็น CSV ได้
        csv = df_grouped.to_csv(index=False).encode('utf-8-sig')
        st.download_button("ดาวน์โหลดข้อมูลที่รวมแล้ว", csv, "grouped_data.csv", "text/csv")

    except KeyError as e:
        st.error(f"ไม่พบคอลัมน์: {e} โปรดตรวจสอบว่าชื่อคอลัมน์ในไฟล์ตรงกับในโค้ด")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการทำงาน")
