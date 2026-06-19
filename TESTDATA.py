import streamlit as st
import pandas as pd
import urllib.parse

st.set_page_config(layout="wide")
st.title("ระบบ Mail Merge สำหรับ Outlook (Fixed)")

st.markdown("""
### 📧 ระบบส่งอีเมลสรุปใบแจ้งหนี้ (Mail Merge)
**วิธีแก้ปัญหาข้อความยาวเกินไป:** 
1. กดปุ่ม **"📋 คัดลอกเนื้อหา"** 
2. กดปุ่ม **"🚀 เปิด Outlook"** 
3. เมื่อหน้าเมลเปิดขึ้นมา ให้กด **วาง (Ctrl+V)** ในช่องเนื้อหาเมลครับ
""")

uploaded_file = st.file_uploader("1️⃣ อัปโหลดไฟล์ Excel ของคุณ", type=["xlsx"])

if uploaded_file is not None:
    # --- Data Processing ---
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    df["เลขที่ใบแจ้งหนี้"] = pd.to_numeric(df["เลขที่ใบแจ้งหนี้"], errors="coerce").fillna(0).astype(int).astype(str)
    money_cols = ["ก่อนVat", "Vat", "รวมทั้งสิ้น", "ค่าเบี้ยปรับ"]
    for col in money_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False).str.strip()
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    date_cols = ["วันที่ออกใบแจ้งหนี้", "วันครบกำหนด", "คงค้างณ.วันที่"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        else:
            df[col] = pd.NaT

    other_cols = ["รายการ", "Email ผู้แทน", "Email บัญชี", "บริษัท"]
    for col in other_cols:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("")

    def format_row(row):
        inv_no       = row.get("เลขที่ใบแจ้งหนี้", "-")
        issue_date   = row["วันที่ออกใบแจ้งหนี้"].strftime("%d/%m/%Y") if pd.notna(row.get("วันที่ออกใบแจ้งหนี้")) else "-"
        item_text    = row.get("รายการ", "-")
        before_vat   = row.get("ก่อนVat", 0)
        vat_val      = row.get("Vat", 0)
        total        = row.get("รวมทั้งสิ้น", 0)
        due_date     = row["วันครบกำหนด"].strftime("%d/%m/%Y") if pd.notna(row.get("วันครบกำหนด")) else "-"
        as_of_date   = row["คงค้างณ.วันที่"].strftime("%d/%m/%Y") if pd.notna(row.get("คงค้างณ.วันที่")) else "-"
        raw          = row.get("จำนวนวันที่เกินกำหนด", 0)
        overdue_days = 0 if pd.isna(raw) else int(raw)
        penalty      = row.get("ค่าเบี้ยปรับ", 0)

        return (
            f"{inv_no} | วันที่ออกใบแจ้งหนี้ {issue_date} | {item_text} | "
            f"ก่อนVat {before_vat:,.2f} | Vat {vat_val:,.2f} | รวม {total:,.2f} | "
            f"วันครบกำหนด {due_date} | คงค้างณ.วันที่ {as_of_date} | "
            f"จำนวนวันที่เกินกำหนด {overdue_days} | ค่าเบี้ยปรับ {penalty:,.2f}"
        )

    df["formatted_text"] = df.apply(format_row, axis=1)

    def combine_items(series):
        return "\n".join([f"{i+1}. {val}" for i, val in enumerate(series)])

    df_grouped = df.groupby("บริษัท").agg({
        "formatted_text": combine_items,
        "รวมทั้งสิ้น": "sum",
        "Email ผู้แทน": "first",
        "Email บัญชี": "first"
    }).reset_index()
    
    df_grouped = df_grouped.rename(columns={"formatted_text": "รายละเอียดรายการสรุป"})

    # --- Template Section ---
    st.subheader("2️⃣ ปรับแต่งรูปแบบจดหมาย")
    custom_subject = st.text_input("หัวข้ออีเมล", value="สรุปข้อมูลใบแจ้งหนี้สำหรับ {บริษัท}")
    custom_body = st.text_area("เนื้อหาอีเมล", value="เรียน ท่านผู้เกี่ยวข้อง\n\nตามที่บริษัท {บริษัท} มีรายการใบแจ้งหนี้ดังนี้:\n\n{รายละเอียด}\n\nยอดรวมสุทธิ: {ยอดรวม} บาท\n\nจึงเรียนมาเพื่อโปรดทราบ", height=150)

    # --- Mail Merge Actions ---
    st.subheader("3️⃣ ดำเนินการส่ง")
    
    if "sent_items" not in st.session_state:
        st.session_state["sent_items"] = set()

    for index, row in df_grouped.iterrows():
        recipients = list(filter(None, set([row["Email ผู้แทน"], row["Email บัญชี"]])))
        to_str = ";".join(recipients)
        
        final_subject = custom_subject.replace("{บริษัท}", row["บริษัท"])
        final_body = custom_body.replace("{บริษัท}", row["บริษัท"]).replace("{รายละเอียด}", row["รายละเอียดรายการสรุป"]).replace("{ยอดรวม}", f"{row['รวมทั้งสิ้น']:,.2f}")
        
        # Short URL (No Body to avoid "Query string too long" error)
        encoded_subject = urllib.parse.quote(final_subject)
        outlook_web_url = f"https://outlook.office.com/mail/deeplink/compose?to={to_str}&subject={encoded_subject}"
        
        status_icon = "✅" if row["บริษัท"] in st.session_state["sent_items"] else "⏳"
        
        with st.expander(f"{status_icon} บริษัท: {row['บริษัท']}"):
            c1, c2, c3 = st.columns([1, 1, 1])
            with c1:
                # Button to copy text to clipboard using Streamlit's built-in functionality
                st.write("**ขั้นตอนที่ 1:**")
                if st.button("📋 คัดลอกเนื้อหา", key=f"copy_{index}"):
                    # In a real app, you'd use a custom JS component for true clipboard copy
                    # For now, we show the code block which has a built-in copy button in Streamlit
                    st.session_state[f"show_code_{index}"] = True
                
                if st.session_state.get(f"show_code_{index}"):
                    st.code(final_body)
                    st.caption("กดปุ่มคัดลอกที่มุมขวาบนของกล่องโค้ดด้านบน")

            with c2:
                st.write("**ขั้นตอนที่ 2:**")
                st.markdown(f'<a href="{outlook_web_url}" target="_blank" style="text-decoration:none; background-color:#0078d4; color:white; padding:10px 20px; border-radius:5px; display:inline-block;">🚀 เปิด Outlook</a>', unsafe_allow_html=True)
                st.caption("แล้วกด Ctrl+V เพื่อวางเนื้อหา")

            with c3:
                st.write("**ขั้นตอนที่ 3:**")
                if st.button("✔️ ติ๊กส่งแล้ว", key=f"done_{index}"):
                    st.session_state["sent_items"].add(row["บริษัท"])
                    st.rerun()

    if st.button("🔄 ล้างสถานะการส่งทั้งหมด"):
        st.session_state["sent_items"] = set()
        st.rerun()

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มกระบวนการ Mail Merge")
