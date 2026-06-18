import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

st.set_page_config(layout="wide")
st.title("ระบบสรุปข้อมูลใบแจ้งหนี้ (Customizable Auto-Email)")

# Email configuration (replace with your actual details or Streamlit secrets)
SMTP_SERVER = "smtp.office365.com"  # SMTP Server for Hotmail/Outlook
SMTP_PORT = 587
SENDER_EMAIL = "your_email@hotmail.com"
SENDER_PASSWORD = "your_email_password"

def send_email(receiver_email, subject, body):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = receiver_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        return str(e)

uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # Data Cleaning & Processing
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

    st.success("รวมข้อมูลสำเร็จ!")

    # Template Editor
    st.subheader("📝 ปรับแต่งรูปแบบอีเมล (Template Editor)")
    st.info("คุณสามารถใช้ตัวแปรดังนี้: {บริษัท}, {รายละเอียด}, {ยอดรวม}")
    
    default_subject = "สรุปข้อมูลใบแจ้งหนี้สำหรับ {บริษัท}"
    default_body = """เรียน ท่านผู้เกี่ยวข้อง

ตามที่บริษัท {บริษัท} มีรายการใบแจ้งหนี้ดังนี้:

{รายละเอียด}

ยอดรวมสุทธิ: {ยอดรวม} บาท

จึงเรียนมาเพื่อโปรดทราบ

ขอแสดงความนับถือ
ระบบสรุปข้อมูลอัตโนมัติ"""

    custom_subject = st.text_input("หัวข้ออีเมล (Subject)", value=default_subject)
    custom_body = st.text_area("เนื้อหาอีเมล (Body)", value=default_body, height=250)

    # Validation Check
    st.subheader("📊 ตรวจสอบสถานะและความถูกต้อง")
    validation_results = []
    for idx, row in df_grouped.iterrows():
        missing = []
        if not row["Email ผู้แทน"] and not row["Email บัญชี"]:
            missing.append("ไม่มีอีเมล")
        if not row["บริษัท"]:
            missing.append("ไม่มีชื่อบริษัท")
        
        status = "✅ พร้อมส่ง" if not missing else f"❌ ไม่พร้อม ({', '.join(missing)})"
        validation_results.append(status)
    
    df_grouped["สถานะ"] = validation_results
    st.dataframe(df_grouped, use_container_width=True)

    # Preview
    if st.checkbox("ดูตัวอย่างอีเมลที่จะส่ง"):
        if not df_grouped.empty:
            sample_row = df_grouped.iloc[0]
            p_subject = custom_subject.replace("{บริษัท}", sample_row["บริษัท"])
            p_body = custom_body.replace("{บริษัท}", sample_row["บริษัท"])\
                                .replace("{รายละเอียด}", sample_row["รายละเอียดรายการสรุป"])\
                                .replace("{ยอดรวม}", f"{sample_row['รวมทั้งสิ้น']:,.2f}")
            
            st.markdown("---")
            st.write("**ตัวอย่างที่จะถูกส่ง (บริษัทแรก):**")
            st.text(f"Subject: {p_subject}")
            st.code(p_body)
            st.markdown("---")

    # Auto Send
    auto_send = st.toggle("เปิดระบบส่งอัตโนมัติ (Auto-Send Mode)")
    
    if auto_send:
        st.warning("ระบบจะใช้รูปแบบ (Template) ด้านบนในการส่งอีเมลทั้งหมด")
        if st.button("เริ่มส่งอีเมลทั้งหมด"):
            if not SENDER_EMAIL or SENDER_EMAIL == "your_email@hotmail.com":
                st.error("กรุณาตั้งค่าอีเมลผู้ส่งในโค้ดก่อน")
            else:
                ready_to_send = df_grouped[df_grouped["สถานะ"] == "✅ พร้อมส่ง"]
                if ready_to_send.empty:
                    st.warning("ไม่มีข้อมูลที่พร้อมส่ง")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    total_rows = len(ready_to_send)
                    sent_count = 0
                    
                    for index, row in ready_to_send.iterrows():
                        recipients = list(filter(None, set([row["Email ผู้แทน"], row["Email บัญชี"]])))
                        
                        # Replace placeholders with actual data
                        final_subject = custom_subject.replace("{บริษัท}", row["บริษัท"])
                        final_body = custom_body.replace("{บริษัท}", row["บริษัท"])\
                                               .replace("{รายละเอียด}", row["รายละเอียดรายการสรุป"])\
                                               .replace("{ยอดรวม}", f"{row['รวมทั้งสิ้น']:,.2f}")
                        
                        success_any = False
                        for recipient in recipients:
                            res = send_email(recipient, final_subject, final_body)
                            if res is True:
                                success_any = True
                                st.write(f"✔️ ส่งถึง {recipient} ({row['บริษัท']}) สำเร็จ")
                            else:
                                st.error(f"❌ ส่งถึง {recipient} ({row['บริษัท']}) ล้มเหลว: {res}")
                        
                        sent_count += 1
                        progress_bar.progress(sent_count / total_rows)
                        status_text.text(f"กำลังส่ง: {sent_count}/{total_rows} บริษัท")
                        time.sleep(1) # Delay for Hotmail/Outlook
                    
                    st.success(f"ดำเนินการเสร็จสิ้น! ส่งข้อมูลสำเร็จทั้งหมด {sent_count} บริษัท")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการทำงาน")
