import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time

st.set_page_config(layout="wide")
st.title("ระบบสรุปข้อมูลใบแจ้งหนี้ (Outlook Auto-Email)")

# --- Sidebar: Email Connection Settings ---
with st.sidebar:
    st.header("🔐 การเชื่อมต่อ Outlook")
    st.write("เชื่อมโยงบัญชีของคุณเพื่อส่งอีเมล")
    
    # Try to get from secrets first, otherwise empty
    default_email = st.secrets.get("SENDER_EMAIL", "your_email@outlook.com")
    default_password = st.secrets.get("SENDER_PASSWORD", "")
    
    sender_email = st.text_input("อีเมลผู้ส่ง", value=default_email)
    sender_password = st.text_input("รหัสผ่าน (App Password)", type="password", value=default_password)
    
    st.info("""
    **วิธีเชื่อมโยงโดยไม่ต้องใส่รหัสในโค้ด:**
    1. กรอกอีเมลและรหัสผ่านที่นี่ (ข้อมูลจะอยู่ใน Session ชั่วคราว)
    2. หรือใช้ไฟล์ `.streamlit/secrets.toml` เพื่อเชื่อมโยงแบบถาวร
    """)
    
    st.divider()
    st.write("⚙️ **การตั้งค่าเซิร์ฟเวอร์**")
    smtp_server = "smtp.office365.com"
    smtp_port = 587

def send_email(receiver_email, subject, body, s_email, s_password):
    try:
        msg = MIMEMultipart()
        msg['From'] = s_email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(s_email, s_password)
            server.send_message(msg)
        return True
    except Exception as e:
        return str(e)

# --- Main App Logic ---
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
    st.subheader("📝 ปรับแต่งรูปแบบอีเมล")
    custom_subject = st.text_input("หัวข้ออีเมล", value="สรุปข้อมูลใบแจ้งหนี้สำหรับ {บริษัท}")
    custom_body = st.text_area("เนื้อหาอีเมล", value="เรียน ท่านผู้เกี่ยวข้อง\n\nตามที่บริษัท {บริษัท} มีรายการใบแจ้งหนี้ดังนี้:\n\n{รายละเอียด}\n\nยอดรวมสุทธิ: {ยอดรวม} บาท\n\nจึงเรียนมาเพื่อโปรดทราบ", height=200)

    # Validation & Table
    st.subheader("📊 ตรวจสอบความพร้อม")
    validation_results = ["✅ พร้อมส่ง" if (row["Email ผู้แทน"] or row["Email บัญชี"]) and row["บริษัท"] else "❌ ไม่พร้อม" for _, row in df_grouped.iterrows()]
    df_grouped["สถานะ"] = validation_results
    st.dataframe(df_grouped, use_container_width=True)

    # Auto Send Logic
    if st.toggle("เปิดระบบส่งอัตโนมัติ"):
        if not sender_password:
            st.warning("⚠️ กรุณากรอกรหัสผ่านที่แถบด้านซ้าย (Sidebar) เพื่อเชื่อมต่อ Outlook")
        else:
            if st.button("เริ่มส่งอีเมลทั้งหมด"):
                ready_to_send = df_grouped[df_grouped["สถานะ"] == "✅ พร้อมส่ง"]
                if not ready_to_send.empty:
                    progress_bar = st.progress(0)
                    sent_count = 0
                    for index, row in ready_to_send.iterrows():
                        recipients = list(filter(None, set([row["Email ผู้แทน"], row["Email บัญชี"]])))
                        final_subject = custom_subject.replace("{บริษัท}", row["บริษัท"])
                        final_body = custom_body.replace("{บริษัท}", row["บริษัท"]).replace("{รายละเอียด}", row["รายละเอียดรายการสรุป"]).replace("{ยอดรวม}", f"{row['รวมทั้งสิ้น']:,.2f}")
                        
                        for recipient in recipients:
                            res = send_email(recipient, final_subject, final_body, sender_email, sender_password)
                            if res is True:
                                st.write(f"✔️ {row['บริษัท']} -> {recipient} สำเร็จ")
                            else:
                                st.error(f"❌ {row['บริษัท']} -> {recipient} ล้มเหลว: {res}")
                        
                        sent_count += 1
                        progress_bar.progress(sent_count / len(ready_to_send))
                        time.sleep(1)
                    st.success(f"ส่งสำเร็จทั้งหมด {sent_count} บริษัท")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการทำงาน")
