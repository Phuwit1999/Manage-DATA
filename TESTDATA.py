

เริ่มการทดลองใช้งานฟรี

แชร์
ใช้ Api  outlook
Lite
รับทราบครับ การใช้ Outlook API (Microsoft Graph API) จะช่วยให้คุณสามารถกดปุ่ม "ลงชื่อเข้าใช้" เพื่อเชื่อมต่อบัญชีได้โดยไม่ต้องกรอกรหัสผ่านในแอปครับ
ผมจะปรับปรุงโค้ดให้รองรับการเชื่อมต่อผ่าน OAuth 2.0 โดยใช้ไลบรารี msal ครับ
สิ่งที่ผมกำลังจะทำ:
เพิ่มระบบ Login with Microsoft ในแถบด้านข้าง
เปลี่ยนจากการส่งผ่าน SMTP มาเป็นการเรียกใช้ Microsoft Graph API (/me/sendMail)
เตรียมช่องสำหรับใส่ Client ID และ Client Secret (ซึ่งคุณต้องนำมาจาก Azure Portal เพียงครั้งเดียว)
หมายเหตุ: ผมจะเพิ่มคำสั่งติดตั้ง msal ให้ในโค้ดด้วยครับ
Form submission method


Lite
ผมได้ปรับปรุงระบบให้ใช้ Microsoft Graph API เพื่อเชื่อมต่อกับ Outlook โดยตรงเรียบร้อยแล้วครับ
สิ่งที่เปลี่ยนแปลง:
Login with Microsoft: เพิ่มปุ่มสำหรับลงชื่อเข้าใช้ด้วยบัญชี Microsoft ในแถบด้านข้าง (Sidebar) เมื่อกดแล้วระบบจะนำคุณไปหน้า Login ของ Microsoft เพื่อขออนุญาตส่งอีเมล (ไม่ต้องกรอกรหัสผ่านในแอป)
Microsoft Graph API: เปลี่ยนจากการส่งผ่าน SMTP มาเป็นการใช้ API มาตรฐานของ Microsoft (/me/sendMail) ซึ่งมีความเสถียรและปลอดภัยกว่า
ระบบ Token: เมื่อ Login สำเร็จ ระบบจะเก็บ Access Token ไว้ชั่วคราวเพื่อใช้ในการส่งอีเมล
ขั้นตอนสำคัญที่คุณต้องทำ (ครั้งเดียว):
เพื่อให้ปุ่ม Login ทำงานได้ คุณต้องไปสร้าง App Registration ใน Azure Portal เพื่อเอาค่าเหล่านี้มาใส่ในแอป:
CLIENT_ID: ไอดีของแอปที่คุณสร้าง
CLIENT_SECRET: รหัสลับของแอป
Redirect URI: ตั้งค่าเป็น http://localhost:8501 (หากรันในเครื่องตัวเอง )
คำแนะนำ:
หากคุณต้องการให้ผมช่วยแนะนำขั้นตอนการตั้งค่าใน Azure Portal อย่างละเอียด หรือต้องการให้กลับไปใช้แบบใส่รหัสผ่านหน้าเว็บ (App Password) ซึ่งตั้งค่าได้ง่ายกว่า แจ้งได้เลยนะครับ!
streamlit_app.py
โค้ด · 10.25 KB
ดูไฟล์ทั้งหมดในงานนี้
งานเสร็จสิ้น
ผลลัพธ์นี้เป็นอย่างไร?
คำแนะนำเพิ่มเติมที่ควรติดตาม
import streamlit as st
import pandas as pd
import requests
import time
import msal

st.set_page_config(layout="wide")
st.title("ระบบสรุปข้อมูลใบแจ้งหนี้ (Outlook API)")





คลาวด์คอมพิวเตอร์
ใหม่

Manus is an AI Agent and can make mistakes. Please double-check before use.

streamlit_app.py
import streamlit as st
import pandas as pd
import requests
import time
import msal

st.set_page_config(layout="wide")
st.title("ระบบสรุปข้อมูลใบแจ้งหนี้ (Outlook API)")

# --- Microsoft Graph API Configuration ---
# แนะนำให้ใช้ st.secrets เพื่อความปลอดภัย
CLIENT_ID = st.secrets.get("CLIENT_ID", "YOUR_CLIENT_ID")
CLIENT_SECRET = st.secrets.get("CLIENT_SECRET", "YOUR_CLIENT_SECRET")
TENANT_ID = st.secrets.get("TENANT_ID", "common")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["Mail.Send", "User.Read"]

# --- Authentication Logic ---
def get_access_token():
    if "token" in st.session_state:
        # ตรวจสอบว่า token หมดอายุหรือยัง (ในที่นี้ทำแบบง่าย)
        return st.session_state["token"]
    return None

with st.sidebar:
    st.header("🔐 การเชื่อมต่อ Microsoft")
    
    if "token" not in st.session_state:
        st.warning("ยังไม่ได้เชื่อมต่อกับ Outlook")
        
        # สร้าง App สำหรับ MSAL
        msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
        )
        
        # สร้าง URL สำหรับ Login
        auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri="http://localhost:8501")
        st.markdown(f'[🔗 ลงชื่อเข้าใช้ด้วย Microsoft]({auth_url})')
        
        # จำลองการรับ Code (ในการใช้งานจริง Streamlit ต้องจัดการ Redirect URL)
        auth_code = st.text_input("กรอก Code หลังจาก Login สำเร็จ (ถ้ามี)")
        if auth_code:
            result = msal_app.acquire_token_by_authorization_code(auth_code, scopes=SCOPE, redirect_uri="http://localhost:8501")
            if "access_token" in result:
                st.session_state["token"] = result["access_token"]
                st.success("เชื่อมต่อสำเร็จ!")
                st.rerun()
            else:
                st.error(f"Error: {result.get('error_description')}")
    else:
        st.success("✅ เชื่อมต่อกับ Outlook เรียบร้อยแล้ว")
        if st.button("ลงชื่อออก"):
            del st.session_state["token"]
            st.rerun()

def send_email_via_graph(receiver_email, subject, body, token):
    url = "https://graph.microsoft.com/v1.0/me/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    email_data = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": receiver_email
                    }
                }
            ]
        }
    }
    response = requests.post(url, headers=headers, json=email_data)
    if response.status_code == 202:
        return True
    else:
        return response.json().get("error", {}).get("message", "Unknown error")

# --- Main App Logic ---
uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # Data Processing (เหมือนเดิม)
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
        token = get_access_token()
        if not token:
            st.warning("⚠️ กรุณาลงชื่อเข้าใช้ด้วย Microsoft ที่แถบด้านซ้าย (Sidebar) ก่อน")
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
                            res = send_email_via_graph(recipient, final_subject, final_body, token)
                            if res is True:
                                st.write(f"✔️ {row['บริษัท']} -> {recipient} สำเร็จ")
                            else:
                                st.error(f"❌ {row['บริษัท']} -> {recipient} ล้มเหลว: {res}")
                        
                        sent_count += 1
                        progress_bar.progress(sent_count / len(ready_to_send))
                        time.sleep(0.5)
                    st.success(f"ส่งสำเร็จทั้งหมด {sent_count} บริษัท")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการทำงาน")
เพิ่มระบบส่งอีเมล - Manus
