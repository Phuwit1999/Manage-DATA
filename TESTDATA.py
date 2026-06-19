import streamlit as st
import pandas as pd
import time
st.set_page_config(layout="wide")
st.title("ระบบ Mail Merge สำหรับ Outlook (ส่งอัตโนมัติ)")

st.markdown("""
### 📧 ระบบส่งอีเมลสรุปใบแจ้งหนี้ (Mail Merge - ส่งอัตโนมัติผ่าน Outlook Desktop)
ระบบนี้จะสั่งให้ **Outlook Desktop ที่เปิด/ล็อกอินอยู่บนเครื่องนี้** สร้างและส่งอีเมลให้ทันที
โดยไม่ต้องคัดลอก-วาง หรือกดส่งเองอีก

> ⚠️ **ข้อกำหนด:** ต้องรันบน **Windows** ที่มี **Outlook Desktop ติดตั้งและล็อกอินไว้แล้ว**
> และต้องติดตั้งไลบรารีก่อน: `pip install pywin32`
> (วิธีนี้ใช้ไม่ได้ถ้ารัน Streamlit บนเซิร์ฟเวอร์ Linux/Cloud หรือไม่มี Outlook Desktop)
""")

# --- ตรวจสอบว่าเชื่อมต่อ Outlook ได้หรือไม่ ---
OUTLOOK_AVAILABLE = False
outlook_error_msg = ""
try:
    import win32com.client  # noqa: F401
    OUTLOOK_AVAILABLE = True
except Exception as e:
    outlook_error_msg = str(e)

if not OUTLOOK_AVAILABLE:
    st.error(
        "ไม่พบไลบรารี pywin32 หรือไม่สามารถใช้งาน Outlook COM ได้บนเครื่องนี้ "
        "(ต้องรันบน Windows ที่มี Outlook Desktop ติดตั้งไว้) "
        f"รายละเอียด: {outlook_error_msg}"
    )

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
    custom_body = st.text_area(
        "เนื้อหาอีเมล",
        value="เรียน ท่านผู้เกี่ยวข้อง\n\nตามที่บริษัท {บริษัท} มีรายการใบแจ้งหนี้ดังนี้:\n\n{รายละเอียด}\n\nยอดรวมสุทธิ: {ยอดรวม} บาท\n\nจึงเรียนมาเพื่อโปรดทราบ",
        height=150
    )

    def send_email_outlook(to_str, subject, body):
        """สั่ง Outlook Desktop สร้างและส่งอีเมลทันที (ไม่เปิดหน้าต่างให้กดส่งเอง)"""
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to_str
        mail.Subject = subject
        mail.Body = body
        mail.Send()

    def build_row_content(row):
        recipients = list(filter(None, set([row["Email ผู้แทน"], row["Email บัญชี"]])))
        to_str = ";".join(recipients)
        final_subject = custom_subject.replace("{บริษัท}", row["บริษัท"])
        final_body = (
            custom_body
            .replace("{บริษัท}", row["บริษัท"])
            .replace("{รายละเอียด}", row["รายละเอียดรายการสรุป"])
            .replace("{ยอดรวม}", f"{row['รวมทั้งสิ้น']:,.2f}")
        )
        return to_str, final_subject, final_body

    # --- Mail Merge Actions ---
    st.subheader("3️⃣ ดำเนินการส่ง")

    if "sent_items" not in st.session_state:
        st.session_state["sent_items"] = set()

    confirm_all = st.checkbox("ฉันตรวจสอบเนื้อหาอีเมลทุกฉบับแล้ว และยืนยันให้ส่งจริงทั้งหมด")

    col_a, col_b = st.columns([1, 1])
    with col_a:
        send_all_clicked = st.button(
            "🚀 ส่งอีเมลทั้งหมดอัตโนมัติ",
            disabled=(not confirm_all) or (not OUTLOOK_AVAILABLE),
            type="primary"
        )
    with col_b:
        if st.button("🔄 ล้างสถานะการส่งทั้งหมด"):
            st.session_state["sent_items"] = set()
            st.rerun()

    if send_all_clicked:
        pending_rows = [row for _, row in df_grouped.iterrows() if row["บริษัท"] not in st.session_state["sent_items"]]
        total = len(pending_rows)
        if total == 0:
            st.info("ทุกบริษัทถูกส่งไปแล้วครับ")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            errors = []
            for i, row in enumerate(pending_rows):
                to_str, final_subject, final_body = build_row_content(row)
                try:
                    send_email_outlook(to_str, final_subject, final_body)
                    st.session_state["sent_items"].add(row["บริษัท"])
                    status_text.text(f"✅ ส่งแล้ว: {row['บริษัท']} ({i+1}/{total})")
                except Exception as e:
                    errors.append((row["บริษัท"], str(e)))
                    status_text.text(f"❌ ส่งไม่สำเร็จ: {row['บริษัท']} - {e}")
                progress_bar.progress((i + 1) / total)
                time.sleep(0.3)  # หน่วงเล็กน้อยกัน Outlook ค้างตอนยิงเมลรัวๆ

            if errors:
                st.error(f"ส่งไม่สำเร็จ {len(errors)} รายการ: " + ", ".join([e[0] for e in errors]))
            else:
                st.success("ส่งอีเมลครบทุกรายการเรียบร้อยแล้ว 🎉")
            st.rerun()

    st.divider()

    for index, row in df_grouped.iterrows():
        to_str, final_subject, final_body = build_row_content(row)
        is_sent = row["บริษัท"] in st.session_state["sent_items"]
        status_icon = "✅" if is_sent else "⏳"

        with st.expander(f"{status_icon} บริษัท: {row['บริษัท']}  |  ผู้รับ: {to_str or '— ไม่มีอีเมล —'}"):
            st.text_area("ตัวอย่างเนื้อหา (Preview)", value=final_body, height=150, key=f"preview_{index}", disabled=True)

            c1, c2 = st.columns([1, 1])
            with c1:
                if st.button(
                    "✉️ ส่งฉบับนี้เลย" if not is_sent else "✉️ ส่งซ้ำอีกครั้ง",
                    key=f"send_{index}",
                    disabled=not OUTLOOK_AVAILABLE or not to_str
                ):
                    try:
                        send_email_outlook(to_str, final_subject, final_body)
                        st.session_state["sent_items"].add(row["บริษัท"])
                        st.success(f"ส่งอีเมลถึง {row['บริษัท']} สำเร็จแล้ว")
                        st.rerun()
                    except Exception as e:
                        st.error(f"ส่งไม่สำเร็จ: {e}")
            with c2:
                if is_sent:
                    if st.button("↩️ ยกเลิกสถานะส่งแล้ว", key=f"undo_{index}"):
                        st.session_state["sent_items"].discard(row["บริษัท"])
                        st.rerun()

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มกระบวนการ Mail Merge")
