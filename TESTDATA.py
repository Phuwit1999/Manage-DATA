import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("ระบบสรุปข้อมูลใบแจ้งหนี้")

uploaded_file = st.file_uploader("เลือกไฟล์ Excel ของคุณ", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # Ensure 'เลขที่ใบแจ้งหนี้' is handled correctly, defaulting to 0 and then string
    df["เลขที่ใบแจ้งหนี้"] = pd.to_numeric(df["เลขที่ใบแจ้งหนี้"], errors="coerce").fillna(0).astype(int).astype(str)

    money_cols = ["ก่อนVat", "Vat", "รวมทั้งสิ้น", "ค่าเบี้ยปรับ"]
    for col in money_cols:
        if col in df.columns:
            # Convert to string first to ensure .str methods work, then clean and convert to numeric
            df[col] = df[col].astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False).str.strip()
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            # If the column does not exist, create it and fill with 0
            df[col] = 0

    date_cols = ["วันที่ออกใบแจ้งหนี้", "วันครบกำหนด", "คงค้างณ.วันที่"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        else:
            # If the column does not exist, create it and fill with NaT (Not a Time) for dates
            df[col] = pd.NaT

    # Ensure 'รายการ', 'Email ผู้แทน', 'Email บัญชี' columns exist and fill NaNs with empty string
    other_cols = ["รายการ", "Email ผู้แทน", "Email บัญชี", "บริษัท"]
    for col in other_cols:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("")

    def format_row(row):
        inv_no       = row.get("เลขที่ใบแจ้งหนี้", "-")
        issue_date   = row["วันที่ออกใบแจ้งหนี้"].strftime("%d/%m/%Y") if pd.notna(row.get("วันที่ออกใบแจ้งหนี้")) else "-"
        item_text    = row.get("รายการ", "-")
        before_vat   = row.get("ก่อนแวต", 0)
        vat_val      = row.get("แวต", 0)
        total        = row.get("รวมทั้งสิ้น", 0)
        due_date     = row["วันครบกำหนด"].strftime("%d/%m/%Y") if pd.notna(row.get("วันครบกำหนด")) else "-"
        as_of_date   = row["คงค้างณ.วันที่"].strftime("%d/%m/%Y") if pd.notna(row.get("คงค้างณ.วันที่")) else "-"
        raw          = row.get("จำนวนวันที่เกินกำหนด", 0)
        overdue_days = 0 if pd.isna(raw) else int(raw)
        penalty      = row.get("ค่าเบี้ยปรับ(1.5%)/เดือน ขั้นต่ำ 200บาท", )

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
    st.dataframe(df_grouped, use_container_width=True)

    if st.checkbox("แสดงตัวอย่างข้อความสำหรับ Email"):
        for index, row in df_grouped.iterrows():
            st.text(f"บริษัท: {row["บริษัท"]}")
            st.code(row["รายละเอียดรายการสรุป"])
            st.write(f"ยอดรวมสุทธิ: {row["รวมทั้งสิ้น"]:,.2f}")
            st.divider()
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มการทำงาน")
