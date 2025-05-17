import streamlit as st
import pandas as pd
import pdfplumber
import tabula
import tempfile
import re
import logging

logging.getLogger("pdfminer").setLevel(logging.ERROR)

# === Helper: Clean Ship To ===
def extract_ship_to(lines):
    for i, line in enumerate(lines):
        if line.startswith("SHIP TO:"):
            clean_line = line.replace("REFERENCE:", "").strip()
            clean_line = re.sub(r"SHIP TO:\s*", "", clean_line)
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                if not next_line.startswith(("VENDOR", "ADDRESS", "REFERENCE", "TERM", "SHIPPING", "PAYMENT")):
                    clean_line += " " + next_line.strip()
            return clean_line.strip()
    return ""

def extract_header_summary(pdf_file):
    header_data = []
    summary_data = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = [line.strip() for line in text.split("\n") if line.strip()]
            joined = "\n".join(lines)

            def extract(pattern, text, group=1):
                match = re.search(pattern, text, re.IGNORECASE)
                return match.group(group).strip() if match else ""

            po_number = extract(r"PO NUMBER:\s*(\d+)", joined)
            date = extract(r"DATE:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", joined)
            ship_to = extract_ship_to(lines)
            vendor_name = extract(r"VENDOR NAME:\s*(.*?)\s+SHIPPING DATE", joined)
            shipping_date = extract(r"SHIPPING DATE:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", joined)

            include_vat = extract(r"INCLUDE VAT\s+([0-9.]+)", joined)
            tax = extract(r"TAX\s+([0-9.]+)", joined)
            usd_total = extract(r"([0-9.]+)\s+USD", joined)

            header_data.append({
                "PO Number": po_number,
                "Date": date,
                "Ship To": ship_to,
                "Vendor Name": vendor_name,
                "Shipping Date": shipping_date
            })

            summary_data.append({
                "INCLUDE VAT": include_vat,
                "TAX": tax,
                "TOTAL (USD)": usd_total
            })

    return header_data, summary_data

def get_column_index_by_keywords(df, keywords):
    for row_idx in range(len(df)):
        row = df.iloc[row_idx]
        for idx, val in enumerate(row):
            if pd.isna(val):
                continue
            cell = str(val).strip().upper()
            for keyword in keywords:
                if keyword.upper() in cell:
                    return idx
    return None

def clean_table(raw_df):
    if raw_df.empty or len(raw_df.columns) < 7:
        return pd.DataFrame()

    unit_qty_col = int(get_column_index_by_keywords(raw_df, ["UNIT QUANTITY"]))
    unit_price_col = int(get_column_index_by_keywords(raw_df, ["UNIT PRICE"]))
    amount_col = int(len(raw_df.columns)-1)

    cleaned_rows = []

    for i in range(len(raw_df)-1):
        row = raw_df.iloc[i]
        next_row = raw_df.iloc[i + 1]

        if re.match(r"^\d{10,}$", str(row[0]).strip()):
            product = str(row[0]).strip()
            unit_qty = str(row[unit_qty_col]).strip() if len(row) > unit_qty_col else ""
            try:
                unit_price = round(float(str(row[unit_price_col]).strip()), 2)
            except:
                unit_price = ""

            try:
                amount = round(float(str(next_row[amount_col]).strip()), 2)
            except:
                amount = ""

            item_no = str(next_row[0]).strip()
            code = str(next_row[1]).strip()
            description = str(next_row[2]).strip()

            if not item_no.isdigit():
                continue

            cleaned_rows.append([
                item_no,
                product,
                code,
                description,
                unit_qty,
                unit_price,
                amount
            ])

    cleaned_df = pd.DataFrame(cleaned_rows, columns=[
        "Item#",
        "Product",
        "Code",
        "Description",
        "Unit Quantity (Unit/Pack/Case)",
        "Unit Price (USD)",
        "Amount (USD)"
    ])

    return cleaned_df

# === Streamlit UI ===
st.title("📄 แปลง PDF ใบสั่งซื้อเป็น Excel")

uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        pdf_path = tmp.name

    header_data, summary_data = extract_header_summary(pdf_path)
    tables = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True)

    # Create Excel
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as out_xlsx:
        with pd.ExcelWriter(out_xlsx.name, engine="openpyxl") as writer:
            for i, (head, summary) in enumerate(zip(header_data, summary_data)):
                sheet = f"Page{i+1}"
                header_df = pd.DataFrame({"Field": list(head.keys()), "Value": list(head.values())})
                summary_df = pd.DataFrame([summary])
                table = clean_table(tables[i]) if i < len(tables) else pd.DataFrame()

                header_df.to_excel(writer, sheet_name=sheet, index=False, startrow=0)
                table.to_excel(writer, sheet_name=sheet, index=False, startrow=len(header_df)+2)
                summary_df.to_excel(writer, sheet_name=sheet, index=False, startrow=len(header_df)+len(table)+4)

        st.success("✅ แปลงเสร็จแล้ว! กดดาวน์โหลดได้เลย")
        with open(out_xlsx.name, "rb") as f:
            st.download_button("⬇️ ดาวน์โหลด Excel", f, file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
