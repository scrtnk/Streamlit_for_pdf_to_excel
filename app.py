import streamlit as st
import pandas as pd
import pdfplumber
import tabula
import tempfile
import re
import logging
import io

logging.getLogger("pdfminer").setLevel(logging.ERROR)

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


def clean_table_flexible(raw_df):
    if raw_df.empty or len(raw_df.columns) < 5:
        return pd.DataFrame()
    desc_col = get_column_index_by_keywords(raw_df, ["DESCRIPTION"])
    if desc_col is None:
        desc_col = 2
    vat_col = desc_col + 1 if desc_col is not None else None
    unit_size_col = vat_col  # อยู่ row ถัดไป
    percent_vat_col = get_column_index_by_keywords(raw_df, ["%VAT", "10.00", "10"])
    pack_size_col = get_column_index_by_keywords(raw_df, ["PACK"])

    unit_qty_col = get_column_index_by_keywords(raw_df, ["UNIT QUANTITY"])
    unit_price_col = get_column_index_by_keywords(raw_df, ["UNIT PRICE"])
    amount_col = len(raw_df.columns) - 1

    import re
    cleaned_rows = []

    for i in range(len(raw_df)-1):
        row = raw_df.iloc[i]
        next_row = raw_df.iloc[i + 1]

        if re.match(r"^\d{10,}$", str(row.iloc[0]).strip()):
            product = str(row.iloc[0]).strip()
            unit_qty = str(row.iloc[unit_qty_col]).strip() if unit_qty_col is not None else ""
            try:
                unit_price = round(float(str(row.iloc[unit_price_col]).strip()), 2)
            except:
                unit_price = ""
            try:
                amount = round(float(str(next_row.iloc[amount_col]).strip()), 2)
            except:
                amount = ""
            item_no = str(next_row.iloc[0]).strip()
            code_number = str(next_row.iloc[1]).strip()
            match = re.match(r"^\d+", code_number)
            if match:
                code = match.group()
            else:
                code = ""

            # --- Description fallback logic ---
            invalid_vals = ["", "nan"]

            def is_invalid(val):
                return str(val).strip().lower() in invalid_vals

            description = ""
            if desc_col is not None:
                candidates = [
                    next_row.iloc[desc_col] if desc_col < len(next_row) else "",
                    next_row.iloc[desc_col - 1] if desc_col - 1 >= 0 else "",
                    next_row.iloc[2] if 2 < len(next_row) else "",
                    next_row.iloc[1] if 1 < len(next_row) else "",  # last resort
                ]
                for candidate in candidates:
                    candidate_str = str(candidate).strip()
                    if not is_invalid(candidate_str):
                        description = candidate_str
                        break

            # --- Unit size fallback ---
            unit_size = ""
            if unit_size_col is not None and unit_size_col < len(next_row):
                unit_size = str(next_row.iloc[unit_size_col]).strip()

            # Try parsing from description if not valid
            if not unit_size or unit_size.lower() == "nan":
                words = description.strip().split()
                if len(words) >= 1:
                    last_word = words[-1]
                    try:
                        float(last_word)  # ถ้าแปลงได้ตรงๆ เป็น float
                        unit_size = last_word
                    except ValueError:
                        # ตรวจว่ามีตัวเลขแบบ float อยู่ในคำไหม เช่น "100ml" หรือ "25.5kg"
                        import re
                        match = re.search(r"\d+(\.\d+)?", last_word)
                        if match:
                            unit_size = last_word
                        elif len(words) >= 2:
                            unit_size = f"{words[-2]} {words[-1]}"
                        else:
                            unit_size = ""

            # --- Optional: Clean unit size if it still looks off ---
            if unit_size.lower() == "nan":
                try:
                    unit_size = str(next_row.iloc[unit_size_col + 1]).strip()
                except:
                    unit_size = ""

            # --- Fix description again if it’s just unit size ---
            if description == unit_size:
                description = str(next_row.iloc[2]).strip()

            vat = str(row.iloc[vat_col]).strip() if vat_col is not None else ""
            if vat == "nan":
                vat = str(next_row.iloc[vat_col]).strip() if vat_col is not None else ""
                if not isinstance(vat, str):
                    vat = str(next_row.iloc[vat_col+1]).strip() if vat_col is not None else ""
                if vat == "nan":
                    vat = "EXC"
            if code and description.startswith(code):
                description = description[len(code):].strip()
            vat_percent = str(row.iloc[percent_vat_col]).strip() if percent_vat_col is not None else ""
            pack_size = str(row.iloc[pack_size_col]).strip() if pack_size_col is not None else ""

            if not item_no.isdigit():
                continue

            cleaned_rows.append([
                item_no,
                product,
                code,
                description,
                unit_size,
                unit_qty,
                unit_price,
                vat,
                vat_percent,
                pack_size,
                amount
            ])

    cleaned_df = pd.DataFrame(cleaned_rows, columns=[
        "Item#",
        "Product",
        "Code",
        "Description",
        "UNIT_SIZE_DESCRIPTION",
        "Unit Quantity (Unit/Pack/Case)",
        "Unit Price (USD)",
        "VAT",
        "%VAT",
        "PACK_SIZE",
        "Amount (USD)"
    ])

    return cleaned_df

# === Streamlit UI ===
st.title("📄 แปลง PDF ใบสั่งซื้อเป็น Excel")

uploaded_file = st.file_uploader("📤 อัปโหลดไฟล์ PDF", type="pdf")

if uploaded_file and "last_uploaded" in st.session_state:
    if uploaded_file.name != st.session_state["last_uploaded"]:
        st.session_state.pop("separated_excel", None)
        st.session_state.pop("summary_excel", None)
st.session_state["last_uploaded"] = uploaded_file.name if uploaded_file else None

if uploaded_file and st.button("▶️ เริ่มประมวลผล"):
    with st.spinner("⏳ กำลังประมวลผล..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.read())
            pdf_path = tmp.name

        header_data, summary_data = extract_header_summary(pdf_path)
        tables = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True)

        # === สร้าง Excel แยกตามหน้า ===
        separated_buffer = io.BytesIO()
        with pd.ExcelWriter(separated_buffer, engine="openpyxl") as writer:
            for i, (head, summary) in enumerate(zip(header_data, summary_data)):
                sheet = f"Page{i+1}"
                header_df = pd.DataFrame({"Field": list(head.keys()), "Value": list(head.values())})
                summary_df = pd.DataFrame([summary])
                table = clean_table_flexible(tables[i]) if i < len(tables) else pd.DataFrame()

                header_df.to_excel(writer, sheet_name=sheet, index=False, startrow=0)
                table.to_excel(writer, sheet_name=sheet, index=False, startrow=len(header_df)+2)
                summary_df.to_excel(writer, sheet_name=sheet, index=False, startrow=len(header_df)+len(table)+4)
        separated_buffer.seek(0)
        st.session_state["separated_excel"] = separated_buffer.getvalue()

        # === สร้าง Excel สรุปรวม ===
        combined_rows = []
        for i, head in enumerate(header_data):
            if i < len(tables):
                table = clean_table_flexible(tables[i])
                for _, row in table.iterrows():
                    combined_rows.append({
                        'SHIP_TO': head['Ship To'],
                        'PO_NUMBER': head['PO Number'],
                        'VENDOR_NAME': head['Vendor Name'],
                        'PO_DATE': head['Date'],
                        'DELIVERY_DATE': head['Shipping Date'],
                        'BARCODE': row['Product'],
                        'ITEM_ID': row['Code'],
                        'ITEM_NAME': row['Description'],
                        'UNIT_SIZE_DESCRIPTION': row['UNIT_SIZE_DESCRIPTION'],
                        'UNIT_PRICE': row['Unit Price (USD)'],
                        'VAT': row['VAT'],
                        '%VAT': row['%VAT'],
                        'PACK_SIZE': row['PACK_SIZE'],
                        'UOM_CASE': row['Unit Quantity (Unit/Pack/Case)']
                    })

        summary_buffer = io.BytesIO()
        with pd.ExcelWriter(summary_buffer, engine="openpyxl") as writer:
            final_df = pd.DataFrame(combined_rows)
            final_df.to_excel(writer, sheet_name="Combined Output", index=False)
        summary_buffer.seek(0)
        st.session_state["summary_excel"] = summary_buffer.getvalue()

# === Download buttons appear independently ===
if "separated_excel" in st.session_state and uploaded_file:
    st.download_button(
        "⬇️ ดาวน์โหลด Excel แบบแยกตามหน้า",
        st.session_state["separated_excel"],
        file_name="output_separated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if "summary_excel" in st.session_state and uploaded_file:
    st.download_button(
        "⬇️ ดาวน์โหลด Excel แบบสรุปรวม",
        st.session_state["summary_excel"],
        file_name="output_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def get_column_index_by_keywords(df, keywords):
    for i in range(min(3, len(df))):  # ตรวจแค่ 3 บรรทัดแรก
        for col_idx, cell in enumerate(df.iloc[i]):
            if pd.isna(cell): continue
            text = str(cell).strip().upper()
            for keyword in keywords:
                if keyword.upper() in text:
                    return col_idx
    return None

def extract_tables_with_plumber(pdf_path):
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if len(table) > 1:  # ต้องมี header + data อย่างน้อย
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)
    return all_tables
