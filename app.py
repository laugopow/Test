
import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment

st.set_page_config(page_title="K-3 Structured Extractor", layout="centered")
st.title("📄 Schedule K-3 Structured Data Extractor (Part II)")
st.markdown("Extracts key fields from Part II of Schedule K-3 and outputs clean, structured Excel data.")

column_labels = {
    "a": "U.S. source",
    "b": "Foreign branch category income",
    "c": "Passive category income",
    "d": "General category income",
    "e": "Other (category code)",
    "f": "Sourced by partner",
    "g": "Total"
}

def extract_structured_part_ii(file):
    lines = []
    capturing = False
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "Part II" in text:
                capturing = True
            elif "Part III" in text and capturing:
                break
            if capturing:
                lines.extend(text.splitlines())

    extracted = []
    current_line = None

    for line in lines:
        main_line_match = re.match(r"^(\d{1,2})\s+(.+)", line.strip())
        if main_line_match:
            current_line = main_line_match.group(1)
            continue

        if re.match(r"^[A-C]\s+", line.strip()) and current_line:
            parts = line.strip().split()
            sub_line = parts[0]
            values = re.findall(r"[\d,]+(?:\.\d+)?|NONE", line)

            mapped = {}
            if len(values) == 2:
                mapped["a"] = values[0]
                mapped["g"] = values[1]
            else:
                for idx, val in enumerate(values):
                    if idx < 7:
                        col_key = list(column_labels.keys())[idx]
                        mapped[col_key] = val

            for col_key in column_labels.keys():
                extracted.append({
                    "Field": f"{current_line}{sub_line} ({col_key}) {column_labels[col_key]}",
                    "Value": mapped.get(col_key, "")
                })

    return pd.DataFrame(extracted)

def convert_df_to_excel(df):
    buffer = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Part II Structured"

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        ws.append(row)
        for cell in ws[r_idx]:
            cell.font = Font(bold=(r_idx == 1))
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    for col in ws.columns:
        max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    wb.save(buffer)
    buffer.seek(0)
    return buffer

uploaded_file = st.file_uploader("Upload a Schedule K-3 PDF", type="pdf")
if uploaded_file:
    with st.spinner("Extracting Part II structured data..."):
        df = extract_structured_part_ii(uploaded_file)
        if df.empty:
            st.warning("No structured data found in Part II.")
        else:
            st.success("✅ Extraction complete.")
            st.dataframe(df)
            excel_data = convert_df_to_excel(df)
            st.download_button("📥 Download Excel", data=excel_data, file_name="k3_part_ii_structured.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
