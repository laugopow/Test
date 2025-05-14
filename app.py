
import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side
from io import BytesIO
import re

st.set_page_config(page_title="K-3 PDF to Excel", layout="centered")
st.title("📄 Schedule K-3 Visual Extractor (v2)")
st.markdown("Extracts each K-3 part into a separate worksheet, preserving visual layout.")

table_settings = {'vertical_strategy': 'lines', 'horizontal_strategy': 'lines', 'snap_tolerance': 3, 'intersection_tolerance': 5, 'edge_min_length': 3, 'min_words_vertical': 1, 'min_words_horizontal': 1, 'keep_blank_chars': True}

def extract_tables_by_part(pdf_file):
    part_tables = {}
    current_part = "General"
    part_pattern = pd.Series(["Part I", "Part II", "Part III", "Part IV", "Part V", "Part VI", "Part VII",
                              "Part VIII", "Part IX", "Part X", "Part XI", "Part XII", "Part XIII"])

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                for part in part_pattern:
                    if part in line:
                        current_part = part
                        break

            tables = page.extract_tables(table_settings)
            for table in tables:
                if table and len(table) > 1:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    part_tables.setdefault(current_part, []).append(df)
    return part_tables

def write_to_excel(part_tables):
    wb = Workbook()
    wb.remove(wb.active)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for part, tables in part_tables.items():
        ws = wb.create_sheet(title=part[:31])
        for df in tables:
            if df.empty:
                continue
            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=ws.max_row + 1):
                ws.append(row)
                for c_idx, cell in enumerate(ws[r_idx], start=1):
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    cell.border = border
                    if r_idx == 1:
                        cell.font = Font(bold=True)

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

uploaded_file = st.file_uploader("Upload a K-3 PDF", type="pdf")
if uploaded_file:
    with st.spinner("Extracting tables and formatting Excel..."):
        try:
            tables_by_part = extract_tables_by_part(uploaded_file)
            if not tables_by_part:
                st.warning("No tables found.")
            else:
                excel_file = write_to_excel(tables_by_part)
                st.success("✅ Done! Download below.")
                st.download_button("📥 Download Excel File", excel_file, file_name="k3_formatted_output.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"An error occurred: {e}")
