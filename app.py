import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

def extract_text(pdf_file):
    text_lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text_lines.extend(text.split('\n'))
    df = pd.DataFrame(text_lines, columns=["Line"])
    return df

def extract_tables(pdf_file):
    all_tables = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append(df)
    return all_tables

st.title("📄 PDF to CSV/Excel Converter")

uploaded_file = st.file_uploader("Upload a PDF", type="pdf")

if uploaded_file:
    st.write("Choose the data type to extract:")
    mode = st.radio("Extraction Mode", ["Text", "Tables"])

    if st.button("Extract and Download"):
        if mode == "Text":
            df = extract_text(uploaded_file)
            buffer = BytesIO()
            df.to_csv(buffer, index=False)
            st.download_button("Download CSV", data=buffer.getvalue(), file_name="output.csv", mime="text/csv")
        else:
            tables = extract_tables(uploaded_file)
            if not tables:
                st.warning("No tables found in the PDF.")
            else:
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    for i, table_df in enumerate(tables):
                        sheet = f'Table_{i+1}'
                        table_df.to_excel(writer, sheet_name=sheet, index=False)
                st.download_button("Download Excel", data=buffer.getvalue(), file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
