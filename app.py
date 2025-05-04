import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
from io import BytesIO

# Extract filled form fields from PDF
def extract_form_fields(uploaded_file):
    reader = PdfReader(uploaded_file)
    fields = reader.get_form_text_fields()
    return fields

# Convert dict to DataFrame
def dict_to_dataframe(fields_dict):
    return pd.DataFrame(fields_dict.items(), columns=["Field", "Value"])

# Streamlit UI
st.title("📄 Extract Filled PDF Form Data")

uploaded_file = st.file_uploader("Upload a filled-out PDF (e.g., IRS Form 1040)", type="pdf")

if uploaded_file:
    try:
        fields = extract_form_fields(uploaded_file)
        if not fields:
            st.warning("No form fields were found or filled in this PDF.")
        else:
            df = dict_to_dataframe(fields)
            st.success("✅ Form fields extracted successfully!")
            st.dataframe(df)

            # Download CSV
            csv_buffer = BytesIO()
            df.to_csv(csv_buffer, index=False)
            st.download_button(
                label="📥 Download as CSV",
                data=csv_buffer.getvalue(),
                file_name="pdf_fields.csv",
                mime="text/csv"
            )

            # Download Excel
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            st.download_button(
                label="📥 Download as Excel",
                data=excel_buffer.getvalue(),
                file_name="pdf_fields.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred: {e}")
