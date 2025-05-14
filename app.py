import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="K-3 PDF to Excel", layout="centered")

st.title("📄 Schedule K-3 to Excel Converter")
st.markdown("Upload a K-3 PDF and extract each section into its own Excel worksheet.")

SECTION_HEADERS = {
    "Part I": "Partner's Share of Partnership's Other Current Year International Information Part I",
    "Part II": "Foreign Tax Credit Limitation Part II",
    "Part III": "Other Information for Preparation of Form 1116 or 1118 Part III",
    "Part IV": "Information on Partner's Section 250 Deduction With Respect to Foreign-Derived Intangible Income (FDII) Part IV",
    "Part VIII": "Partner's Interest in Foreign Corporation Income (Section 960) (continued) Part VIII",
    "Part IX": "Partner's Information for Base Erosion and Anti-Abuse Tax (Section 59A) Part IX",
    "Part X": "Foreign Partner's Character and Source of Income and Deductions Part X",
}

def extract_sections(file):
    with pdfplumber.open(file) as pdf:
        full_text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    sections = {}
    for part, header in SECTION_HEADERS.items():
        pattern = rf"({re.escape(header)})(.*?)(?=\n(?:{'|'.join(map(re.escape, SECTION_HEADERS.values()))})|\Z)"
        match = re.search(pattern, full_text, re.DOTALL)
        if match:
            raw_lines = match.group(2).strip().splitlines()
            key_value_pairs = []

            for line in raw_lines:
                match = re.match(r"^(.+?)\s{2,}([\d,]+\.?|NONE)$", line.strip())
                if match:
                    key = match.group(1).strip()
                    value = match.group(2).strip()
                    key_value_pairs.append((key, value))

            if key_value_pairs:
                df = pd.DataFrame(key_value_pairs, columns=["Field", "Value"])
                sections[part] = df
    return sections

def convert_to_excel(sections_dict):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for part, df in sections_dict.items():
            df.to_excel(writer, sheet_name=part[:31], index=False)
    buffer.seek(0)
    return buffer

# Upload widget
uploaded_file = st.file_uploader("Upload a K-3 PDF", type="pdf")

if uploaded_file:
    with st.spinner("Extracting data..."):
        try:
            sections = extract_sections(uploaded_file)
            if not sections:
                st.warning("No relevant data sections found.")
            else:
                excel_data = convert_to_excel(sections)
                st.success("✅ Data extracted successfully!")
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_data,
                    file_name="k3_extracted_sections.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error processing PDF: {e}")
