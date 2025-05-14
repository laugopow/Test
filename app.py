import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Test", layout="centered")

st.title("📄 Test v1")
st.markdown("Upload a PDF and extract each section into its own Excel worksheet.")

SECTION_PATTERNS = {
    "Part I": r"Part I\b",
    "Part II": r"Part II\b",
    "Part III": r"Part III\b",
    "Part IV": r"Part IV\b",
    "Part VIII": r"Part VIII\b",
    "Part IX": r"Part IX\b",
    "Part X": r"Part X\b"
}

def extract_sections(file):
    with pdfplumber.open(file) as pdf:
        full_text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    # Define a regex that matches "Part X"
    part_pattern = re.compile(r"(Part\s+(I{1,3}|IV|VIII|IX|X))", re.IGNORECASE)

    # Find all part headers and their positions
    matches = list(part_pattern.finditer(full_text))
    sections = {}

    for i, match in enumerate(matches):
        part_name = match.group(1).strip().title()  # E.g., "Part Ii" => "Part II"
        start_idx = match.start()
        end_idx = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        section_text = full_text[start_idx:end_idx]

        # Extract key-value pairs from this section
        lines = section_text.strip().splitlines()
        key_value_pairs = []
        for line in lines:
            match = re.match(r"^(.+?)\s{2,}([\d,]+\.?|NONE)$", line.strip())
            if match:
                key = match.group(1).strip()
                value = match.group(2).strip()
                key_value_pairs.append((key, value))

        if key_value_pairs:
            df = pd.DataFrame(key_value_pairs, columns=["Field", "Value"])
            sections[part_name] = df

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
