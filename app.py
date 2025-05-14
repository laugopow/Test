import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Test", layout="centered")

st.title("📄 Test v1.2")
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
    import pdfplumber
    import pandas as pd
    import re

    with pdfplumber.open(file) as pdf:
        full_text = "\n".join(page.extract_text() for page in pdf.pages if page.extract_text())

    lines = full_text.splitlines()
    sections = {"General Data": []}
    current_section = "General Data"

    part_marker = re.compile(r"Part\s+(I{1,3}|IV|VIII|IX|X)\b", re.IGNORECASE)
    
    # 💡 Looser match: any line with a value at the end
    kv_pattern = re.compile(r"^(.{3,}?)\s+([\d,]+\.?|NONE)$")

    debug_hits = []  # store what it matches

    for line in lines:
        line = line.strip()

        # Detect part section headers
        part_match = part_marker.search(line)
        if part_match:
            current_section = f"Part {part_match.group(1).upper()}"
            if current_section not in sections:
                sections[current_section] = []
            continue

        # Try to find a key/value line
        kv_match = kv_pattern.match(line)
        if kv_match:
            key = kv_match.group(1).strip()
            value = kv_match.group(2).strip()
            sections[current_section].append((key, value))
            debug_hits.append(f"{key} → {value}")

    # Debug output to help you verify
    st.subheader("🔍 Key/Value Lines Detected")
    if debug_hits:
        st.text("\n".join(debug_hits[:30]))  # Show the first 30 matches
    else:
        st.warning("No lines matched the label/value pattern.")

    # Convert each group to a DataFrame
    result = {}
    for part, items in sections.items():
        if items:
            df = pd.DataFrame(items, columns=["Field", "Value"])
            result[part] = df

    return result

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
