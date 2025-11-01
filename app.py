import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
from io import BytesIO
from docx import Document
import pandas as pd

# Title and logo
st.markdown("<h1 style='text-align: center; color: red;'>MDH</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center;'>PDF Converter App</h3>", unsafe_allow_html=True)

# Sidebar help section
with st.sidebar.expander("📘 Help & Instructions"):
    st.markdown("""
    **How to Use:**
    1. Upload a PDF file.
    2. Choose conversion type: PDF to Word or PDF to Excel.
    3. Click Convert and download the result.

    **Supported Formats:**
    - PDF to Word (.docx)
    - PDF to Excel (.xlsx) with table detection

    **Troubleshooting:**
    - Make sure the PDF is not password protected.
    - Tables may not be detected if they are scanned images.

    **Contact:** support@mdhconverter.com
    """)

# Sidebar version history
with st.sidebar.expander("📄 Version History"):
    st.markdown("""
    #### Version 1.1.0 — 2024-04-03
    - Version history section added
    - Improved layout and sidebar navigation
    - Minor bug fixes and performance improvements

    #### Version 1.0.0 — 2024-04-01
    - Initial release with PDF to Word and Excel conversion
    - Table detection for Excel
    - MDH branding and logo
    - Help section added
    """)

# File uploader
uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

# Conversion type
conversion_type = st.radio("Choose conversion type", ["PDF to Word", "PDF to Excel"])

# Convert button
if uploaded_file and st.button("Convert"):
    pdf_bytes = uploaded_file.read()
    pdf_name = uploaded_file.name.rsplit(".", 1)[0]

    if conversion_type == "PDF to Word":
        doc = Document()
        with fitz.open(stream=pdf_bytes, filetype="pdf") as pdf:
            for page in pdf:
                text = page.get_text()
                doc.add_paragraph(text)
        buffer = BytesIO()
        doc.save(buffer)
        st.download_button("Download Word File", buffer.getvalue(), file_name=f"{pdf_name}.docx")

    elif conversion_type == "PDF to Excel":
        xls_buffer = BytesIO()
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            with pd.ExcelWriter(xls_buffer, engine='openpyxl') as writer:
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables):
                        df = pd.DataFrame(table[1:], columns=table[0])
                        sheet_name = f"Page{i+1}_Table{j+1}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
        st.download_button("Download Excel File", xls_buffer.getvalue(), file_name=f"{pdf_name}.xlsx")
