import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
import zipfile
import os

# --------------------------
# Static Contact Info (editable)
# --------------------------
contact_info = {
    "Contact Name": "Seth Tenore",
    "Phone": "877-780-4848",
    "Fax": "506-675-8989",
    "E-mail": "communicationonlinefiling@avalara.com"
}

# --------------------------
# Extract text from a PDF file
# --------------------------
def extract_pdf_data(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

# --------------------------
# Parse required fields from PDF text
# --------------------------
def parse_data(text):
    lines = text.splitlines()
    data = {}

    for i, line in enumerate(lines):
        if "Company" in line:
            data["Company"] = lines[i+1].strip()
        elif "Filing Period" in line:
            data["Filing Period"] = lines[i+1].strip()
        elif "Form" in line:
            data["Form"] = lines[i+1].strip()
        elif "State" in line:
            data["State"] = lines[i+1].strip()
        elif "Registration ID" in line:
            data["Registration ID"] = lines[i+1].strip()
        elif "Filing Date" in line:
            data["Filing Date"] = lines[i+1].strip()
        elif "Payment Amount" in line:
            data["Payment Amount"] = lines[i+1].strip()

    return data

# --------------------------
# Streamlit UI
# --------------------------
st.set_page_config(page_title="PDF to Excel Batch Converter", layout="centered")
st.title("📄 PDF to Excel Converter")
st.write("Upload multiple Avalara-style confirmation PDFs and download Excel files in a ZIP.")

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    output_zip = io.BytesIO()

    with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
        for uploaded_file in uploaded_files:
            with st.spinner(f"Processing {uploaded_file.name}..."):
                text = extract_pdf_data(uploaded_file.read())
                pdf_data = parse_data(text)
                full_data = {**pdf_data, **contact_info}

                df = pd.DataFrame([full_data])

                # Create Excel in memory
                excel_buffer = io.BytesIO()
                df.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)

                excel_filename = os.path.splitext(uploaded_file.name)[0] + ".xlsx"
                zipf.writestr(excel_filename, excel_buffer.read())

    output_zip.seek(0)
    st.success("✅ All PDFs converted to Excel!")

    st.download_button(
        label="📦 Download ZIP of Excel Files",
        data=output_zip,
        file_name="converted_excels.zip",
        mime="application/zip"
    )
