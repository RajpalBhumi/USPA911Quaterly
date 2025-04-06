import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
import zipfile
import os
from openpyxl import load_workbook
from openpyxl.writer.excel import save_virtual_workbook

# Static Contact Info (SECTION I)
contact_info = {
    "Contact Name": "Seth Tenore",
    "Phone": "877-780-4848",
    "Fax": "506-675-8989",
    "E-mail": "communicationonlinefiling@avalara.com"
}

# PDF Text Extractor
def extract_pdf_text(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = "".join([page.get_text() for page in doc])
    doc.close()
    return text

# Parse PDF to extract needed values
def parse_pdf_data(text):
    lines = text.splitlines()
    data = {}

    for i, line in enumerate(lines):
        if "Company" in line:
            data["Company"] = lines[i + 1].strip()
        elif "Filing Period" in line:
            data["Filing Period"] = lines[i + 1].strip()
        elif "Form" in line:
            data["Form"] = lines[i + 1].strip()
        elif "State" in line:
            data["State"] = lines[i + 1].strip()
        elif "Registration ID" in line:
            data["Registration ID"] = lines[i + 1].strip()
        elif "Filing Date" in line:
            data["Filing Date"] = lines[i + 1].strip()
        elif "Payment Amount" in line:
            data["Payment Amount"] = lines[i + 1].strip().replace("$", "").replace(",", "")
        elif "Unified Communications LLC" in line and "46" in line:
            data["Provider Name"] = "Affiliated Unified"
            data["Federal Tax ID"] = "465746085"
            data["Customer ID"] = "1148455"
        elif "Walter Road" in line:
            data["Address Line 1"] = "358 Walter Road"
        elif "Cochranville" in line:
            data["Address Line 2"] = "Cochranville, PA 19330"
        elif "/2024" in line or "2024" in line:
            if "Period Ending" not in data:
                data["Period Ending"] = "3-31-2024"

    return data

# Fill Excel from template
def fill_excel_template(template_bytes, data_dict):
    wb = load_workbook(filename=template_bytes)
    ws = wb["Remittance Report"]

    # Fill SECTION I cells
    ws["B7"] = data_dict.get("Provider Name", "")
    ws["B8"] = data_dict.get("Federal Tax ID", "")
    ws["B9"] = data_dict.get("Customer ID", "")
    ws["B11"] = data_dict.get("Address Line 1", "")
    ws["B12"] = data_dict.get("Address Line 2", "")

    ws["E7"] = contact_info["Contact Name"]
    ws["E8"] = contact_info["Phone"]
    ws["E9"] = contact_info["Fax"]
    ws["E10"] = contact_info["E-mail"]

    ws["E12"] = data_dict.get("Period Ending", "")
    ws["F13"] = float(data_dict.get("Payment Amount", "0"))

    return wb

# Streamlit App UI
st.set_page_config(page_title="911 Remittance Report Generator", layout="centered")
st.title("📄 Avalara PDF ➜ Branded Excel Generator")
st.caption("Upload multiple Avalara PDFs to auto-fill official Remittance Report format with branding and contact info.")

template_file = "Affiliated Unified 2403 Uniform-911-Surcharge-Remittance-Report.xlsx"
uploaded_files = st.file_uploader("Upload Avalara PDF(s)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    with open(template_file, "rb") as f:
        template_bytes = io.BytesIO(f.read())

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for pdf in uploaded_files:
            text = extract_pdf_text(pdf.read())
            data = parse_pdf_data(text)
            wb = fill_excel_template(template_bytes, data)

            excel_bytes = save_virtual_workbook(wb)
            excel_name = os.path.splitext(pdf.name)[0] + ".xlsx"
            zipf.writestr(excel_name, excel_bytes)

    zip_buffer.seek(0)
    st.success("✅ All branded Excel files generated successfully!")

    st.download_button(
        label="📦 Download All as ZIP",
        data=zip_buffer,
        file_name="911_remittance_excels.zip",
        mime="application/zip"
    )
