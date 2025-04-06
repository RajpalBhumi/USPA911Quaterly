import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import io
import zipfile
import os
import re
from datetime import date
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage

# SECTION I Contact Info (Static)
contact_info = {
    "Contact Name": "Seth Tenore",
    "Phone": "877-780-4848",
    "Fax": "506-675-8989",
    "E-mail": "communicationonlinefiling@avalara.com"
}

# Extract text from PDF
def extract_pdf_text(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = "".join([page.get_text() for page in doc])
    doc.close()
    return text

# Parse required fields from text
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
            data["Payment Amount"] = lines[i + 1].strip()
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

# Fill Excel using correct cell anchors and optional logo
def fill_excel_template(template_bytes, data_dict, section_v_data):
    wb = load_workbook(filename=template_bytes)
    ws = wb["Remittance Report"]

    # SECTION I
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

    # Payment
    payment_raw = data_dict.get("Payment Amount", "")
    match = re.search(r"[\d,]+\.\d{2}", payment_raw)
    ws["F13"] = float(match.group(0).replace(",", "")) if match else 0.0

    # SECTION V - Certification (correct anchor cells)
    ws["B41"] = section_v_data["initials"]
    ws["D41"] = section_v_data["title"]
    ws["F41"] = section_v_data["date"]
    ws["B43"] = section_v_data["full_name"]

    # Re-insert logo if it was removed during open/save
    try:
        logo = ExcelImage("logo.png")  # Make sure this file exists
        logo.width = 150
        logo.height = 50
        ws.add_image(logo, "G1")
    except FileNotFoundError:
        print("⚠️ Logo not found: logo.png — skipping image.")

    return wb

# Streamlit UI
st.set_page_config(page_title="911 Remittance Excel Generator", layout="centered")
st.title("📄 Avalara PDF ➝ Branded Excel Report Generator")
st.caption("Upload Avalara confirmation PDFs to auto-fill branded Excel remittance reports.")

# Sidebar inputs for Section V
st.sidebar.header("✍️ Section V – Certification Info")
initials = st.sidebar.text_input("Initials", "Rhenry")
title = st.sidebar.text_input("Title", "Sr Tax Analyst")
full_name = st.sidebar.text_input("Full Name", "Rachel Henry")
cert_date = st.sidebar.date_input("Date", value=date.today())

section_v_data = {
    "initials": initials,
    "title": title,
    "full_name": full_name,
    "date": cert_date.strftime("%-m/%-d/%Y")  # e.g., 4/15/2024
}

# Template and PDF uploads
template_file = "Template Report.xlsx"
uploaded_files = st.file_uploader("Upload Avalara PDF(s)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for pdf in uploaded_files:
            with st.spinner(f"Processing {pdf.name}..."):
                try:
                    text = extract_pdf_text(pdf.read())
                    extracted_data = parse_pdf_data(text)

                    with open(template_file, "rb") as f:
                        template_bytes = io.BytesIO(f.read())  # fresh copy each time

                    wb = fill_excel_template(template_bytes, extracted_data, section_v_data)

                    excel_buffer = io.BytesIO()
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)

                    file_name = os.path.splitext(pdf.name)[0] + ".xlsx"
                    zipf.writestr(file_name, excel_buffer.read())
                except Exception as e:
                    st.error(f"❌ Error processing {pdf.name}: {e}")

    zip_buffer.seek(0)
    st.success("✅ All Excel files generated!")

    st.download_button(
        label="📦 Download ZIP of Reports",
        data=zip_buffer,
        file_name="911_remittance_reports.zip",
        mime="application/zip"
    )
