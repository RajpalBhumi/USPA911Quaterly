import streamlit as st
import pandas as pd
import pdfplumber
import io
import zipfile
import os
import re
from datetime import date
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage

# SECTION I Contact Info (Fixed)
contact_info = {
    "Contact Name": "Seth Tenore",
    "Phone": "877-780-4848",
    "Fax": "506-675-8989",
    "E-mail": "communicationonlinefiling@avalara.com"
}

# Extract Filing Period, Payment, and Address
def extract_basic_data(text):
    lines = text.splitlines()
    data = {}

    for i, line in enumerate(lines):
        if "Filing Period" in line:
            data["Period Ending"] = lines[i + 1].strip()
        elif "Payment Amount" in line:
            data["Payment Amount"] = lines[i + 1].strip()
        elif "Walter Road" in line:
            data["Address Line 1"] = "358 Walter Road"
        elif "Cochranville" in line:
            data["Address Line 2"] = "Cochranville, PA 19330"
        elif "Unified Communications LLC" in line and "46" in line:
            data["Provider Name"] = "Affiliated Unified"
            data["Federal Tax ID"] = "465746085"
            data["Customer ID"] = "1148455"
    return data

# Extract Section II and III values using pdfplumber
def extract_surcharge_rows_pdfplumber(file):
    months = {
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    }

    result = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 3:
                        continue
                    row_clean = [str(cell).strip() if cell else "" for cell in row]
                    if row_clean[0] in months:
                        try:
                            assessed = float(row_clean[1].replace(",", "").replace("$", ""))
                            collected = float(row_clean[2].replace(",", "").replace("$", ""))
                            result.append({
                                "month": row_clean[0],
                                "assessed": assessed,
                                "collected": collected
                            })
                        except:
                            continue
    return result

# Fill Excel template
def fill_excel_template(template_bytes, data_dict, section_v_data, surcharge_rows):
    wb = load_workbook(filename=template_bytes)
    ws = wb["Remittance Report"]

    # Unlock sheet
    ws.protection.sheet = False

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

    # SECTION V – Signature
    try:
        for r in ["B41:D41", "E41:F41", "B43:D43"]:
            if r in [str(rng) for rng in ws.merged_cells.ranges]:
                ws.unmerge_cells(r)
        ws["B41"] = section_v_data["initials"]
        ws["D41"] = section_v_data["title"]
        ws["F41"] = section_v_data["date"]
        ws["B43"] = section_v_data["full_name"]
    except Exception as e:
        print("Section V Error:", e)

    # SECTION II – Start at row 17
    start_row_ii = 17
    for i, row in enumerate(surcharge_rows[:3]):
        ws[f"B{start_row_ii + i}"] = row["month"]
        ws[f"C{start_row_ii + i}"] = row["assessed"]
        ws[f"D{start_row_ii + i}"] = row["collected"]

    # SECTION III – Start at row 30
    start_row_iii = 30
    for i, row in enumerate(surcharge_rows[3:6]):
        ws[f"B{start_row_iii + i}"] = row["month"]
        ws[f"C{start_row_iii + i}"] = row["assessed"]
        ws[f"D{start_row_iii + i}"] = row["collected"]

    # Logo
    try:
        logo = ExcelImage("logo.png")
        logo.width = 150
        logo.height = 50
        ws.add_image(logo, "B1")
    except FileNotFoundError:
        print("logo.png not found")

    return wb

# Streamlit UI
st.set_page_config(page_title="911 Remittance Excel Generator", layout="centered")
st.title("📄 Avalara PDF ➝ Branded Excel Report Generator")

# Section V inputs
st.sidebar.header("✍️ Section V – Certification Info")
initials = st.sidebar.text_input("Initials", "Rhenry")
title = st.sidebar.text_input("Title", "Sr Tax Analyst")
full_name = st.sidebar.text_input("Full Name", "Rachel Henry")
cert_date = st.sidebar.date_input("Date", value=date.today())

section_v_data = {
    "initials": initials,
    "title": title,
    "full_name": full_name,
    "date": cert_date.strftime("%-m/%-d/%Y")
}

template_file = "Template Report.xlsx"
uploaded_files = st.file_uploader("Upload Avalara Confirmation PDF(s)", type="pdf", accept_multiple_files=True)

if uploaded_files:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for pdf in uploaded_files:
            with st.spinner(f"Processing {pdf.name}..."):
                try:
                    with pdfplumber.open(pdf) as pdf_doc:
                        full_text = "\n".join([page.extract_text() for page in pdf_doc.pages if page.extract_text()])
                    basic_data = extract_basic_data(full_text)
                    surcharge_rows = extract_surcharge_rows_pdfplumber(pdf)

                    with open(template_file, "rb") as f:
                        template_bytes = io.BytesIO(f.read())

                    wb = fill_excel_template(template_bytes, basic_data, section_v_data, surcharge_rows)

                    excel_io = io.BytesIO()
                    wb.save(excel_io)
                    excel_io.seek(0)
                    zipf.writestr(os.path.splitext(pdf.name)[0] + ".xlsx", excel_io.read())
                except Exception as e:
                    st.error(f"Error processing {pdf.name}: {e}")

    zip_buffer.seek(0)
    st.success("✅ All files processed.")
    st.download_button("📦 Download All Reports (ZIP)", data=zip_buffer, file_name="remittance_reports.zip", mime="application/zip")
