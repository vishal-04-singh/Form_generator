import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
from datetime import datetime
from reportlab.lib.utils import ImageReader

import os

# File paths
excel_path = "p1.xlsx"               # Your Excel file
template_pdf_path = "sample.pdf"         # Your 3-page blank consent form
output_folder = "filled_forms"
os.makedirs(output_folder, exist_ok=True)

# Load Excel
df = pd.read_excel(excel_path)

# Helper function to create overlay for a specific page
def create_overlay(name, email, date, target_page):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    if target_page == 0:
        # Page 1: Red box - Name only
        can.setFont("Helvetica-Bold", 10.9)
        can.drawString(445, 546.5, name)

    elif target_page == 2:
        # Page 3: Name, Email, Date
        can.setFont("Helvetica-Bold", 10.9)
        can.drawString(184, 476.5, name)

        can.setFont("Helvetica", 10.9)
        can.drawString(60, 446.5, email)
        can.drawString(60, 388.5, date)

        # 🖋 Try to place signature image if available
        signature_path = os.path.join("signatures", f"{name}.png")
        if os.path.exists(signature_path):
            try:
                signature = ImageReader(signature_path)
                # Draw signature image (adjust x, y, width, height as needed)
                can.drawImage(signature, 140, 398, width=160, height=50, mask='auto')
            except Exception as e:
                print(f"⚠️ Could not place signature for {name}: {e}")

    can.save()
    packet.seek(0)
    return PdfReader(packet).pages[0]


# Main processing loop
for _, row in df.iterrows():
    name = str(row['Name'])
    email = str(row['Email'])

    # Format date in DD-MM-YYYY (Indian format)
    date_raw = row['Date']
    if pd.notna(date_raw):
        try:
            date = pd.to_datetime(date_raw).strftime('%m-%d-%Y')
        except:
            date = str(date_raw)
    else:
        date = ""

    batch_id = str(row['Unique Batch ID']).strip().replace(" ", "_")

    # Read original PDF
    existing_pdf = PdfReader(open(template_pdf_path, "rb"))
    writer = PdfWriter()

    # Process each page
    for i, page in enumerate(existing_pdf.pages):
        if i == 0:
            overlay = create_overlay(name, email, date, 0)
            page.merge_page(overlay)
        elif i == 2:
            overlay = create_overlay(name, email, date, 2)
            page.merge_page(overlay)

        writer.add_page(page)

    # Save the output file
    output_file = os.path.join(output_folder, f"{batch_id}.pdf")
    with open(output_file, "wb") as f:
        writer.write(f)

print("✅ All consent forms generated successfully with correct page overlays and date format!")
