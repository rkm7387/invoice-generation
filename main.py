import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Generating the text of file path.
filepaths = glob.glob("invoices/*.xlsx")

# Creating PDF files.
for filepath in filepaths:

    # Extracting excel file.
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # PDF file format
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extracting Invoice and Date from the Excel
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Printing Invoice number
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}",ln=1)

    # Printing Date below the invoice number.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Dat:{date}")

    pdf.output(f"PDFs/{filename}.pdf")