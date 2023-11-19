import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Generating the text of file path.
filepaths = glob.glob("invoices/*.xlsx")

# Creating PDF files.
for filepath in filepaths:

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
    pdf.cell(w=50, h=8, txt=f"Dat:{date}",ln=1)

    # Add table to PDF
    # Extracting Excel file.
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Assigning the column name into the pdf.
    columns = list(df.columns)
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Creating cell in pdf doc
    for index ,row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")