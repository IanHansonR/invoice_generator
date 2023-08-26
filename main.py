import pandas as pd
import fpdf
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

# Loop for multiple file generation.
for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split("-")

    pdf = fpdf.FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Invoice number.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=8, txt=f"Invoice # {invoice_nr}", ln=1)
    # Invoice date.
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=8, txt=f"Date: {invoice_date}", ln=1)
    pdf.cell(w=0, h=8, txt=" ", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Table headers.
    table_headers = df.columns
    table_headers = [header.replace("_", " ").title() for header in table_headers]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(h=8, w=30, txt=table_headers[0], border=1)
    pdf.cell(h=8, w=70, txt=table_headers[1], border=1)
    pdf.cell(h=8, w=35, txt=table_headers[2], border=1)
    pdf.cell(h=8, w=30, txt=table_headers[3], border=1)
    pdf.cell(h=8, w=30, txt=table_headers[4], border=1, ln=1)

    # Loop for repetitive table rows.
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(h=8, w=30, txt=str(row['product_id']), border=1)
        pdf.cell(h=8, w=70, txt=str(row['product_name']), border=1)
        pdf.cell(h=8, w=35, txt=str(row['amount_purchased']), border=1)
        pdf.cell(h=8, w=30, txt=str(row['price_per_unit']), border=1)
        pdf.cell(h=8, w=30, txt=str(row['total_price']), border=1, ln=1)

    # Total cost row.
    total_sum = df['total_price'].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(h=8, w=30, txt=" ", border=1)
    pdf.cell(h=8, w=70, txt=" ", border=1)
    pdf.cell(h=8, w=35, txt=" ", border=1)
    pdf.cell(h=8, w=30, txt=" ", border=1)
    pdf.cell(h=8, w=30, txt=str(total_sum), border=1, ln=1)
    pdf.cell(h=8, w=35, txt=" ", ln=1)
    pdf.cell(h=8, w=35, txt=" ", ln=1)

    # Total cost sentence.
    pdf.set_font(family="Times", size=14, style="BU")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(h=8, w=30, txt=f"The total cost is ${total_sum}", ln=1)

    # Company name and logo.
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(h=3, w=0, txt=" ", ln=1)
    pdf.cell(h=8, w=33, txt="PythonHow.com")
    pdf.image("pythonhow.png", w=8, h=8)

    pdf.output(f"PDFs/{filename}.pdf")
