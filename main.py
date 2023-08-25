import pandas as pd
import fpdf
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split("-")

    pdf = fpdf.FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Invoice number
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=8, txt=f"Invoice # {invoice_nr}", ln=1)
    # Invoice date
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=8, txt=f"Date: {invoice_date}", ln=1)
    pdf.cell(w=0, h=8, txt=" ", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Table headers
    table_headers = df.columns
    table_headers = [header.replace("_", " ").title() for header in table_headers]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(h=8, w=30, txt=table_headers[0], border=1)
    pdf.cell(h=8, w=70, txt=table_headers[1], border=1)
    pdf.cell(h=8, w=35, txt=table_headers[2], border=1)
    pdf.cell(h=8, w=30, txt=table_headers[3], border=1)
    pdf.cell(h=8, w=30, txt=table_headers[4], border=1, ln=1)

    # Table data
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(h=8, w=30, txt=str(row['product_id']), border=1)
        pdf.cell(h=8, w=70, txt=str(row['product_name']), border=1)
        pdf.cell(h=8, w=35, txt=str(row['amount_purchased']), border=1)
        pdf.cell(h=8, w=30, txt=str(row['price_per_unit']), border=1)
        pdf.cell(h=8, w=30, txt=str(row['total_price']), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
