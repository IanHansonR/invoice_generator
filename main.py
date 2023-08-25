import pandas as pd
import fpdf
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath)

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

    pdf.output(f"PDFs/{filename}.pdf")
