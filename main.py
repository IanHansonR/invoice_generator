import pandas as pd
import fpdf
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath)
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    pdf = fpdf.FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=0, txt=f"Invoice # {invoice_nr}")
    pdf.output(f"PDFs/{filename}.pdf")
