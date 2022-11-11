import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', style='B', size=12)
    pdf.set_text_color(100, 100, 100)
    filename = Path(filepath).stem
    invoice_nr = filename.split('-')[0]
    invoice_date = filename.split('-')[1]
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", align='L', ln=1, border=0)
    pdf.cell(w=50, h=8, txt=f"Invoice date: {invoice_date}", align='L', ln=1, border=0)
    pdf.output(f'invoices/{filename}.pdf')
