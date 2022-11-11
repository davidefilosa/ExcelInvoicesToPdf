import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    total_sum = df['total_price'].sum()
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', style='B', size=12)
    pdf.set_text_color(100, 100, 100)
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split('-')
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", align='L', ln=1, border=0)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", align='L', ln=1, border=0)
    for column in df.columns:
        pdf.cell(w=40 if column == 'product_name' or column == 'amount_purchased' else 35, h=8, txt=f"{column.replace('_', ' ').title()}", align='L', ln=0, border=1)
    pdf.ln()
    for index, row in df.iterrows():
        pdf.set_font(family='Times', style='I', size=10)
        pdf.set_text_color(80, 80, 80)
        for column in df.columns:
            pdf.cell(w=40 if column == 'product_name' or column == 'amount_purchased' else 35, h=8, txt=f"{row[column]}", align='L', ln=0, border=1)
        pdf.ln()
    for column in df.columns:
        pdf.cell(w=40 if column == 'product_name' or column == 'amount_purchased' else 35,
                 h=8, txt=f"{total_sum if column == 'total_price' else ''}",
                 align='L', ln=0, border=1)
    pdf.ln()
    pdf.set_font(family='Times', style='B', size=10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=40, h=8, txt=f"The total price is {total_sum}",
             align='L', ln=0, border=0)

    pdf.output(f'invoices/{filename}.pdf')
