import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')


    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}', ln=1)

    #pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Date: {date}', ln=1)
    pdf.cell(w=0, h=8, ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')

# titles
    columns = df.columns
    columns = [i.replace('_', ' ').title() for i in columns]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(40, 40, 40)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], ln=1, border=1)

# table
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(40, 40, 40)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), ln=1, border=1)


# sum raw
    sum_price = df['total_price'].sum()
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(w=30, h=8,  border=1)
    pdf.cell(w=60, h=8,  border=1)
    pdf.cell(w=35, h=8,  border=1)
    pdf.cell(w=30, h=8,  border=1)
    pdf.cell(w=30, h=8, txt=str(sum_price), ln=1, border=1)

    pdf.cell(w=0, h=16, ln=1)

# ending
    pdf.set_font(family='Times', size=14, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, txt=f'The total due amount is {sum_price} Euros.', ln=1)
    pdf.cell(w=30, h=8, txt='PythonHow')
    pdf.image('pythonhow.png', w=10)


    pdf.output(f'PDFs/{filename}.pdf')

