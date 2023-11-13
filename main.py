from fpdf import FPDF
import pandas as pd
import os

filepath = r"C:\Users\TAAFERU2\Desktop\Python\Invoice Generator\invoice"
files = os.listdir(filepath)
print(files)

for fn in files:
    invoice_no = fn[:5]
    invoice_date = fn[6:15]
    df = pd.read_excel(f'{filepath}/{fn}', sheet_name='Sheet 1')
    print(df)
    for index, row in df.iterrows():
        pdf = FPDF(orientation="P", unit='mm', format='A4')
        pdf.add_page()
        pdf.set_font(family='Times', style='B', size=24)
        pdf.set_auto_page_break(auto=False, margin=0)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(w=0, h=10, txt=f"Invoice nr. {invoice_no}", border=0,
                 ln=1, align='L')
        pdf.cell(w=0, h=10, txt=f"Date {invoice_date}",
                 border=0, ln=1, align='L')

        pdf.output(f'{invoice_no}{invoice_date}.pdf')
