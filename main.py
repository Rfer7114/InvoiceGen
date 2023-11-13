from fpdf import FPDF
import pandas as pd
import os

# Create a list with the filenames
filepath = r"C:\Users\TAAFERU2\Desktop\Python\Invoice Generator\invoice"
files = os.listdir(filepath)

for fn in files:
    # Get the invoice number and date as separate variables
    invoice_no = fn[:5]
    invoice_date = fn[6:15]

    # Create PDF
    pdf = FPDF(orientation="P", unit='mm', format='A4')
    pdf.add_page()

    pdf.set_font(family='Times', style='B', size=16)
    pdf.set_text_color(0, 0, 0)

    # Write Invoice number and Date to PDF page
    pdf.cell(w=0, h=10, txt=f"Invoice nr. {invoice_no}", border=0,
             ln=1, align='L')

    pdf.cell(w=0, h=10, txt=f"Date {invoice_date}",
             border=0, ln=1, align='L')

    # Read the Excels from invoice directory
    df = pd.read_excel(f'{filepath}/{fn}', sheet_name='Sheet 1')
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family='Helvetica', size=10, style='B')
    pdf.set_text_color(170, 50, 125)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=50, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=40, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=40, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=8)
        pdf.set_text_color(100, 24, 120)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=50, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    pdf.output(f'{invoice_no}{invoice_date}.pdf')
