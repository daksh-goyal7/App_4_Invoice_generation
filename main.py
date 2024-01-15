import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths=glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf=FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()
    filename=Path(filepath).stem
    invoice_nr,date=filename.split("-")
    pdf.set_font(family="Times",size=12,style="B")
    pdf.cell(w=0,h=12,txt=f"Invoice nr.{invoice_nr}",ln=1)

    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=0, h=12, txt=f"Date:{date}",ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns=df.columns
    columns=[item.replace("_"," ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=50, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=40, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)


    for index,rows in df.iterrows():
        pdf.set_font(family="Times",size=10,style="I")
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30,h=8,txt=str(rows["product_id"]),border=1)
        pdf.cell(w=50, h=8, txt=str(rows["product_name"]),border=1)
        pdf.cell(w=35, h=8, txt=str(rows["amount_purchased"]),border=1)
        pdf.cell(w=40, h=8, txt=str(rows["price_per_unit"]),border=1)
        pdf.cell(w=30, h=8, txt=str(rows["total_price"]),border=1,ln=1)
        total_price=df["total_price"].sum()

    pdf.set_font(family="Times",size=10,style="I")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=f"${total_price}",border=1,ln=1)

    pdf.set_font(family="Times",size=10,style="B")
    pdf.cell(w=0,h=10,txt=f"The total due amount is ${total_price}",ln=1    )

    pdf.set_font(family="Times",size=10,style="UI")
    pdf.cell(w=25,h=10,txt=f"PythonHow")
    pdf.image("pythonhow.png",w=10)

    pdf.output(f"PDFs/{filename}.pdf")