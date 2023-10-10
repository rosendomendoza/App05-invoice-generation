import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    filename = Path(filepath).stem
    nr, date = filename.split("-")

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_auto_page_break(auto=False, margin=0)

    # Set the header Page
    pdf.set_font(family="Times", style="B", size=24)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=f"Invoice nr. {nr}", align="L", ln=1)
    pdf.cell(w=0, h=12, txt=f"Date {date}", align="L", ln=2)

    df = pd.read_excel(f"{filepath}", sheet_name="Sheet 1")

    # Set the header Table
    columns = list(df.columns)
    columns = [column.replace("_"," ").title() for column in columns]
    pdf.set_font(family="Times", style="B", size=14)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=columns[0], align="C", border=1)
    pdf.cell(w=60, h=10, txt=columns[1], align="C", border=1)
    pdf.cell(w=60, h=10, txt=columns[2], align="C", border=1)
    pdf.cell(w=50, h=10, txt=columns[3], align="C", border=1)
    pdf.cell(w=30, h=10, txt=columns[4], align="C", border=1, ln=1)

    pdf.set_font(family="Times", style="", size= 12)

    for index, row in df.iterrows():
        p_id = row["product_id"]
        p_name = row["product_name"]
        amount = row["amount_purchased"]
        price = row["price_per_unit"]
        totalP = row["total_price"]

        # Set the body Table
        pdf.cell(w=30, h=10, txt=str(p_id), align="C", border=1)
        pdf.cell(w=60, h=10, txt=p_name, align="C", border=1)
        pdf.cell(w=60, h=10, txt=str(amount), align="C", border=1)
        pdf.cell(w=50, h=10, txt=str(price), align="C", border=1)
        pdf.cell(w=30, h=10, txt=str(totalP), align="C", border=1, ln=1)


    # Set the footer Table
    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=10, border=1)
    pdf.cell(w=60, h=10, border=1)
    pdf.cell(w=60, h=10, border=1)
    pdf.cell(w=50, h=10, border=1)
    pdf.cell(w=30, h=10, txt=str(total_sum), align="C", border=1, ln=1)

    # Set the footer Page
    pdf.set_font(family="Times", style="B", size=14)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30, h=12, txt=f"The total due amount is {total_sum} Euros.", align="L", ln=1)
    pdf.cell(w=30, h=12, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)
    pdf.output(f"PDFs/{nr}.pdf")