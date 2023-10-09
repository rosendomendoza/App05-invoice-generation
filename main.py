import pandas as pd
from fpdf import FPDF
import glob

filepaths = glob.glob("invoices/*xlsx")

for filepath in filepaths:
    filename = filepath.lstrip("invoices/")[1:]

    nr = filename.split("-")[0]
    date = filename.split("-")[1].rstrip(".xlsx")

    df = pd.read_excel(f"{filepath}", sheet_name="Sheet 1")

    pdf = FPDF(orientation="L", unit="mm", format="A5")

    pdf.add_page()
    pdf.set_auto_page_break(auto=False, margin=0)



    head = "Product ID     Product Name                 Amount     Price per Unit     Total Price"

    # Set the header
    pdf.set_font(family="Times", style="B", size=24)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=f"Invoice nr. {nr}", align="L", ln=1)
    pdf.cell(w=0, h=12, txt=f"Date {date}", align="L", ln=2)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=0, h=12, txt=f"{head}", align="L", ln=1)
    #pdf.line(10, 21, 200, 21)
    print(f"Invoice nr. {nr}")
    print(f"Date {date}")
    print(head)

    totalInv = 0

    pdf.set_font(family="Times", style="", size=12)

    for index, row in df.iterrows():
        p_id = row["product_id"]
        p_name = row["product_name"]
        price =  row["price_per_unit"]
        totalP = row["total_price"]
        amount = totalP // price
        pdf.cell(w=0, h=10, txt=f"  {p_id}        {p_name}          {amount}\
                {price}           {totalP}", align="L", ln=1)
        print(f"  {p_id}        {p_name}          {amount}      {price}           \
            {totalP}")
        totalInv += totalP

    pdf.cell(w=0, h=10, txt=f"                                                    \
            {totalInv}", align="L", ln=1)
    print(f"                                                                       \
        {totalInv}")

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=0, h=12, txt=f"The total due amount is {totalInv} Euros.", align="L", ln=1)
    print()
    print(f"The total due amount is {totalInv} Euros.")

    pdf.output(f"{filepath}.pdf")