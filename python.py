import pandas as pd
from fpdf import FPDF
import os


excel_file = "recipients.xlsx"
output_folder = "output_letters"
os.makedirs(output_folder, exist_ok=True)


template = """
Dear [Name],

We hope this message finds you well.

This is a reminder that your current balance is [Balance] and is due by [DueDate].

Address: [Address]

Please make your payment at your earliest convenience.

Sincerely,
Your Company
"""


df = pd.read_excel(excel_file)


for index, row in df.iterrows():
    content = template
    for col in df.columns:
        content = content.replace(f"[{col}]", str(row[col]))

   
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for line in content.strip().split('\n'):
        pdf.multi_cell(0, 10, line)

 
    filename = f"{output_folder}/{row['Name'].replace(' ', '_')}.pdf"
    pdf.output(filename)

print("✅ file  ", output_folder , "created")
