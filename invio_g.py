"""
Created on Sat Jun 27 18:02:37 2026

@author: Kamila Dudzińska

Supporting module with functions

Dataset: Procurement Department 
Contain: Data from SAP Ariba - Material POs
Characteristics: 2500 records,
                 outliers 0,04%, 
                 null values < 0,02%
Goal:   script created for procurement specialist and expert, who want to train 
        data analysis skills in Python/Pandas. The dataset reflects the SAP
        Ariba architecture. 
        
"""

# IMPORT MODULE I
import pandas as pd
import os
import win32com.client
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import random 


# LOAD DATA II
file_path = "C:\\Users\\lila_\\Desktop\\GitHub\\invio\\procurement_mock_dataset_inv.xlsx"
df_ariba = pd.read_excel(file_path, sheet_name="Ariba")
df_invoices = pd.read_excel(file_path, "Invoices")

# MERGE III
df = pd.merge(df_ariba, df_invoices, on="PO Number", how="inner", suffixes=("", "_inv"))

# DATA CLEANING III
df['Requester Name'] = df['Requester Name'].astype(str)
df['PO Number'] = df['PO Number'].astype(str)
df['Order Status'] = df['Order Status'].astype(str)
df['Invoice Status'] = df['Invoice Status'].astype(str)
df['Invoice Amount'] = df['Invoice Amount'].astype(float)
df['Amount'] = df['Amount'].astype(float)


filtered = df[(df['Order Status'] == "received") &
              (df['Invoice Status']== "hold")]

df.columns = df.columns.str.strip()


# OUTLOOK IV
outlook =  win32com.client.Dispatch("Outlook.Application")

#zminenna - liczba mail
emails_sent = 0

# MAIN LOOP V
for index, row in filtered.iterrows():
    status = row['Order Status_inv']
    invoice_status = row['Invoice Status']
    email = row['Requester Mail_inv']
    name = row['Requester Name_inv']
    po_number = row['PO Number']
    amount = row['Amount_inv']
    invoice_amount = row['Invoice Amount']

    below = amount * 0.95
    over = amount * 1.05
    difference = invoice_amount - amount

    if status == 'received' and invoice_status == 'hold':
        if (invoice_amount > below) and (invoice_amount < over) and difference < 20:
            continue
        else:
            mail_item = outlook.CreateItem(0)
            mail_item.To = email
            mail_item.Subject = f"Price mismatch to check {po_number}"
            mail_item.Body = f"""
            Hello {name},

            There is an amount difference on PO {po_number}.
            Order amount: {amount}
            Invoice amount: {invoice_amount}
            Difference: {difference}

            Could you please check and clarify why the difference exceeds tolerance level?
            """
            mail_item.Display()
            emails_sent += 1
            print(f"Email prepared for {email}")

        
# DATA PREPARATION FOR REPORT VI
ordered_count = (df['Order Status_x'] == "ordered").sum()
received_count = (df['Invoice Status'] == "received").sum()
suma = ordered_count + received_count
print(type(suma), type(ordered_count), type(received_count))

# statistics*   
a = ordered_count
b = received_count  
suma = a + b

a = float(a)
b = float(b)
suma = float(suma)

percentage_ordered = divide_z(a, suma) *100
percentage_received = divide_z(b, suma) *100
today = datetime.today().strftime("%d-%m-%Y")
         
# ADMIN PART VII

# create new file
pdf_file = os.path.abspath("C:\\Users\\lila_\\Desktop\\Statistics.pdf")

c=canvas.Canvas(pdf_file, pagesize=A4)

# title
y=720
c.setFont("Helvetica-Bold", 14)
c.setFillColorRGB(0.5, 0.08, 0.5)
c.drawString(100, y, f"Report of PO status")
c.setLineWidth(1)
c.line(100, y-2, 100+250, y-3)

# report
text = c.beginText(100, y-50)
c.setFont("Helvetica", 12)
c.setFillColorRGB(0,0,0)
text.setLeading(30)


text.textLine(f"PO with the status 'ordered' {ordered_count}.")
text.textLine(f"PO with the status 'recieved' {received_count}")
text.textLine(f"PO with the status 'recieved' {emails_sent} which are put on hold due to the price discrepancy")
text.textLine(f"In the report we have {suma} POs")
text.textLine(f"there {ordered_count} so {percentage_ordered:.2f} % PO with the status ordered")
text.textLine(f" and {received_count} so {percentage_received:.2f} % PO with the status received.")
text.textLine(f" Invio sent {today} - - > {emails_sent} emails.")

c.drawText(text)
c.save()

print("Pdf exists: ", pdf_file)   #spr czy plik istnieje

# path to file
pdf_file = os.path.abspath("C:\\Users\\lila_\\Desktop\\Statistics.pdf")

print(os.path.exists(pdf_file))

# email to admin
admin_email = "admin@test.com"
report_mail = outlook.CreateItem(0)
report_mail.To = admin_email
report_mail.Subject = " Daily report of PO status"
report_mail.Body = """ Hello Admin,

here the report of the PO status. Our Python program -Invio has done a great work!
Invio has sent to requestors and asked to check, if GR can be done.
                        
Invio wishes you a great day! :)
                        
                        
  """

#attachment
report_mail.Attachments.Add(pdf_file)

#wysyłka maila
#report_mail.Send()

print(f'Email to admin {admin_email} sent.')

