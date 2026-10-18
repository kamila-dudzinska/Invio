"""
Created on Sat Jun 27 18:02:37 2026

@author: Kamila Dudzińska

Invio - Price Mismatch automation (PO vs invoice)

Dataset: Procurement Department
Contain: Data from SAP Ariba - Material POs
Characteristics: 2500 records,
                 outliers 0,04%,
                 null values < 0,02%
Goal:   finds POs received in Ariba with invoices on hold, checks the net
        amount difference and asks the buyer to clarify it.
        Run dataset_mock_invio.py first to generate the input file.
"""

# IMPORT MODULE I
import os
from datetime import datetime
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

try:
    import win32com.client
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False   # no Outlook (e.g. Mac/Linux) -> dry run


# SETTINGS
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, "procurement_mock_dataset_inv.xlsx")
PDF_FILE = os.path.join(BASE_DIR, "Statistics.pdf")
ADMIN_EMAIL = "admin@test.com"
TOLERANCE_PCT = 0.05    # 5% of the PO amount
TOLERANCE_ABS = 20      # 20 in PO currency (simplification: no FX rates)
DRY_RUN = True          # True = only print, no Outlook windows               
MAX_PREVIEW = 3         # max. Outlook windows to open in test mode           

def divide_z(a: float, b: float) -> float:
    """Safe division - returns 0 instead of an error when b is 0."""
    return a / b if b else 0.0


# LOAD DATA II
df_ariba = pd.read_excel(FILE_PATH, sheet_name="Ariba")
df_invoices = pd.read_excel(FILE_PATH, sheet_name="Invoices")

# MERGE III
df = pd.merge(df_ariba, df_invoices, on="PO Number", how="inner", suffixes=("", "_inv"))
df.columns = df.columns.str.strip()

# DATA CLEANING IV
df['Requester Name'] = df['Requester Name'].astype(str)
df['PO Number'] = df['PO Number'].astype(str)
df['Order Status'] = df['Order Status'].astype(str)
df['Invoice Status'] = df['Invoice Status'].astype(str)
df['Invoice Amount'] = df['Invoice Amount'].astype(float)
df['Amount'] = df['Amount'].astype(float)

# PRICE MISMATCH CHECK V (vectorised - no row-by-row comparison)
df['Difference'] = (df['Invoice Amount'] - df['Amount']).round(2)
df['Difference %'] = (df['Difference'].abs() / df['Amount']).round(4)

on_hold = (df['Order Status'] == "received") & (df['Invoice Status'] == "hold")
out_of_tolerance = (df['Difference'].abs() > TOLERANCE_ABS) | \
                   (df['Difference %'] > TOLERANCE_PCT)
to_clarify = df[on_hold & out_of_tolerance]

# OUTLOOK VI
outlook = win32com.client.Dispatch("Outlook.Application") if (OUTLOOK_AVAILABLE and not DRY_RUN) else None   

emails_sent = 0     # number of prepared emails

# MAIN LOOP VII - only POs that need clarification
for _, row in to_clarify.iterrows():
    body = f"""Hello {row['Requester Name']},

There is an amount difference on PO {row['PO Number']}.
Order amount: {row['Amount']:.2f} {row['Currency']}
Invoice amount: {row['Invoice Amount']:.2f} {row['Currency']}
Difference: {row['Difference']:.2f} {row['Currency']} ({row['Difference %']:.1%})

Could you please check and clarify why the difference exceeds the tolerance level?
"""
    if outlook and emails_sent < MAX_PREVIEW:                                 
        mail_item = outlook.CreateItem(0)
        mail_item.To = row['Requester Mail']
        mail_item.Subject = f"Price mismatch to check {row['PO Number']}"
        mail_item.Body = body
        mail_item.Display()     # Display = test mode, change to Send() in production
    emails_sent += 1
    print(f"Email prepared for {row['Requester Mail']} (PO {row['PO Number']})")


# DATA PREPARATION FOR REPORT VIII
ordered_count = int((df['Order Status'] == "ordered").sum())
received_count = int((df['Order Status'] == "received").sum())
total = ordered_count + received_count

percentage_ordered = divide_z(ordered_count, total) * 100
percentage_received = divide_z(received_count, total) * 100
today = datetime.today().strftime("%d-%m-%Y")

# ADMIN PART IX
c = canvas.Canvas(PDF_FILE, pagesize=A4)

# title
y = 720
c.setFont("Helvetica-Bold", 14)
c.setFillColorRGB(0.5, 0.08, 0.5)
c.drawString(100, y, "Report of PO status")
c.setLineWidth(1)
c.line(100, y - 2, 100 + 250, y - 3)

# report
text = c.beginText(100, y - 50)
text.setFont("Helvetica", 12)
c.setFillColorRGB(0, 0, 0)
text.setLeading(30)

text.textLine(f"POs with the status 'ordered': {ordered_count}")
text.textLine(f"POs with the status 'received': {received_count}")
text.textLine(f"Received POs on hold with a price mismatch above tolerance: {emails_sent}")
text.textLine(f"POs in the report (ordered + received): {total}")
text.textLine(f"Share of 'ordered': {percentage_ordered:.2f} %")
text.textLine(f"Share of 'received': {percentage_received:.2f} %")
text.textLine(f"On {today} Invio prepared {emails_sent} emails.")

c.drawText(text)
c.save()

print(f"PDF report saved: {PDF_FILE}")

# email to admin
if outlook:
    report_mail = outlook.CreateItem(0)
    report_mail.To = ADMIN_EMAIL
    report_mail.Subject = "Daily report of PO status"
    report_mail.Body = f"""Hello Admin,

here is the report of the PO status. Invio prepared {emails_sent} emails
to requesters and asked them to clarify the price differences.

Invio wishes you a great day! :)
"""
    report_mail.Attachments.Add(PDF_FILE)
    report_mail.Display()       # change to Send() in production
    print(f"Email to admin {ADMIN_EMAIL} prepared.")
else:
    print("Outlook not available - dry run, no emails were created.")
