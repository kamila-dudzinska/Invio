"""
Created on Sat Jun 27 18:02:37 2026
Extenden on Wed Oct 21  13:59:20 2026

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
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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
DRY_RUN = True          # True = only MAX_PREVIEW email previews, nothing is sent
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
outlook = win32com.client.Dispatch("Outlook.Application") if OUTLOOK_AVAILABLE else None

emails_sent = 0     # number of prepared emails

# MAIN LOOP VII - only POs that need clarification
for _, row in to_clarify.iterrows():
    body = f"""Hello {row['Requester Name']},

There is an amount difference on PO {row['PO Number']}.
Order amount: {row['Amount']:.2f} {row['Currency']}
Invoice amount: {row['Invoice Amount']:.2f} {row['Currency']}
Difference: {row['Difference']:.2f} {row['Currency']} ({row['Difference %']:.1%})

Could you please check and clarify why the difference exceeds the tolerance level?

Thank you in advance!

Have a great day! :)
Invio
"""
    if outlook and (not DRY_RUN or emails_sent < MAX_PREVIEW):
        mail_item = outlook.CreateItem(0)
        mail_item.To = row['Requester Mail']
        mail_item.Subject = f"Price mismatch to check {row['PO Number']}"
        mail_item.Body = body
        if DRY_RUN:
            mail_item.Display()     # test mode: preview only
        else:
            mail_item.Send()
        emails_sent += 1
        print(f"Email created for {row['Requester Mail']} (PO {row['PO Number']})")
    else:
        print(f"[dry run] {row['Requester Mail']} (PO {row['PO Number']})")


# STATISTICS VIII
today = datetime.today()
po_matched = len(df)                          # POs with an invoice
po_on_hold = int(on_hold.sum())               # received + invoice on hold
po_to_clarify = len(to_clarify)               # on hold + out of tolerance
invoice_status_counts = df['Invoice Status'].value_counts()

# overbilling vs underbilling
invoice_above_po = int((to_clarify['Difference'] > 0).sum())
invoice_below_po = int((to_clarify['Difference'] < 0).sum())

# size of the difference (% of PO amount)
diff_size = pd.cut(to_clarify['Difference %'],
                   bins=[0, 0.05, 0.10, 0.25, float('inf')],
                   labels=["up to 5% (above 20 abs)", "5-10%", "10-25%", "over 25%"])
diff_size_counts = diff_size.value_counts().sort_index()

# net difference to clarify per currency (absolute values)
difference_value = (to_clarify.assign(abs_diff=to_clarify['Difference'].abs())
                    .groupby('Currency')['abs_diff'].sum()
                    .sort_values(ascending=False))

# requestors with the most price mismatches
requestors_with_max_mismatches = (to_clarify.groupby('Requester Name').size()
                                  .sort_values(ascending=False)
                                  .head(5))

# ADMIN PART - PDF REPORT IX
# font with Polish letters (Windows Arial), Helvetica as a fallback
FONT, FONT_BOLD = "Helvetica", "Helvetica-Bold"
for regular, bold in [(r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
                      ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                       "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")]:
    if os.path.exists(regular) and os.path.exists(bold):
        pdfmetrics.registerFont(TTFont("ReportFont", regular))
        pdfmetrics.registerFont(TTFont("ReportFont-Bold", bold))
        FONT, FONT_BOLD = "ReportFont", "ReportFont-Bold"
        break

c = canvas.Canvas(PDF_FILE, pagesize=A4)
width, height = A4
PLUM = colors.HexColor("#5B1A5B")     # Invio colour (Mobi is navy)
LILAC = colors.HexColor("#EBD7F0")
GREY = colors.HexColor("#6B7280")


def section_title(y, title):
    c.setFillColor(PLUM)
    c.setFont(FONT_BOLD, 13)
    c.drawString(50, y, title)
    c.setStrokeColor(LILAC)
    c.setLineWidth(2)
    c.line(50, y - 6, width - 50, y - 6)


def draw_bar(y, label, count, share, bar_max=260):
    """One row: label, grey background bar, plum bar for the share, number."""
    c.setFont(FONT, 10)
    c.setFillColor(colors.black)
    c.drawString(50, y + 3, label)
    c.setFillColor(colors.HexColor("#E5E7EB"))
    c.rect(185, y, bar_max, 13, stroke=0, fill=1)
    c.setFillColor(PLUM)
    c.rect(185, y, bar_max * share / 100, 13, stroke=0, fill=1)
    c.setFillColor(colors.black)
    c.drawString(185 + bar_max + 10, y + 3, f"{count:,}  ({share:.1f}%)")


# header bar
c.setFillColor(PLUM)
c.rect(0, height - 90, width, 90, stroke=0, fill=1)
c.setFillColor(colors.white)
c.setFont(FONT_BOLD, 22)
c.drawString(50, height - 50, "Invio - Price Mismatch Report")
c.setFont(FONT, 11)
c.drawString(50, height - 72, f"Generated: {today:%Y-%m-%d %H:%M}")

# KPI tiles
tiles = [("PO with invoice", po_matched),
         ("Received, invoice on hold", po_on_hold),
         ("Out of tolerance", po_to_clarify),
         ("Emails created", emails_sent)]
tile_w, tile_h, gap = 115, 70, 12
y_tiles = height - 190
for i, (label, value) in enumerate(tiles):
    x = 50 + i * (tile_w + gap)
    c.setFillColor(LILAC)
    c.roundRect(x, y_tiles, tile_w, tile_h, 8, stroke=0, fill=1)
    c.setFillColor(PLUM)
    c.setFont(FONT_BOLD, 22)
    c.drawCentredString(x + tile_w / 2, y_tiles + 32, f"{value:,}")
    c.setFont(FONT, 8.5)
    c.setFillColor(GREY)
    c.drawCentredString(x + tile_w / 2, y_tiles + 14, label)

# 1. invoice status - all invoices
y = y_tiles - 40
section_title(y, "Invoice status (all invoices)")
y -= 10
for status, count in invoice_status_counts.items():
    y -= 19
    draw_bar(y, status.capitalize(), count, divide_z(int(count), po_matched) * 100)

# 2. overbilling vs underbilling + size of the difference
y -= 32
section_title(y, "Price mismatch - direction and size")
y -= 10
for label, count in [("Invoice above PO", invoice_above_po),
                     ("Invoice below PO", invoice_below_po)] + list(diff_size_counts.items()):
    y -= 19
    draw_bar(y, label, count, divide_z(int(count), po_to_clarify) * 100)

# 3. difference value per currency
y -= 32
section_title(y, "Net difference to clarify per currency")
y -= 10
for currency, amount in difference_value.items():
    y -= 17
    c.setFont(FONT, 10)
    c.setFillColor(colors.black)
    c.drawString(50, y, currency)
    c.drawRightString(260, y, f"{amount:,.2f}")

# 4. requestors with max mismatches
y -= 32
section_title(y, "Requestors with max mismatches")
y -= 10
for rank, (name, count) in enumerate(requestors_with_max_mismatches.items(), start=1):
    y -= 19
    draw_bar(y, f"{rank}. {name}", count, divide_z(int(count), po_to_clarify) * 100)

# footer
c.setFont(FONT, 8)
c.setFillColor(GREY)
c.drawString(50, 30, f"Rule: PO 'received', invoice 'hold', difference above {TOLERANCE_ABS} "
                     f"or {TOLERANCE_PCT:.0%} of the PO amount. | Invio by Kamila Dudzińska")
c.save()

print(f"PDF report saved: {PDF_FILE}")

# email to admin
if outlook:
    report_mail = outlook.CreateItem(0)
    report_mail.To = ADMIN_EMAIL
    report_mail.Subject = "Daily price mismatch report"
    report_mail.Body = f"""Hello Admin,

here is today's price mismatch report.
Invio found {po_to_clarify} received PO(s) with an invoice on hold and a price
difference above tolerance, and asked the requestors to clarify them.
Details are in the attached PDF.

Invio wishes you a great day! :)
"""
    report_mail.Attachments.Add(PDF_FILE)
    if DRY_RUN:
        report_mail.Display()   # test mode: preview only
    else:
        report_mail.Send()
    print(f"Email to admin {ADMIN_EMAIL} prepared.")
else:
    print("Outlook not available - admin email not created.")
