# -*- coding: utf-8 -*-
"""
Created on Mon Mar 23 22:09:01 2026

@author: Kamila Dudzińska

"""

#import modułów

import pandas as pd
import os
import win32com.client
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


# wczytywanie pliku excel
file_path = "C:\\Users\\lila_\\Desktop\\Github - moje projekty\\PO_status.xlsx"
df_ariba = pd.read_excel(file_path, sheet_name="Ariba")
df_invoices = pd.read_excel(file_path, "Faktury")

# łączenie danych z obu zakładek
df = pd.merge(df_ariba, df_invoices, on="PO Number", how="inner")

print(df.columns)

filtered = df[(df['Status1'] == "Received") &
              (df['Invoice Status']== "hold")]


print(df.head())

# czyszczenie danych
df['Name'] = df['Name'].astype(str)
df['PO Number'] = df['PO Number'].astype(str)
df['Amount2'] = df['Amount2'].astype(int)


# przygotowanie maila
outlook =  win32com.client.Dispatch("Outlook.Application")

#zminenna - liczba mail
emails_sent = 0

#iteracja po wierszach excela
for index, row in filtered.iterrows():
    
    status = row['Status']
    mail = row['Mail']
    name = row['Name']
    po_number = row['PO Number']
    amount =  row['Amount']
    
    mail = outlook.CreateItem(0)
    
    mail.To = mail
    mail.Subject = f"Price missmatch to check {po_number}"
    
    mail.Body = f"""
    
        Hello {name},
        
        There is an amount differnce on the PO {po_number}. 
        Could you be so kind and check it and clarify why the difference is bigger than tolerance level 
        and inform, if the invoice can be paid?
        
        Thank you in advance!
        
        Kind regards,
        Kamila
        
    
        """

        
    #wysyłanie maila
    mail.Display()
    
    emails_sent +=1
        
    print(f'Email sent to {mail}.')
        


#ta funkcja mogłaby się znależć w osobnym module, gdyby kod był bardziej rozbudowany
def divide_z(a, b, default=0):
    
    """
    Robimy funkcję dzielenia z zabezpieczeniem dzielenia przez zero.
    
    """
    
    try:
        # Sprawdzenie typu danych
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise TypeError("Oba argumenty muszą być liczbami.")
        
        return a / b
    except ZeroDivisionError:
        return default
    except TypeError as e:
        print(f"Błąd: {e}")
        return default


#zmienne do statystyk
ordered_count = (df['Status1'] == "Ordered").sum()
received_count = (df['Status1'] == "Received").sum()
suma = ordered_count + received_count

print(type(suma), type(ordered_count), type(received_count))

# dzisiejszy dzień
today = datetime.today().strftime("%d-%m-%Y")



# obliczenia do statystyk   
# korzystamy ze stworzonej przez nas funkcji 
a = ordered_count
b = received_count  
suma = a + b

a = a.astype(float)
b= b.astype(float)
suma = suma.astype(float)

percentage_ordered = divide_z(a, suma) *100
percentage_received = divide_z(b, suma) *100

         
# czesc adminowa

#tworzenie nowego pliku pdf
pdf_file = os.path.abspath("C:\\Users\\lila_\\Desktop\\Statistics.pdf")

c=canvas.Canvas(pdf_file, pagesize=A4)

# tytuł
y=720
c.setFont("Helvetica-Bold", 14)
c.setFillColorRGB(0.5, 0.08, 0.5)
c.drawString(100, y, f"Report of PO status")
c.setLineWidth(1)
c.line(100, y-2, 100+250, y-3)

#reszta raportu
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

#podajemy cieżkę do pliku
pdf_file = os.path.abspath("C:\\Users\\lila_\\Desktop\\Statistics.pdf")

print(os.path.exists(pdf_file))

# dane administatora
admin_email = "your.email@example.com"
report_mail = outlook.CreateItem(0)
report_mail.To = admin_email
report_mail_Subject = " Daily report of PO status"
report_mail.Body = """ Hello Admin,

here the report of the PO status. Our Python program -Invio has done a great work!
Invio has sent to requestors and asked to check, if GR can be done.
                        
Invio wishes you a great day! :)
                        
                        
  """

#załącznik - plik pdf
report_mail.Attachments.Add(pdf_file)

#wysyłka maila
report_mail.Send()
