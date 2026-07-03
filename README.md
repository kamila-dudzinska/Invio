# Invio


⚙️ Program „Invio” — Automatyzacja procesu Price Mismatch w dziale AP
Autor: Kamila Dudzińska
Obszar: Procurement / Accounts Payable (AP)
Technologie: Python, Pandas, Outlook, Excel
Moduły: pandas, win32com, datetime, reportlab, random# Invio
Źródło: procurement_mock_dataset_invio.xlsx - stworzony na podstawie własnego skryptu "dataset_mock_invio.py"



🎯 Cel projektu

Program Invio automatyzuje analizę danych zakupowych oraz wysyłkę maili dotyczących różnic kwot netto (Price Mismatch) pomiędzy zamówieniem (PO) a fakturą w systemie CORA/Ariba.
Narzędzie eliminuje konieczność ręcznego sprawdzania setek pozycji i wysyłania follow‑upów do kupców, co pozwala zaoszczędzić znaczną liczbę FTE oraz przyspiesza proces wyjaśniania niezgodności.



💱 Jak działa program: 

Program analizuje tabelę zamówień oraz tabelę faktur, porównując statusy i wartości netto.
1. Program iteruje wiersz po wierszu w tabeli za zamówieniami i porównuje dane z raportem z tabeli z fakturami. 
2. Jeśli znajdzie zamówienie (PO) ze statusem "received" ("otrzymane") w raporcie "Ariba" oraz ze statusem"hold" w tabeli "Faktury" to sprawdzi dodatkowo kwoty netto.
3. Jeżeli różnica kwot netto będzie większa niż 20 EUR lub 5% wartości zamówienia to program wyśle maila do kupca z prośbą o wyjaśnienie różnic. 
4. Po wykonaniu zadania program poinformuje administratora, ile maili zostało wysłanych - w przypadku aktywnej konsoli IDE oraz dodatkowo wyśle raport ze statystykami w formacie pdf na maila administratora.
   


🚀 Zalety projektu:

--> odpowiada na realny problem w wielu procesach operacyjnych, gdzie wymagane jest sprawdzanie i repetetywne wysyłanie przypominajek/follow-upów
--> zmniejsza problem z wyjaśnianiem price missmatch (różnic cenowych) i przyczynia się do redukcji zaległych faktur (invoice overdue) i zminimalizować ryzyko kłopotów z dostawcami, czy utraty wizerunku
--> administrator programu otrzymuje statystyki, dzięki czemu łatwiej kontrolować proces Price Missmatch
--> program automatyzuje pracę w obrębie działu zakupów/AP
--> program napisany pod typowe środowisko korporacyjne z zalogowanym "Outlookiem"
--> program dedykowany SAP, ale można go szybko dopasować do innych systemów - wystarczy przeanalizować raporty generowane przez dowolny inny program.



🗂️ Struktura projektu

-->invio_g.py — główny program automatyzujący analizę i wysyłkę maili
-->dataset_mock_invio.py — generator danych do testów
-->procurement_mock_functions.py — moduł wspierający logikę danych


Przykładowe fragmenty kodu oraz screen z maila i raportów.


Tabela z zamówieniami:
![zamowienia](po_invio.png)

Tabela z fakturami:
![faktury](invoices_invio.png)


Fragmenty kodu:
<img width="621" height="257" alt="image" src="https://github.com/user-attachments/assets/78fd50f9-c48c-4ed1-9b90-2cb651b777c5" />




<img width="788" height="356" alt="image" src="https://github.com/user-attachments/assets/5a1bf282-d0c8-4962-a568-f6e8160cd111" />




Email do buyera:
![buyer email](email_buyer.png)




Email dla administratora:
![admin email](email_admin.png)
















