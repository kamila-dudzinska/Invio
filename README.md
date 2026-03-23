# Invio

Program do usprawnienia procesu Price Missmatch w dziale AP

Autor: Kamila Dudzińska

Projekt: Program 'Invio' do automatyzacji maili 
	 dedykowany dla procesów operacyjnych dla działu zakupów (Procurement)

Źródło: sama wygenerowałam próbkę danych na potrzeby spr. automatyzacji

IDE: Python, Pandas, Excel, Outlook

Cel: Stworzenie programu do analizy tabeli excel z danymi o zamówieniach w systemie CORA (invoicing) oraz automatycznego wysyłania maili do kupców z prośbą o wyjaśnienie różnic kwot netto pomiędzy zamówieniem (PO), a otrzymana fakturą. Program generuje też raport dla administratora, do kogo maile zostały wysłane i jakie są statystyki zamówień. Dzięki temu można jednym kliknięciem zaoszczędzić sporo FTE, a administrator może szybko uzyskać realny "stan rzeczy".


Jak działa program:
1. Program iteruje wiersz po wierszu w tabeli za zamówieniami i porównuje dane z raportem z tabeli z fakturami.
2.Jeśli znajdzie zamówienie (PO) ze statusem "received" ("otrzymane") w raporcie "Ariba" oraz ze statusem
"hold" w tabeli "Faktury" to sprawdzi dodatkowo kwoty netto.
3. Jeżeli różnica kwot netto będzie większa niż 20 EUR lub 5% wartości zamówienia to program wyśle maila do kupca z prośbą o wyjaśnienie różnic.
4. Po wykonaniu zadania program poinformuje administratora, gdzie udało mu się wysłać maila - w przypadku aktywnej konsoli IDE oraz dodatkowo wyśle raport ze statystykami w formacie pdf na maila administratora. 

Zalety projektu:
--> odpowiada na realny problem w wielu procesach operacyjnych, gdzie wymagane jest sprawdzanie i repetetywne wysyłanie przypominajek/follow-upów
--> zmniejsza problem z wyjaśnianiem price missmatch (różnic cenowych) i przyczynia się do redukcji zaległych faktur (invoice overdue) i zminimalizować ryzyko kłopotów z dostawcami, czy utraty wizerunku
--> administrator programu otrzymuje statystyki, dzięki czemu łatwiej kontrolować proces Price Missmatch
--> program automatyzuje pracę w obrębie działu zakupów/AP
--> program napisany pod typowe środowisko korporacyjne z zalogowanym "Outlookiem"
--> program dedykowany SAP, ale można go szybko dopasować do innych systemów - wystarczy przeanalizować raporty generowane przez dowolny inny program.

Kod: cały kod znajduje się w osobnym pliku

Przykładowe fragmenty kodu oraz screen z maila i raportów.

Tabela z zamówieniami:
<img width="1512" height="196" alt="image" src="https://github.com/user-attachments/assets/6024f8b9-1ab1-4cd7-8052-193fe168e2dd" />

Tabela z fakturami:
<img width="1423" height="197" alt="image" src="https://github.com/user-attachments/assets/6ade4b18-40b5-4368-856c-6c36ad083c38" />

Fragmenty kodu:
<img width="621" height="257" alt="image" src="https://github.com/user-attachments/assets/78fd50f9-c48c-4ed1-9b90-2cb651b777c5" />

<img width="788" height="356" alt="image" src="https://github.com/user-attachments/assets/5a1bf282-d0c8-4962-a568-f6e8160cd111" />

Email do buyera:
<img width="888" height="342" alt="image" src="https://github.com/user-attachments/assets/b2ded9d4-fd11-4dc0-b1a2-a11e3bf0254d" />

Statystyki dla administratora:<img width="763" height="458" alt="image" src="https://github.com/user-attachments/assets/1eeee929-e497-48b1-91e9-a46b7bfc5d8a" />


















