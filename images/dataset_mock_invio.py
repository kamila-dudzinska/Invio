"""
Created on Sat Jun 20 20:00:27 2026

@author: Kamila Dudzińska

Dataset: Procurement Department 
Contain: Data from SAP Ariba - Material POs
Characteristics: 2500 records,
                 outliers 0,04%, 
                 null values: 0,02%
Goal:   script created for procurement specialist and expert, who want to train 
        data analysis skills in Python/Pandas.
"""


# %%
# IMPORT MODULES
import random
from datetime import datetime, timedelta
import openpyxl
import procurement_mock_functions as pmf


# %%
# RANDOM SEED
random.seed(42)

# GENERATED DATA FOR FURTHER ACTIONS
# company codes - created manually basing on my work expierience
company_codes = ["A001", "A002", "B001", "B002", "CH01", "CH02", "D001", "D002", "D003", "D004", "N001", "F001", "F002", "F003", "S001", "S002"]

#created by copilot basing on the rules in attched excel file
suppliers = {
    2004839201: "Adecco sp. z o. o.",
    2001297744: "BluePrint SA",
    2005938472: "Lila",
    2002846619: "Pedro",
    2009183340: "Januszex sp. z o. o.",
    2007745128: "Tech Solutions",
    2006629183: "Green Energy",
    2003374920: "Fast Logistics",
    2008457712: "Alpha Systems",
    2001183499: "Blue Ocean",
    2009927344: "Silverline",
    2004412870: "NextGen",
    2007139044: "Bright Future",
    2005578219: "Global Trade",
    2006291180: "Sunrise Corp",
    2007740033: "Kraftwerk GmbH",
    2008894412: "Bauhaus AG",
    2003319077: "Müller & Söhne",
    2006674921: "Schmidt GmbH",
    2002245190: "Weber AG",
    2009910044: "Fischer GmbH",
    2001138455: "NovaTech",
    2007746621: "EcoSmart",
    2005591188: "Urban Solutions",
    2008823410: "Pioneer Co.",
    2006619923: "Summit Industries",
    2004477129: "Quantum Corp",
    2003390041: "Visionary Ltd.",
    2007745510: "Everest Supplies",
    2006620033: "BlueSky",
    2005579912: "Ironclad",
    2008890021: "Nimbus",
    2003317744: "Crescent",
    2006675519: "Falcon",
    2002246610: "Atlas",
    2009917711: "Vanguard",
    2001139922: "Harbor",
    2007748820: "Legacy",
    2005597741: "Summit",
    2008826612: "Zenith",
    2006614477: "Pinnacle",
    2004475510: "Stratus",
    2003398821: "Nimbus",
    2007741199: "Echo",
    2006627740: "Solstice",
    2005573311: "Aurora",
    2008896612: "Celestial",
    2003315510: "Nimbus",
    2006678821: "Helios",
    2002247744: "Lumen",
    2009915512: "Orion",
    2001137740: "Vortex",
    2007743319: "BlueWave sp. z o. o.",
    2005596610: "IronGate sp. z o. o.",
    2008827741: "ClearWater sp. z o. o.",
    2006615512: "NextLevel sp. z o. o.",
    2004477744: "BrightStar sp. z o. o.",
    2003396611: "Skyline sp. z o. o.",
    2007748822: "EverGreen sp. z o. o.",
    2006623310: "MountainPeak sp. z o. o.",
    2005577741: "Oceanic sp. z o. o.",
    2008895512: "SilverStone sp. z o. o.",
    2003317744: "CrystalClear sp. z o. o.",
    2006676611: "RapidFlow sp. z o. o.",
    2002248822: "TrueNorth sp. z o. o.",
    2009913310: "BlueHorizon sp. z o. o.",
    2001137741: "Sunset sp. z o. o.",
    2007745512: "IronClad sp. z o. o.",
    2005597744: "StormRider sp. z o. o.",
    2008826611: "CloudNine sp. z o. o.",
    2006618822: "BrightPath sp. z o. o.",
    2004473310: "GoldenGate sp. z o. o.",
    2003397741: "NordicTech GmbH",
    2007746612: "Bergmann AG",
    2006625510: "Schneider & Sohn",
    2005578822: "Fischer GmbH",
    2008893310: "Weiss AG",
    2003318821: "Albatros SA",
    2006677744: "Bison SA",
    2002246611: "Cobra SA",
    2009918822: "Delta SA",
    2001135510: "Eagle SA"
}


#created by copilot basing on the rules in attched excel file (module re)
users = [
    {"Requestor_ID": "PLANNMAC", "Name": "Anna Maciejewska", "Mail": "anna.maciejewska@firma.com"},
    {"Requestor_ID": "PLJANNOW", "Name": "Jan Nowak", "Mail": "jan.nowak@firma.com"},
    {"Requestor_ID": "PLEWAZIE", "Name": "Ewa Zielinska", "Mail": "ewa.zielinska@firma.com"},
    {"Requestor_ID": "PLPIWONO", "Name": "Piotr Wozniak", "Mail": "piotr.wozniak@firma.com"},
    {"Requestor_ID": "PLKAMAZU", "Name": "Katarzyna Mazur", "Mail": "katarzyna.mazur@firma.com"},
    {"Requestor_ID": "PLMIWISI", "Name": "Michał Wiśniewski", "Mail": "michal.wisniewski@firma.com"},
    {"Requestor_ID": "PLAGNNO", "Name": "Agnieszka Nowak", "Mail": "agnieszka.nowak@firma.com"},
    {"Requestor_ID": "PLTOZIE", "Name": "Tomasz Zieliński", "Mail": "tomasz.zielinski@firma.com"},
    {"Requestor_ID": "PLMOLEW", "Name": "Monika Lewandowska", "Mail": "monika.lewandowska@firma.com"},
    {"Requestor_ID": "PLPAKAC", "Name": "Paweł Kaczmarek", "Mail": "pawel.kaczmarek@firma.com"},
    {"Requestor_ID": "PLKIWOJ", "Name": "Kinga Wójcik", "Mail": "kinga.wojcik@firma.com"},
    {"Requestor_ID": "PLLUKAM", "Name": "Łukasz Kamiński", "Mail": "lukasz.kaminski@firma.com"},
    {"Requestor_ID": "PLNASZY", "Name": "Natalia Szymańska", "Mail": "natalia.szymanska@firma.com"},
    {"Requestor_ID": "PLJADUD", "Name": "Jakub Duda", "Mail": "jakub.duda@firma.com"},
    {"Requestor_ID": "PLMAPAW", "Name": "Magdalena Pawlak", "Mail": "magdalena.pawlak@firma.com"},
    {"Requestor_ID": "PLMARKRA", "Name": "Marcin Krawczyk", "Mail": "marcin.krawczyk@firma.com"},
    {"Requestor_ID": "PLBANO", "Name": "Barbara Nowicka", "Mail": "barbara.nowicka@firma.com"},
    {"Requestor_ID": "PLGRWRO", "Name": "Grzegorz Wrona", "Mail": "grzegorz.wrona@firma.com"},
    {"Requestor_ID": "PLJOLIS", "Name": "Joanna Lis", "Mail": "joanna.lis@firma.com"},
    {"Requestor_ID": "PLDASZA", "Name": "Dariusz Zając", "Mail": "dariusz.zajac@firma.com"},
]


# percentage of statuses in dataset - basing on my real life expierience
status_choices = ["ordered"] * 30 + ["confirmed"] * 8 + ["received"] * 22 + ["invoiced"] * 27 + ["canceled"] * 3

# percentage of cureency_codes - basing on my real life expierience
currency_choices = ["EUR"] * 60 + ["CHF"] * 12 + ["GBP"] * 8 + ["PLN"] * 20

# percentage of invoice statuses in dataset:
invoice_choices =['entered']*15 + ['vouched']*18 +['hold']*12 + ['pending approval']*6 + ['approved']*8 + ['selected']*5 +['paid']*30 + ['canceled']*2


#amount range % percentage in dataset
amount_ranges = [
    (0, 990, 40),
    (1000, 10000, 26),
    (10001, 20000, 13),
    (20001, 50000, 5),
    (50001, 70000, 9),
    (70001, 80000, 3.02),
    (80001, 1000000, 0.06),  
    (250001, 250001, 0.02)  
]


# DATA CREATION
existing_po = set()             #emoty set
records = []                    #empty lists
existing_inv = set()

start_date = datetime.strptime("01.01.2026", "%d.%m.%Y")
end_date = datetime.strptime("31.05.2026", "%d.%m.%Y")
today = datetime.today()
# do daty początkowej dodajemy randomową liczbę dni z przedziału (data końcowa -data początkowa)
creation_date = start_date + timedelta(days=random.randint(0, (end_date - start_date).days))

#MAIN LOOP
if __name__ =="__main__":
    print("I'm starting to generate data for the SAP Ariba report")
    
    for _ in range(2500):
        po_number = generate_po_number(existing_po)
        existing_po.add(po_number)

        company_code = random.choice(company_codes)
        
        #choose a random.choice() from keys()
        supplier_number = random.choice(list(suppliers.keys()))
        supplier_name = suppliers[supplier_number]
    
        user = random.choice(users)
        
        #from starting date we add the date from range(end_date - start_date)
        creation_date = start_date + timedelta(days=random.randint(0, (end_date - start_date).days))
        
        # to creation_date add a number of days from range (1,30)
        delivery_date = generate_delivery_date(creation_date, today)
        
        order_status = get_order_status(delivery_date, today)
        
        invoice_number = generate_invoice_number(existing_inv)
        existing_inv.add(invoice_number)
        
        #invoice_date add a number of days from range(10,90)
        invoice_date, payment_terms = generate_invoice_date(delivery_date, today)
    
        invoice_status = random.choice(invoice_choices)
        amount = weighted_choice(amount_ranges)
        
        invoice_amount = calculate_invoice_amount(amount, order_status, invoice_status)
        currency = random.choice(currency_choices)
        
        
    
        record = {
            "PO Number": po_number,
            "Company Code": company_code,
            "Supplier ID": supplier_number,
            "Supplier Name": supplier_name,
            "Requester ID": user["Requestor_ID"],
            "Requester Name": user["Name"],
            "Requester Mail": user["Mail"],
            "Order Status": order_status,
            "Create Date": creation_date.strftime("%d.%m.%Y"),
            "Delivery Date": delivery_date.strftime("%d.%m.%Y"),
            'Invoice Date': invoice_date,
            "Invoice Status": invoice_status,
            "Amount": amount,
            "Invoice Amount": invoice_amount,
            "Currency": currency,
            "Payment Terms": payment_terms,
            "Invoice Number": invoice_number 
        }
        records.append(record)
        
    #konwersja listy na obiekt DataFrame
    df_all = pd.DataFrame(records)    
    
    #definiujemy, które kolumny mają ić do której zakładki
    columns_tab1 = ['PO Number', 
                          'Company Code', 
                          'Supplier ID',
                          'Supplier Name',
                          'Requester ID',
                          'Requester Name',
                          'Requester Mail',
                          'Create Date',
                          'Delivery Date',
                          'Order Status',
                          'Amount',
                          'Currency']
    
    
    columns_tab2 = ['PO Number', 
                          'Company Code', 
                          'Supplier ID',
                          'Supplier Name',
                          'Requester ID',
                          'Requester Name',
                          'Requester Mail',
                          'Delivery Date',
                          'Order Status',
                          'Amount',
                          'Invoice Number',
                          'Invoice Date',
                          'Invoice Status',
                          'Invoice Amount',
                          'Currency',
                          'Payment Terms']
    
    #filtrowanie tabeli na 2 podzbiory:
    df_tab1 = df_all[columns_tab1]
    df_tab2 = df_all[columns_tab2]

    # write excel
    with pd.ExcelWriter("procurement_mock_dataset_inv.xlsx", 
                        engine='openpyxl') as writer:
        df_tab1.to_excel(writer,
                         sheet_name='Ariba',
                         index=False)
        df_tab2.to_excel(writer,
                         sheet_name='Invoices',
                         index=False)
    

    
    print("Generated the file procurement_mock_dataset_inv with 2 sheets.")
