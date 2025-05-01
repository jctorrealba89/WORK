
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import win32com.client as win32

# 1. Setup Paths
documents_folder = "C://Users//h409422//OneDrive - Santander Office 365//Documents//"
forecast_folder = "C://Users//h409422//OneDrive - Santander Office 365//"
sharepoint_folder = "C://Users//h409422//Santander Office 365//Sales - Quell_Dateien/"
rcd_folder = "C://Users/h409422/Santander Office 365/HCBE Dashboard Retail - Retail Contract Details via Databricks/"
os.chdir(documents_folder)

# 2. Find Most Recent RCD File
def find_rcd_file():
    files = sorted(os.listdir(rcd_folder), reverse=True)
    files = [f for f in files if all(x not in f for x in ["KMD", "HMD", "desktop"])]
    files = [f for f in files if all(x in f for x in [".xlsx", "Retail", "2025"])]
    target_month = (datetime.today() - timedelta(days=90)).month
    files = [f for f in files if f"_{target_month:02}" in f]
    return os.path.join(rcd_folder, files[0]) if files else None

rcd_file_path = find_rcd_file()

# 3. Load Data
df = pd.read_excel(rcd_file_path)
df.rename(columns={
    "StornoJ/N": "Storno J/N",
    "Fahrgestellnr": "Fahrgestellnr.",
    "Rabattierfähiger Betrag Netto": "Rabattierfähiger Betrag netto",
    "Fin Betrag brutto": "Fin. Betrag brutto",
    "Fin Betrag netto": "Fin. Betrag netto",
    "1 Rate brutto": "1. Rate brutto",
    "1 Rate netto": "1. Rate netto",
    "WKZ Hdl netto": "WKZ Hdl. netto",
    "Händler Nr": "Händlernr.",
    "Händlernr": "Händler Nr."
}, inplace=True)

# 4. Format Dates
for col in ["Anfragedatum", "Aktivierung", "Erstzulassung"]:
    df[col] = pd.to_datetime(df[col], errors='coerce')

# 5. Fill Missing Premiums with 0
premium_cols = [
    "RSV - Prämie", "GAP - Prämie", "MOBI PLUS - Prämie", "EW - Prämie",
    "Rabattierter Kaufpreis brutto", "Rabattierter Kaufpreis netto",
    "Anzahlung in EUR brutto", "Anzahlung in EUR netto"
]
df[premium_cols] = df[premium_cols].fillna(0)

# 6. Calculate Finb1 and Finb2
df["Finb1"] = np.where(
    df["Produktgattung"] == "Finanzierung",
    df["Rabattierter Kaufpreis brutto"] - df["Anzahlung in EUR brutto"],
    df["Rabattierter Kaufpreis netto"] - df["Anzahlung in EUR netto"]
)

df["Finb2"] = np.where(
    df["Produktgattung"] == "Finanzierung",
    df["Finb1"] + df["RSV - Prämie"] + df["GAP - Prämie"] + df["MOBI PLUS - Prämie"] + df["EW - Prämie"],
    df["Finb1"] + df["RSV - Prämie"] + df["GAP - Prämie"] + df["MOBI PLUS - Prämie"] + df["EW - Prämie"]
)

# Continue with enrichment, export, and emailing logic...
