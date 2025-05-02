
import os
import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32

# === CONFIGURATION ===
documents_folder = "C://Users//h409422//OneDrive - Santander Office 365//Documents//"
sharepoint_folder = "C://Users//h409422//Santander Office 365//Sales - Quell_Dateien//"
rcd_folder = "C://Users//h409422//Santander Office 365//HCBE Dashboard Retail - Retail Contract Details via Databricks//"

# === LOAD MOST RECENT FILE ===
target_month = f"_0{(datetime.today() - timedelta(days=90)).month}"
files = sorted([f for f in os.listdir(rcd_folder) if all([
    f.endswith(".xlsx"),
    "Retail" in f,
    "2025" in f,
    target_month in f,
    not any(x in f for x in ["KMD", "HMD", "desktop"])
])], reverse=True)
rcd_file = os.path.join(rcd_folder, files[0])
df = pd.read_excel(rcd_file)

# === RENAME COLUMNS ===
df.rename(columns={
    "StornoJ/N": "Storno J/N",
    "Fahrgestellnr": "Fahrgestellnr.",
    "Rabattierfähiger Betrag Netto": "Rabattierfähiger Betrag netto",
    "Fin Betrag brutto": "Fin. Betrag brutto",
    "Fin Betrag netto": "Fin. Betrag netto",
    "1 Rate brutto": "1. Rate brutto",
    "1 Rate netto": "1. Rate netto",
    "WKZ Hdl netto": "WKZ Hdl. netto",
    "Händler Nr": "Händlernr"
}, inplace=True)

# === CLEAN DATES ===
for col in ["Anfragedatum", "Aktivierung", "Erstzulassung"]:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

# === FILL NAs ===
na_cols = [
    "RSV - Prämie", "GAP - Prämie", "MOBI PLUS - Prämie", "EW - Prämie",
    "Rabattierter Kaufpreis brutto", "Rabattierter Kaufpreis netto",
    "Anzahlung in EUR brutto", "Anzahlung in EUR netto"
]
df[na_cols] = df[na_cols].fillna(0)

# === CALCULATE Finb1 & Finb2 ===
is_finance = df["Produktgattung"] == "Finanzierung"
df["Finb1"] = df["Rabattierter Kaufpreis brutto"].where(is_finance, df["Rabattierter Kaufpreis netto"]) -               df["Anzahlung in EUR brutto"].where(is_finance, df["Anzahlung in EUR netto"])
df["Finb2"] = df["Finb1"] + df[["RSV - Prämie", "GAP - Prämie", "MOBI PLUS - Prämie", "EW - Prämie"]].sum(axis=1)

# === FILL MISSING HÄNDLERNR ===
zuordnung_file = os.path.join(sharepoint_folder, "00_Zuordnung_Hdl_Nr_Vertragsnr.xlsx")
zuordnung = pd.read_excel(zuordnung_file, sheet_name="Sheet1")
zuordnung["Vertragsnummer"] = zuordnung["Angebots-nummer"]
df = df.merge(zuordnung[["Vertragsnummer", "HDL Nummer"]], on="Vertragsnummer", how="left")
df["Händlernr"] = df["Händlernr"].combine_first(df["HDL Nummer"])
df.drop(columns=["HDL Nummer"], inplace=True)

# === FILL MISSING KRAFTSTOFFART ===
df["Kraftstoffart"] = df["Kraftstoffart"].replace("NA", "").fillna("")
fuel_map = df[df["Kraftstoffart"] != ""].drop_duplicates("Fahrzeugmodell")[["Fahrzeugmodell", "Kraftstoffart"]]
df = df.merge(fuel_map, on="Fahrzeugmodell", how="left", suffixes=("", "_y"))
df["Kraftstoffart"] = df["Kraftstoffart"].combine_first(df["Kraftstoffart_y"]).replace("", "Benzin")
df.drop(columns=["Kraftstoffart_y"], inplace=True)

# === QUALITY CHECK SUMMARY ===
check = {
    "Empty_vertragsnummer": df["Vertragsnummer"].isna().sum(),
    "Empty_Status": df["Status"].isna().sum(),
    "Empty_Handlernr": df["Händlernr"].isna().sum(),
    "Empty_Kraftstoffart": (df["Kraftstoffart"] == "").sum(),
    "Empty_Dealer_Codes": df["Händlernr"].isna().sum(),
    "Empty_Fahrzeugtyp": df["Fahrzeugtyp"].isna().sum(),
    "Wrong_Dealer_Code": df["Händlernr"].fillna("").str[:3].ne("C07").sum()
}
check_df = pd.DataFrame([check])

# === SAVE FILES ===
main_output = os.path.join(sharepoint_folder, "HCBE_RCD_Mainsource.xlsx")
check_output = os.path.join(documents_folder, "RCD_Check.xlsx")
df.to_excel(main_output, sheet_name="Export", index=False)
check_df.to_excel(check_output, index=False)

# === SEND EMAIL ===
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = "juan.torrealba@de.hcs.com; anastasia.taranuha@de.hcs.com; philipp.hedwig@de.hcs.com"
mail.Subject = f"RCD_Check_{datetime.today().date()}"
mail.Body = "This is an automatic email: With the weekly check on RCD."
mail.Attachments.Add(check_output)
mail.Send()

print("✅ Done. Email sent and data saved.")
