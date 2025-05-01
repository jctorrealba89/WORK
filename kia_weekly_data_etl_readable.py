import pandas as pd
import logging
from datetime import datetime
import os
log_dir = 'logs'
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, 'kia_data_etl.log')
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
try:
    logging.info('Starting KIA weekly data ETL job...')
    url = 'https://santandernet.sharepoint.com/sites/HCBE_Sales/Freigegebene%20Dokumente/Sales_Operation/Reporting/Quell_Dateien/HCBE_RCD_MainSource.xlsx'
    df = pd.read_excel(url, sheet_name=0)
    df = df[df['Hersteller'] == 'KIA']
    df = df[((df['Produktgattung'] == 'Leasing') | (df['Produkttyp'] == 'Retail Aktion')) & (df['Status'] == 'Active') & ~df['#Händlernr.'].astype(str).str.contains('VA', na=False)]
    columns_to_drop = ['Brand', 'Kundennr.', 'Storno J/N', 'Anfragedatum', 'Fahrgestellnr.', 'Fahrzeugtyp', 'Hersteller', 'Erstzulassung', 'Produktgattung', 'Vertragsart', 'Produkt', 'Kondition', 'Laufzeit', 'WKZ OEM netto', 'Versteckte Sub']
    df.drop(columns=columns_to_drop, errors='ignore', inplace=True)
    df.rename(columns={'Versteckte Sub': 'Hidden Subvention'}, inplace=True)
    df = df.sort_values(by=['Aktivierung', 'Vertragsnummer'], ascending=[True, True])
    df['Hidden Subvention'] = pd.to_numeric(df['Hidden Subvention'], errors='coerce')
    export_dir = 'output'
    os.makedirs(export_dir, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    csv_path = os.path.join(export_dir, f'kia_data_{timestamp}.csv')
    excel_path = os.path.join(export_dir, f'kia_data_{timestamp}.xlsx')
    df.to_csv(csv_path, index=False)
    df.to_excel(excel_path, index=False)
    logging.info(f'ETL job completed successfully. Files saved: {csv_path}, {excel_path}')
except Exception as e:
    logging.error(f'ETL job failed: {e}')
    raise