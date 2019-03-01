#!/usr/bin/env python3

import eatools

creds = eatools.get_credentials()
drive = eatools.start_drive(creds)
sheets = eatools.start_sheets(creds)

# Hier wird nun "TESTKOPIE von EA_ALL KNGLMRT 2017" ausgelesen

spreadsheet_id = "10-tp2M3fTWA5jEFl5DouieR-t9K3KOOYtJ9NiKaW6kw"
# sheet_id       = "1527830234"
sheet = sheets.spreadsheets()

# Lese Kopfzeile aus der Tabelle um die Buchstaben zuordnen zu können
result = sheet.values().get(spreadsheetId=spreadsheet_id, range='EA_ALL!A1:1').execute()
values = result.get('values', [])
