#!/usr/bin/env python3

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# ==============================================================================
# Google API functions
# ------------------------------------------------------------------------------

def get_credentials(SCOPES):
    """Authenticate and return credentials.
    Obtained from https://github.com/gsuitedevs/python-samples/blob/master/drive/quickstart/quickstart.py
    """
    creds = None
    credentials_filename = 'client_id.json' # credentials.json
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_filename, SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds


def start_drive(creds):
    """start the Google Drive service"""
    service = build('drive', 'v3', credentials=creds)
    return service


def start_sheets(creds):
    """start the Google Sheets service"""
    service = build('sheets', 'v4', credentials=creds)
    return service


# ==============================================================================
# Utilities
# ------------------------------------------------------------------------------

def number2letter(number):
    """Convert number to column letter (e.g. 0->"A", 1->"B", 26->"AA")"""
    from string import ascii_uppercase
    letters = []
    while number >= 0:
        letters.append(ascii_uppercase[number % 26])
        number = number // 26 - 1
        return ''.join(reversed(letters))



# ==============================================================================
# Class for things to be done with the EA_ALL table
# ------------------------------------------------------------------------------

class EA_ALL:

    def __init__(self):
        # If modifying these scopes, delete the file token.pickle.
        self.SCOPES = [ 'https://www.googleapis.com/auth/drive.metadata.readonly',
                        'https://www.googleapis.com/auth/spreadsheets.readonly'   ]
        self.creds = get_credentials(self.SCOPES)
        self.drive = start_drive(self.creds)
        self.sheets = start_sheets(self.creds)
        #self.spreadsheet_id = "10-tp2M3fTWA5jEFl5DouieR-t9K3KOOYtJ9NiKaW6kw"  # EA_ALL 2017 Testkopie
        self.spreadsheet_id = "1azciJGKkr3oIEmrFDnMCcQabPz3IeNMHxTJUmoPWbmQ"  # EA_ALL 2018
        #self.sheet_id       = "1527830234"
        self.sheet_name     = "EA_ALL"

    def generate_pdf_of_bookings( self, year=2019, month=2 ):
        """Erzeuge PDF mit allen Buchungen des Buchungszeitraums"""

        # Parameters
        sheets = self.sheets
        column_name_dates = "Datum Überweisung"
        spreadsheet_id = self.spreadsheet_id
        sheet_name     = self.sheet_name

        # Get Table column titles (from the first row)
        sheet = sheets.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=sheet_name+'!A1:1').execute()
        values = result.get('values', [])
        columns = { v:i for i,v in enumerate(values[0])  }

        # Get row numbers and corresponding booking dates
        column = number2letter(columns[column_name_dates])
        rowID0 = 2
        range = '{0}!{1}{2:d}:{1}'.format(sheet_name, column, rowID0)
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range).execute()
        bookingdates = result.get('values', [])
        rowIDs = [ rowID0 + i for i,v in enumerate(bookingdates) ]

        # Filter booking dates to obtain desired rows
        i_select = [ i for i,v in enumerate(bookingdates) if v[0].startswith('{:04d}.{:02d}'.format(year,month)) ]
        ranges = [ '{0}!A{1:d}:{1:d}'.format(sheet_name, rowIDs[i]) for i in i_select ]

        # Get entries for the selected rows
        result = sheet.values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges).execute()
        valueRanges = result.get('valueRanges', [])

        # Loop over the rows and do something
        for valueRange in valueRanges:
            values = valueRange.get('values', [])[0]
            print(valueRange['range'])
            print(values)
            print('')

        print("Jetzt muss noch das PDF erzeugt werden. TODO für Bastian.")

        return


# ==============================================================================
# Some demo functions
# ------------------------------------------------------------------------------

def list_some_files(service):
    # Call the Drive v3 API
    results = service.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))
