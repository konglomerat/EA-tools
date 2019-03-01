#!/usr/bin/env python3

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = [ 'https://www.googleapis.com/auth/drive.metadata.readonly',
           'https://www.googleapis.com/auth/spreadsheets.readonly'   ]

def get_credentials():
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
