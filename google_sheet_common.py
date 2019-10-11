from __future__ import print_function

import warnings

import pandas as pd
from googleapiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools

warnings.filterwarnings('ignore')


def update_sheet(sheetid, range, df):
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    store = file.Storage('storage_spreadsheet_footprint.json')

    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials (1).json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = discovery.build('sheets', 'v4', http=creds.authorize(Http()))
    value_range_body = {'values': df.fillna('').astype(str).values.tolist()}
    request = service.spreadsheets().values().append(spreadsheetId=sheetid, range=range,
                                                     valueInputOption='USER_ENTERED', body=value_range_body).execute()
    newrows = request['updates']['updatedRows']
    return newrows


def getsheet(sheetid, range):
    # Setup the Sheets API
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    store = file.Storage('storage_spreadsheet_footprint.json')

    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials (1).json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = discovery.build('sheets', 'v4', http=creds.authorize(Http()))
    result = service.spreadsheets().values().get(spreadsheetId=sheetid, range=range).execute()
    values = result.get('values', [])
    df1 = pd.DataFrame(values)
    new_header = df1.iloc[0]  # grab the first row for the header
    df1 = df1[1:]  # take the data less the header row
    df1.columns = new_header  # set the header row as the df header
    return df1


def append_sheet(sheetid, range, reclist):
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    store = file.Storage('storage_spreadsheet_footprint.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials (1).json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = discovery.build('sheets', 'v4', http=creds.authorize(Http()))
    value_range_body = {'values': reclist}
    service.spreadsheets().values().append(spreadsheetId=sheetid, range=range,
                                           valueInputOption='USER_ENTERED', body=value_range_body).execute()

