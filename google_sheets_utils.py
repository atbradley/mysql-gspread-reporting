import os.path
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
import httplib2
import gspread
import yaml
from time import sleep
import sys

import io
from os import SEEK_SET, SEEK_CUR, SEEK_END
import csv

#TODO: Make this a class.

here = os.path.dirname(os.path.realpath(__file__))
settings_file = os.path.join(here, 'ocra-data.conf.yaml')
with open(settings_file, 'r') as f:
    settings = yaml.load(f)

scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    settings['apikeyfile'], scopes=scopes)

service = build('drive', 'v3')
http = credentials.authorize(httplib2.Http())
gc = gspread.authorize(credentials)

discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
            'version=v4')
sheetsClient = build('sheets', 'v4', http=http,
                        discoveryServiceUrl=discoveryUrl)

toc = [['Sheet Name', 'Description']] 

def create_spreadsheet(name, folder):
    body = {
      'mimeType': 'application/vnd.google-apps.spreadsheet',
      'name': name,
      'parents': [folder],
    }
    file = service.files().create(body=body).execute(http=http)
    wkbid = file.get('id')
    wks = gc.open_by_key(wkbid).add_worksheet('Index', 1, 2)
    delete_worksheet(wkbid, 0)

    return wkbid

def create_toc(spreadsheet_id):
    global sheetsClient, toc

    wkb = sheetsClient.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheetid = wkb['sheets'][0]['properties']['sheetId']
    
    csvdata = io.StringIO()
    cw = csv.writer(csvdata)
    cw.writerows(toc)
    
    body = {'requests': [
        {
            'pasteData': { 
                    'coordinate': {
                        'sheetId': sheetid,
                        'rowIndex': 0,
                        'columnIndex': 0,
                    },
                    'data': csvdata.getvalue(),
                    'type': 'PASTE_VALUES',
                    'delimiter': ',',
            }
        }
    ]}
    rsp = sheetsClient.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def data_to_worksheet(spreadsheet_id, name, description, headers, data):
    global service, http, gc, toc
    wkb = gc.open_by_key(spreadsheet_id)
    if not data:
        #Maybe write something to the spreadsheet indicating no data.
        return 0
    
    body = {'requests': [
        {
            'addSheet': {
                'properties': { 
                    'title': name,
                    'gridProperties': {
                        'rowCount': 2,
                        'columnCount': len(headers),
                        'frozenRowCount': 1
                    },
                },
            }
        }
    ]}
    rsp = sheetsClient.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    sheetid = rsp['replies'][0]['addSheet']['properties']['sheetId']
    csvdata = io.StringIO()
    cw = csv.writer(csvdata)
    cw.writerow(headers)
    
    for row in data:
        vals = [z.decode() if type(z)==bytes else '' if z==None else z for z in row]
        cw.writerow(vals)

    body = {'requests': [
        {
            'pasteData': { 
                    'coordinate': {
                        'sheetId': sheetid,
                        'rowIndex': 0,
                        'columnIndex': 0,
                    },
                    'data': csvdata.getvalue(),
                    'type': 'PASTE_VALUES',
                    'delimiter': ',',
            }
        }
    ]}
    rsp = sheetsClient.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    toc.append([name, description])

def freeze_rows(spreadsheet_id, rows=1):
    return
    global sheetsClient
    wkb = gc.open_by_key(spreadsheet_id)

    body = {'requests': []}
    for ws in wkb.worksheets():
        body['requests'].append({
            'updateSheetProperties': {
                'properties': {
                    'sheet_id': ws['properties']['sheetId'],
                    'gridProperties': {
                        'frozenRowCount': 1
                    }
                },
                'fields': 'gridProperties.frozenRowCount'
            }
        })

    rsp = sheetsClient.spreadsheets().batchUpdate(spreadsheetId=fileid, body=body).execute()
        
def delete_worksheet(spreadsheet_id, sheet_index=0):
    gc = gspread.authorize(credentials)
    wkb = gc.open_by_key(spreadsheet_id)
    if sheet_index == 'last':
        sheet_index = len(wkb.worksheets()) - 1
    wkb.del_worksheet(wkb.get_worksheet(sheet_index))
