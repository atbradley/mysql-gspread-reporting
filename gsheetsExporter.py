import os.path
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
import httplib2
import gspread
import yaml
from time import sleep
import sys

import io
import csv

class gsheetsDataExporter:
    def __init__(self, apikeyfile, folder, name, latest=False):
        self.toc = [['Sheet Name', 'Description']] 
        
        scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        
        #TODO: This shouldn't be needed any more--Google's library now uses an environment variable instead.
        # https://cloud.google.com/docs/authentication/getting-started#command-line
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
                                    apikeyfile, scopes=scopes)
        service = build('drive', 'v3')
        http = credentials.authorize(httplib2.Http())

        #gspread client handle
        self.gc = gspread.authorize(credentials)

        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
        
        #Google Sheets API client handle
        self.sheetsClient = build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
        
        body = {
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'name': name,
            'parents': [folder],
        }
        file = service.files().create(body=body).execute(http=http)
        self.wkbid = file.get('id')

        self.gs_wkb = self.gc.open_by_key(self.wkbid)
        self.gs_wkb.add_worksheet('Index', 1, 2)

        self.gs_wkb = self.gc.open_by_key(self.wkbid)
        self.gs_wkb.del_worksheet(self.gs_wkb.sheet1)


        self.wkb = self.sheetsClient.spreadsheets().get(spreadsheetId=self.wkbid).execute()
        self.toc_sheetid = self.wkb['sheets'][0]['properties']['sheetId']

        if latest:
            self.latest_id = latest
            #open latest as self.gs_latest
            self.gs_latest = self.gc.open_by_key(self.latest_id)
            #delete all sheets except 'Index' in latest.
            for sheet in self.gs_latest:
                if 'Index' != sheet.title:
                    body = {'requests': [
                        {
                            'deleteSheet': { 
                                    'sheetId': sheet.id,
                            }
                        }
                    ]}
                    rsp = self.sheetsClient.spreadsheets().batchUpdate(spreadsheetId=self.latest_id, body=body).execute()
                else:
                    #Save self.latest_toc_sheetid
                    self.latest_toc_sheetid = sheet.id
    
    def __enter__(self):
        return self
    
    def __exit__(self, *args):
        self.create_toc()
    
    @staticmethod
    def _list_to_csv(data):
        csvdata = io.StringIO()
        cw = csv.writer(csvdata)
        cw.writerows(data)
        return csvdata.getvalue()

    def create_toc(self):
        def makereq(sheetid, data):
            return {'requests': [
                {
                    'pasteData': { 
                            'coordinate': {
                                'sheetId': sheetid,
                                'rowIndex': 0,
                                'columnIndex': 0,
                            },
                            'data': data,
                            'type': 'PASTE_VALUES',
                            'delimiter': ',',
                    }
                }
            ]}
        body = makereq(self.toc_sheetid, type(self)._list_to_csv(self.toc))
        rsp = self.sheetsClient.spreadsheets().batchUpdate(spreadsheetId=self.wkbid, body=body).execute()
        
        if self.latest_id:
            body = makereq(self.latest_toc_sheetid, type(self)._list_to_csv(self.toc))
            rsplatest = self.sheetsClient.spreadsheets().batchUpdate(spreadsheetId=self.latest_id, body=body).execute()

    def data_to_worksheet(self, name, description, headers, data):
        if not data:
            #Maybe write something to the spreadsheet indicating no data.
            return 0
        
        def makesheet(nm, cols):
            return {'requests': [
            {
                'addSheet': {
                    'properties': { 
                        'title': nm,
                        'gridProperties': {
                            'rowCount': 2,
                            'columnCount': cols,
                            'frozenRowCount': 1
                        },
                    },
                }
            }
        ]}

        body = makesheet(name, len(headers))
        rsp = self.sheetsClient.spreadsheets().batchUpdate(spreadsheetId=self.wkbid, body=body).execute()
        sheetid = rsp['replies'][0]['addSheet']['properties']['sheetId']
        
        #Fix the data.
        data = [headers] + [[z.decode() if type(z)==bytes else '' if z==None else z for z in x] for x in data]

        '''
        inpt = [headers]
        for row in data:
            inpt.append([z.decode() if type(z)==bytes else '' if z==None else z for z in row])
        '''
        
        def insertdata(sid, data):
            return {'requests': [
                    {
                        'pasteData': { 
                                'coordinate': {
                                    'sheetId': sid,
                                    'rowIndex': 0,
                                    'columnIndex': 0,
                                },
                                'data': data,
                                'type': 'PASTE_VALUES',
                                'delimiter': ',',
                        }
                    }
                ]}
        
        body = insertdata(sheetid, type(self)._list_to_csv(data))
        rsp = self.sheetsClient.spreadsheets().batchUpdate(spreadsheetId=self.wkbid, body=body).execute()
        self.toc.append([name, description])

        if self.gs_latest:
            body = makesheet(name, len(headers))
            rsp = self.sheetsClient.spreadsheets().batchUpdate(spreadsheetId=self.latest_id, body=body).execute()
            sheetid = rsp['replies'][0]['addSheet']['properties']['sheetId']

            body =  insertdata(sheetid, type(self)._list_to_csv(data))
            rsp = self.sheetsClient.spreadsheets().batchUpdate(spreadsheetId=self.latest_id, body=body).execute()
            
    def delete_worksheet(self, sheet_index=0):
        if sheet_index == 'last':
            sheet_index = len(self.gs_wkb.worksheets()) - 1
        self.gs_wkb.del_worksheet(self.gs_wkb.get_worksheet(sheet_index))