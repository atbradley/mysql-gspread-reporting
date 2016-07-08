import os.path
from oauth2client.service_account import ServiceAccountCredentials
from apiclient.discovery import build
import httplib2
import gspread
import yaml
from time import sleep
import sys

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

# maybe create a folder and upload CSVs as individual spreadsheet files.
# Folders are mimeType application/vnd.google-apps.folder.

def create_spreadsheet(name, folder):
    body = {
      'mimeType': 'application/vnd.google-apps.spreadsheet',
      'name': name,
      'parents': [folder],
    }
    file = service.files().create(body=body).execute(http=http)
    wkbid = file.get('id')
    
    wks = gc.open_by_key(wkbid).add_worksheet('Index', 1, 2)
    wks.update_cell(1,1, 'Sheet Name')
    wks.update_cell(1,2, 'Description')
    
    delete_worksheet(wkbid, 0)

    return wkbid

def data_to_worksheet(spreadsheet_id, name, description, headers, data):
    global service, http, gc
    wkb = gc.open_by_key(spreadsheet_id)
    if not data:
        #Maybe write something to the spreadsheet indicating no data.
        return 0
    
    wksindex = wkb.sheet1
    wksindex.append_row((name, description))
    
    wks = wkb.add_worksheet(name, 1, len(headers))

    for hc in range(len(headers)): 
        wks.update_cell(1, hc+1, headers[hc])
        
    print(str(len(data)))    
    
    for row in data:
        vals = [z.decode() if type(z)==bytes else '' if z==None else z for z in row]
        print("Appending row: "+str(vals))
        try:
            wks.append_row(vals)
        except gspread.exceptions.HTTPError as ex:
            #TODO: Real logging
            print("HTTP Error. Trying to reauthenticate.")
            print(str(ex))
            
            #Pause a few seconds, reauthenticate and try again. Go ahead and crash if we fail again.
            sleep(30)
            service = build('drive', 'v3')
            http = credentials.authorize(httplib2.Http())
            gc = gspread.authorize(credentials)
            
            print("Reauthenticated. Trying to open spreadsheet.")
            wkb = gc.open_by_key(spreadsheet_id)
            wks = wkb.worksheet(name)
            print("adding values.")

            wks.append_row(vals)
        sys.stdout.flush()
        
def delete_worksheet(spreadsheet_id, sheet_index=0):
    gc = gspread.authorize(credentials)
    wkb = gc.open_by_key(spreadsheet_id)
    wkb.del_worksheet(wkb.get_worksheet(sheet_index))
    
#notes:

#Getting started
#https://developers.google.com/gdata/articles/python_client_lib#running-tests-and-samples

#Create an empty spreadsheet and get the key
#http://stackoverflow.com/questions/10527705/create-new-spreadsheet-google-api-python
#or maybe http://stackoverflow.com/questions/12741303/create-a-empty-spreadsheet-in-google-drive-using-google-spreadsheet-api-python

#Authentication using a service account
#https://developers.google.com/identity/protocols/OAuth2ServiceAccount#authorizingrequests
#http://stackoverflow.com/questions/16026286/using-oauth2-with-service-account-on-gdata-in-python
#http://gclient-service-account-auth.readthedocs.org/en/latest/