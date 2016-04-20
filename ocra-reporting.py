import os.path
from datetime import date

import _mysql
import yaml

from google_sheets_utils import create_spreadsheet, data_to_worksheet, delete_worksheet

here = os.path.dirname(os.path.realpath(__file__))
settings_file = os.path.join(here, 'ocra-data.conf.yaml')
with open(settings_file, 'r') as f:
    settings = yaml.load(f)

wbname = 'Reports '+date.today().strftime('%Y-%m-%d')
workbook_id = create_spreadsheet(wbname, settings['folder'])


dbs = settings['database']
#db = _mysql.connect(host=dbs['host'], port=dbs['port'], user=dbs['user'], passwd=dbs['passwd'], db=dbs['db'])
db = _mysql.connect(**dbs)

qry = 'SELECT * FROM reports ORDER BY run_order'

db.query(qry)
r = db.store_result()

report = r.fetch_row(how=1)
while report:
    report = report[0]
    db.query(report['query'])
    res = db.store_result()
    report_results = res.fetch_row(maxrows=0)
    headers = tuple([x[0] for x in res.describe()])
    data_to_worksheet(workbook_id, report['name'].decode(), report['description'].decode(), headers, report_results)
    report = r.fetch_row(how=1)
    
#delete_worksheet(workbook_id)