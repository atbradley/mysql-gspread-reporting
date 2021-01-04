import csv
import sys
import re
import os.path
from datetime import date

import _mysql
import yaml

from gsheetsExporter import gsheetsDataExporter

here = os.path.dirname(os.path.realpath(__file__))
settings_file = os.path.join(here, 'ocra-data.conf.yaml')
with open(settings_file, 'r') as f:
    settings = yaml.load(f)

wbname = 'Reports '+date.today().strftime('%Y-%m-%d')

dbs = settings['database']
db = _mysql.connect(**dbs)

qry = 'SELECT * FROM {} ORDER BY run_order'.format(settings['report_table_name'])

db.query(qry)
r = db.store_result()
reports = r.fetch_row(maxrows=0, how=1)

with gsheetsDataExporter(settings['apikeyfile'], settings['folder'], wbname) as gde:
    for report in reports:
        db.query(report['query'])
        res = db.store_result()
        report_results = res.fetch_row(maxrows=0)
        headers = tuple([x[0] for x in res.describe()])
        gde.data_to_worksheet(report['name'].decode(), report['description'].decode(), headers, report_results)

        csvname = re.sub('[^\\w]', '_', report['name'].decode())
        with open(csvname+'.csv','w', newline="") as fout:
            writer = csv.writer(fout, delimiter = ',')
            table = [headers] + list(report_results)
            writer.writerows(table)