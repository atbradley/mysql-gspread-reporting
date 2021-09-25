#TODO use pathlib
import os.path
from datetime import date

from MySQLdb import _mysql
import yaml

from google_sheets_utils import create_spreadsheet, data_to_worksheet

#TODO: Define conf file on command line.
here = os.path.dirname(os.path.realpath(__file__))
settings_file = os.path.join(here, '..', "ocra-data.conf.yaml")
with open(settings_file, "r") as f:
    settings = yaml.load(f, Loader=yaml.SafeLoader)

wbname = "Reports " + date.today().strftime("%Y-%m-%d")
workbook_id = create_spreadsheet(wbname, settings["folder"])


dbs = settings["database"]
db = _mysql.connect(**dbs)

qry = "SELECT * FROM {} ORDER BY run_order".format(settings["report_table_name"])

db.query(qry)
r = db.store_result()

report = r.fetch_row(how=1)
while report:
    report = report[0]
    db.query(report["query"])
    res = db.store_result()
    report_results = res.fetch_row(maxrows=0)
    headers = ([x[0] for x in res.describe()])
    data_to_worksheet(
        workbook_id,
        report["name"].decode(),
        report["description"].decode(),
        headers, 
        report_results,
    )
    report = r.fetch_row(how=1)
