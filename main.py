from flask import Flask, render_template, request, send_file
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import exc, cast, Date
from sqlalchemy.ext.automap import automap_base 
from sqlalchemy.orm import Session
from datetime import date, datetime, timedelta, timezone
from smtplib import SMTPException
from dotenv import load_dotenv
from logging.config import dictConfig

import xlsxwriter
import os
import csv


dictConfig(
            {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
            'default': {
                        'format': '[%(asctime)s] %(levelname)s in %(module)s: %(message)s',
                       },
            'simpleformatter' : {
                        'format' : '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            }
    },
    'handlers':
    {
        'custom_handler': {
            'class' : 'logging.FileHandler',
            'formatter': 'default',
            'filename' : 'warnings.log',
            'level': 'WARN',
        }
    },
    'root': {
        'level': 'WARN',
        'handlers': ['custom_handler']
    },
})

os.environ["WERKZEUG_RUN_MAIN"] = "true"
load_dotenv()

app = Flask(__name__)

# Database 
driver = 'ODBC Driver 17 for SQL Server'
user_name = os.getenv("USER_NAME")
server = os.getenv("SERVER_NAME")
db_name = os.getenv("DB_NAME")
password = os.getenv("DB_PASS")
# app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc://{user_name}:{password}@{server}/{db_name}?driver={driver}"
app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc://{server}/{db_name}?driver={driver}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False



db = SQLAlchemy(app)

Base = automap_base()

def _name_for_collection_relationship(base, local_cls, referred_cls, constraint):
    if constraint.name:
        return constraint.name.lower()
    # if this didn't work, revert to the default behavior
    return name_for_collection_relationship(base, local_cls, referred_cls, constraint)

Base.prepare(db.engine, reflect=True, name_for_collection_relationship=_name_for_collection_relationship)
session = Session(db.engine,autocommit=False)

# DB Model Classes
Orders = Base.classes.Orders
OrderScans = Base.classes.OrderScans
OrderPackageItems = Base.classes.OrderPackageItems


def check_last_audit():
    # Check if any previous audit has been conducted
    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = os.path.join(dir_path, "lastaudit.txt")
    last_audit = None
    if os.path.exists(file_path) and os.stat(file_path).st_size > 0:
        with open(file_path, 'r') as f:
            last_line = f.readlines()[-1]
            last_audit = datetime.strptime(last_line.strip('\n'), "%Y-%m-%d %H:%M:%S.%f")
            return last_audit
    return last_audit

# Return list of files with scancodes to be scanned
def generate_scan_file_list(last_audit):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    dir_path = os.path.join(dir_path, "packagesreceived")
    file_list = []
    for file in os.listdir(dir_path):
        if file.startswith("PackagesReceived") and file.endswith(".csv"):
            scan_file = os.path.join(dir_path, file)
            lmt = os.path.getmtime(scan_file)
            modified = datetime.fromtimestamp(lmt)
            if last_audit is not None:
                if modified >= last_audit:
                    file_list.append(scan_file)
            else:
                file_list.append(scan_file)
    return file_list

def generate_master_list_scan_codes(file_list):

    scan_codes = {}
    for file in file_list:
        with open(file) as csvfile: 
            csvreader =  csv.reader(csvfile, delimiter=',')
            header = next(csvreader)
            for line in csvreader:
                scan_codes[line[0]] = line[1]
    
    return scan_codes


# Cross-reference scan codes 
def get_unscanned_codes(master_scan_codes): 
    last_hour = datetime.today() - timedelta(hours=1)
    last_hour = last_hour.replace(minute=0, second=0, microsecond=0)
    scan_codes = master_scan_codes.keys()
    db_scan_codes = session.query(OrderScans.SCANcode)
    db_scan_codes = db_scan_codes.filter(
        OrderScans.SCANlocation == 'R', 
        OrderScans.aTimeStamp.cast(Date) >= last_hour
    )
    db_scan_codes = [r._asdict() for r in db_scan_codes.all()]
    order_scans = [d['SCANcode'] for d in db_scan_codes]
    
    unscanned_codes = list(set(scan_codes) - set(order_scans))
   
    return unscanned_codes

# Get OrderTrackingID for packages without scan codes
def get_order_tracking_ids(unscanned_codes):
    threshold =  datetime.today() - timedelta(days=14)
    threshold = threshold.date()
    db_orders = session.query(OrderPackageItems.OrderTrackingID, OrderPackageItems.RefNo)
    db_orders = db_orders.join(Orders, OrderPackageItems.OrderTrackingID == Orders.OrderTrackingID)
    db_orders = db_orders.filter(
        OrderPackageItems.RefNo.in_(unscanned_codes),
        Orders.oDate.cast(Date) >= threshold
        ).all()
    db_orders = [r._asdict() for r in db_orders]
    orders = {}
    for order in db_orders:
        k = order['RefNo']
        v = order['OrderTrackingID']
        orders[k] = v
    return orders


# Generate audit report file
def generate_audit_report(scan_codes, unscanned_orders):

    today = date.today()
    today = today.strftime("%m_%d_%y")
    file_name = 'Audit_Report-' + today + '.xlsx'
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()

    headers = ['OrderTrackingID', 'ScanCode', 'TimeStamp']
    for x in range(len(headers)):
        worksheet.write(0, x, headers[x])

    for idx, scan in enumerate(scan_codes.keys()):
        if scan in unscanned_orders: 
            worksheet.write(idx+1, 0, unscanned_orders[scan])
        else:
            worksheet.write(idx+1, 0, 'None')
        worksheet.write(idx+1, 1, scan)
        worksheet.write(idx+1, 2, scan_codes[scan])
    
    workbook.close()

    # Write current timestamp to lastaudit
    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = os.path.join(dir_path, "lastaudit.txt")
    with open(file_path, 'w') as f:
        f.write(str(datetime.now()))
        f.write('\n')

    return send_file(
        file_name,
        mimetype='application/vnd.ms-excel', 
        as_attachment=True
    )


@app.route('/')
def home_rte():
    return render_template('home.html')
    

@app.route('/report',  methods=["GET", "POST"])
def report_rte():
    passcode=request.form.get("passcode")
    if passcode != os.getenv("PASSCODE"):
        return render_template('403.html')

    last_audit = check_last_audit()    
    file_list = generate_scan_file_list(last_audit)
    scan_codes = generate_master_list_scan_codes(file_list)
    unscanned_codes = get_unscanned_codes(scan_codes)
    unscanned_orders = get_order_tracking_ids(unscanned_codes)
    return generate_audit_report(scan_codes, unscanned_orders)
    
    

if __name__ == "__main__":
    app.run()