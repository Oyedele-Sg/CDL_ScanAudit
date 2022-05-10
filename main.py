from asyncio.log import logger
from time import time
from flask import Flask, render_template, request, send_file
from flask_mail import Mail, Message
from flask_sqlalchemy import SQLAlchemy
from pymysql import Time
from sqlalchemy import exc, cast, Date, Time
from sqlalchemy.ext.automap import automap_base 
from sqlalchemy.orm import Session
from datetime import date, datetime, timedelta
from dotenv import load_dotenv
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

import logging, logging.handlers
import xlsxwriter
import os
import csv



os.environ["WERKZEUG_RUN_MAIN"] = "true"
load_dotenv()

def setup_log(name):
    logger = logging.getLogger(name)   

    logger.setLevel(logging.DEBUG)

    log_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    filename = f"{name}.log"
    log_handler = logging.FileHandler(filename)
    smtp_handler = logging.handlers.SMTPHandler(mailhost=('smtp.gmail.com', 587),
                                                fromaddr=str(os.getenv('EMAIL')),
                                                toaddrs=[str(os.getenv('SUPPORT'))],
                                                subject='Error In CDL ScanAudit',
                                                credentials=(str(os.getenv('EMAIL')), str(os.getenv('MAIL_PASS'))),
                                                secure=())
    log_handler.setLevel(logging.DEBUG)
    smtp_handler.setLevel(logging.WARNING)
    log_handler.setFormatter(log_format)
    smtp_handler.setFormatter(log_format)

    logger.addHandler(log_handler)
    logger.addHandler(smtp_handler)

    return logger



def start_log(name):
    logger = setup_log(name)
    logger.info("Just logged from %s", name)   

start_log("scanaudit")



app = Flask(__name__)
mail = Mail(app)

# Database 
driver = 'SQL Server'
user_name = os.getenv("USER_NAME")
server = os.getenv("SERVER_NAME")
db_name = os.getenv("DB_NAME")
password = os.getenv("DB_PASS")
app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc://{user_name}:{password}@{server}/{db_name}?driver={driver}"
# app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc://{server}/{db_name}?driver={driver}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN'] = True
app.config['SQLALCHEMY_NATIVE_UNICODE'] = True
# configuration of mail
app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = str(os.getenv('EMAIL'))
app.config['MAIL_DEFAULT_SENDER'] = str(os.getenv('EMAIL'))
app.config['MAIL_PASSWORD'] = str(os.getenv('MAIL_PASS'))
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
recipients = []
for r in  os.getenv('ADMINS').split(','):
    recipients.append(str(r))

support = []
for r in  os.getenv('SUPPORT').split(','):
    support.append(str(r))

mail = Mail(app)

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
            last_line = last_line.strip('\n')
            if len(last_line) > 0:
                last_audit = datetime.strptime(last_line.strip('\n'), "%Y-%m-%d %H:%M:%S.%f")
            return last_audit
    return last_audit
        

# Return list of files with scancodes to be scanned
def generate_scan_file_list(last_audit):
    dir_path = os.getenv('FILE_DIR')
    file_list = []
    for file in os.listdir(dir_path):
        if 'PackagesReceived' in file and file.endswith(".csv"):
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
    last_hour = datetime.today() - timedelta(hours=int(os.getenv('HOUR_THRESHOLD')))
    scan_codes = master_scan_codes.keys()
    db_scan_codes = session.query(OrderScans.SCANcode)
    db_scan_codes = db_scan_codes.filter(
        OrderScans.SCANlocation == 'R', 
        OrderScans.aTimeStamp >= last_hour
    ).all()
    db_scan_codes = [r._asdict() for r in db_scan_codes]
    order_scans = [d['SCANcode'] for d in db_scan_codes]
    unscanned_codes = list(set(scan_codes) - set(order_scans))
    return unscanned_codes
   

# Get OrderTrackingID for packages without scan codes
def get_order_tracking_ids():
    threshold =  datetime.today() - timedelta(days=14)
    threshold = threshold.date()
    db_orders = session.query(OrderPackageItems.OrderTrackingID, OrderPackageItems.RefNo)
    db_orders = db_orders.join(Orders, OrderPackageItems.OrderTrackingID == Orders.OrderTrackingID)
    db_orders = db_orders.filter(Orders.oDate.cast(Date) >= threshold).all()
    db_orders = [r._asdict() for r in db_orders]
    unscanned_orders = {}
    for order in db_orders:
        k = order['RefNo']
        v = order['OrderTrackingID']
        unscanned_orders[k] = v
    return unscanned_orders

def write_timestamp():
    # Write current timestamp to lastaudit
    dir_path = os.path.dirname(os.path.realpath(__file__))
    file_path = os.path.join(dir_path, "lastaudit.txt")
    with open(file_path, 'w') as f:
        f.write(str(datetime.now()))
        f.write('\n')

# Generate audit report file
def generate_audit_report(master_scan_list, order_package_items, unscanned_codes):
    today = datetime.now()
    today = today.strftime("%m_%d_%y_%H_%M_%S")
    file_name = 'Audit_Report-' + today + '.xlsx'
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()

    headers = ['OrderTrackingID', 'ScanCode', 'TimeStamp']
    for x in range(len(headers)):
        worksheet.write(0, x, headers[x])
    
    for idx, scan in enumerate(unscanned_codes):
        if scan in order_package_items: 
            worksheet.write(idx+1, 0, order_package_items[scan])
        else:
            worksheet.write(idx+1, 0, 'None')
        worksheet.write(idx+1, 1, scan)
        worksheet.write(idx+1, 2, master_scan_list[scan])

    workbook.close()

    subject = 'Scan Audit - ' + today
    msg = Message(
                    sender=('btech@cdldelivers.com', str(os.getenv('EMAIL'))),
                    subject=subject,
                    recipients = recipients
                )
    msg.body = 'Find attached the scan audit report in the email'
    file = open(file_name, 'rb')

    
    msg.attach(file_name, '	application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', file.read())
    mail.send(msg)
    return send_file(
        file_name,
        mimetype='application/vnd.ms-excel', 
        as_attachment=True
    )

def get_scan_report():
    with app.test_request_context(): 
        last_audit = check_last_audit()    
        print("LAST AUDIT: ", last_audit)
        file_list = generate_scan_file_list(last_audit)
        print(len(file_list))
        master_scan_list = generate_master_list_scan_codes(file_list)
        unscanned_codes = get_unscanned_codes(master_scan_list)
        order_package_items = get_order_tracking_ids()
        write_timestamp()
        generate_audit_report(master_scan_list, order_package_items, unscanned_codes)

    return 'Reported generated successfully'
    
    # return generate_audit_report(master_scan_list, order_package_items, unscanned_codes)
        


@app.route('/')
def home_rte():
    return render_template('home.html')


@app.route('/auditscan',  methods=["GET", "POST"])
def audit_scan_rte():
    passcode=request.form.get("passcode")
    if passcode != os.getenv("PASSCODE"):
        return render_template('403.html')

    return get_scan_report()


@app.route('/report',  methods=["GET", "POST"])
def report_rte():
    return get_scan_report()

 
@app.errorhandler(500)
def internal_error(exception):
    return render_template('500.html'), 500


sched = BackgroundScheduler(daemon=True)
sched.add_job(get_scan_report,'interval', hours=4)
sched.add_job(write_timestamp,'interval', hours=4)
sched.start()

if __name__ == "__main__":
    app.run(host="localhost", port=8083)



