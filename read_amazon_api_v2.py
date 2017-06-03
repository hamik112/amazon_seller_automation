import robobrowser,math,random,os
from bs4 import BeautifulSoup
import time

import sys, dateutil.parser, textwrap, csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A6
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.utils import ImageReader

ORDER_FILE = "input_orders.txt"
PREFS_FILE = "input_prefs.txt"
REFER_FILE = "input_reference.txt"
#os.chdir(os.path.dirname(os.path.realpath(__file__)))
if not os.path.isfile(PREFS_FILE) or not os.path.isfile(REFER_FILE):
    print("All 2 files are required.  " + PREFS_FILE + ", " + REFER_FILE)
    sys.exit()


br = robobrowser.RoboBrowser(history=True,user_agent='Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.110 Safari/537.36')

print ("logging in...")
sign_in = br.open("https://sellercentral.amazon.in/")

#br.select_form(name="signinWidget")
form = br.get_form(id="ap_signin_form")
form["email"].value = "hakeem.bracevor@gmail.com"
form["password"].value = "HakeemShiv408"
br.submit_form(form)
open("u.html","w").write(br.parsed.getText().decode().decode("cp1252"))
"""try:
    br.select_form(name="ap_dcq_form")
    br["dcq_question_subjective_1"] = "9739657055"
    submitted = br.submit()
    print ("verification needed...")
except:
    print ("No verif needed")
"""
print ("generating reports list...")



##################################
#open("u.txt","w").write(br.response().read())
try:
    orders_list = br.open("https://sellercentral.amazon.in/gp/transactions/actionableOrderPickup.html")
    form = br.get_form(id="requestForm")
except:
    print ("Err")
    orders_list = br.open("https://sellercentral.amazon.in/gp/transactions/actionableOrderPickup.html")
    br.select_form(name="requestForm")
sys.exit()
submitted = br.submit()
list = br.open( '/gp/upload-download-utils/reportStatusData.html?callerContext=actionableOrderReports&includeRequested=1&includeDateRange=1&includeReportType=1&diffScheduled=0&diffOrderReportVersion=0&marketplaceIDs=&nocache=' + str(math.floor(random.random() * 10000) ))
a = BeautifulSoup(list.read())
while(str(a.findAll("tr")[1].findAll("td")[-3].text) == "Not Complete"):
    print  ("waiting for the list to be ready")
    time.sleep(10)
    
    list = br.open( '/gp/upload-download-utils/reportStatusData.html?callerContext=actionableOrderReports&includeRequested=1&includeDateRange=1&includeReportType=1&diffScheduled=0&diffOrderReportVersion=0&marketplaceIDs=&nocache=' + str(math.floor(random.random() * 10000) ))
    a = BeautifulSoup(list.read())
print ("writing into file %s ..."%(ORDER_FILE))
for k in a.findAll("tr")[1].findAll("td")[-1].find("a").attrs:
    if str(k[0])=="href":
        open(ORDER_FILE,"w").write(br.open(k[1]).read())
        break

#output
print ("generating output file ...")
# Parse orders file, remove duplicates
ORDERS_FORMATTED = []
with open(ORDER_FILE,"rU") as f:
    csvreader = csv.reader(f, delimiter="\t")
    for index,row in enumerate(csvreader):
        if index == 0 or row == []:
            continue
        order_dict = {
            "order-id": row[0],
            "order-item-id": row[1],
            "purchase-date": row[2],
            "payment-date": row[3],
            "reporting-date": row[4],
            "promise-date": row[5],
            "days-past-promise": row[6],
            "buyer-email": row[7],
            "buyer-name": row[8],
            "buyer-phone-number": row[9],
            "sku": row[10],
            "product-name": row[11],
            "quantity-purchased": row[12],
            "quantity-shipped": row[13],
            "quantity-to-ship": row[14],
            "ship-service-level": row[15],
            "recipient-name": row[16],
            "ship-address-1": row[17],
            "ship-address-2": row[18],
            "ship-address-3": row[19],
            "ship-city": row[20],
            "ship-state": row[21],
            "ship-postal-code": row[22],
            "ship-country": row[23],
            "gift-wrap-type": row[24],
            "gift-message-text": row[25],
            "payment-method": "Pre-Paid" if not row[26] else row[26],
            "cod-collectible-amount": row[27],
            "already-paid": row[28],
            "payment-method-fee": row[29],
            "fulfilled-by": row[30]
            
        }
        if order_dict['fulfilled-by'].lower() != "easy ship":
            continue
        if not any(d['order-id'] == order_dict['order-id'] or d['buyer-email'] == order_dict['buyer-email'] for d in ORDERS_FORMATTED):
            ORDERS_FORMATTED.append(order_dict)
        else:
            for o in ORDERS_FORMATTED:
                if o['order-id'] == order_dict['order-id'] or o['buyer-email'] == order_dict['buyer-email']:
                    ORDERS_FORMATTED.remove(o);
"""
row = ["wal" for i in range(31)]
order_dict = {
    "order-id": row[0],
    "order-item-id": row[1],
    "purchase-date": row[2],
    "payment-date": row[3],
    "reporting-date": row[4],
    "promise-date": row[5],
    "days-past-promise": row[6],
    "buyer-email": row[7],
    "buyer-name": row[8],
    "buyer-phone-number": row[9],
    "sku": row[10],
    "product-name": row[11],
    "quantity-purchased": 20,
    "quantity-shipped": row[13],
    "quantity-to-ship": row[14],
    "ship-service-level": row[15],
    "recipient-name": row[16],
    "ship-address-1": row[17],
    "ship-address-2": row[18],
    "ship-address-3": row[19],
    "ship-city": row[20],
    "ship-state": row[21],
    "ship-postal-code": row[22],
    "ship-country": row[23],
    "gift-wrap-type": row[24],
    "gift-message-text": row[25],
    "payment-method": "Pre-Paid" if not row[26] else row[26],
    "cod-collectible-amount": row[27],
    "already-paid": row[28],
    "payment-method-fee": row[29],
    "fulfilled-by": row[30]
    
}
ORDERS_FORMATTED.append(order_dict)"""
# Parse reference file
REFERS_FORMATTED = []
with open(REFER_FILE,"rU") as fn:
    csvreader = csv.reader(fn, delimiter="\t")
    for index,row in enumerate(csvreader):
        if index == 0 or row == []:
            continue
        refer_dict = {
            "sku": row[0],
            "price": float(row[1]),
            "box": row[2],
            "weight": row[3],
            "l": row[4],
            "w": row[5],
            "h": row[6],
        }
        REFERS_FORMATTED.append(refer_dict)

# Parse preferences file
PREFS = {}
new_file_lines = []
with open(PREFS_FILE,"rU") as f:
    for line in f:
        new_file_line = line
        if not line.startswith("#"):
            line = line.strip()
            if "title" not in PREFS:
                PREFS["title"] = line
            elif "tin" not in PREFS:
                PREFS["tin"] = line
            elif "invoicestart" not in PREFS:
                PREFS["invoicestart"] = line
                new_file_line = str(int(line) + (len(ORDERS_FORMATTED) - 1) + 1) + "\n"
            elif "pickdate" not in PREFS:
                PREFS["pickdate"] = line
            elif "picktime" not in PREFS:
                PREFS["picktime"] = line
        new_file_lines.append(new_file_line)
with open(PREFS_FILE, "w") as f:
    f.writelines(new_file_lines)

# Write to PDF file
pagesize = (21*cm,25*cm)
c = canvas.Canvas("output_invoices.pdf", pagesize=pagesize)
width, height = pagesize
for index, order in enumerate(ORDERS_FORMATTED):
    defaultTop = height - (inch * 0.8)
    fromTop = 0.5
    ##
    #order["quantity-purchased"]=20
    #order_reference = {"price":20}
    try:
        order_reference = next((item for item in REFERS_FORMATTED if item["sku"] == order['sku']))
    except:
        with open("output_badsku.txt", "a") as f:
            f.write(order['sku'] + "\n")
        continue
    c.setFont("Helvetica-Bold", 25)
    c.drawCentredString(width / 2.0, defaultTop,"Tax Invoice")
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(width / 2.0, defaultTop - (inch * fromTop), PREFS['title'])
    fromTop += 0.4
    c.setFont("Helvetica", 12)
    c.drawRightString(width - inch, defaultTop - (inch * fromTop), "VAT/TIN: " + PREFS['tin'])
    fromTop += 0.2
    c.drawRightString(width - inch, defaultTop - (inch * fromTop), "Invoice no: " + str(int(PREFS['invoicestart']) + index))
    fromTop += 0.2
    c.drawRightString(width - inch, defaultTop - (inch * fromTop), "Order ID: " + order['order-id'])
    fromTop += 0.3
    c.line(inch, defaultTop - (inch * fromTop), width - inch, defaultTop - (inch * fromTop))
    c.setFont("Helvetica-Bold", 12)
    fromTop += 0.4
    c.drawString(inch, defaultTop - (inch * fromTop), "Delivery Address:")
    c.setFont("Helvetica", 11)
    fromTop += 0.25
    c.drawString(inch, defaultTop - (inch * fromTop), order['recipient-name'])
    ship_address = order['ship-address-1'].replace("\"", "")
    ship_address_splitted = textwrap.wrap(ship_address, 40)
    for add in ship_address_splitted:
        fromTop += 0.2
        c.drawString(inch, defaultTop - (inch * fromTop), unicode(add.decode("cp1252")))
    fromTop += 0.2
    c.drawString(inch, defaultTop - (inch * fromTop), order['ship-city'])
    fromTop += 0.2
    c.drawString(inch, defaultTop - (inch * fromTop), order['ship-state'] + ", " + order['ship-country'] + " " + order['ship-postal-code'])
    fromTop = 1.85
    c.drawRightString(width - inch, defaultTop - (inch * fromTop), "Order Date: " + dateutil.parser.parse(order['purchase-date']).strftime("%d-%b-%y"))
    fromTop += 0.2
    c.drawRightString(width - inch, defaultTop - (inch * fromTop), "Shipping Service: " + order['ship-service-level'])
    fromTop += 0.2
    c.drawRightString(width - inch, defaultTop - (inch * fromTop), "Buyer Name: " + order['buyer-name'])
    fromTop += 0.2
    c.drawRightString(width - inch, defaultTop - (inch * fromTop), "Seller Name: " + PREFS['title'])
    c.setFont("Helvetica", 12)
    fromTop += 0.75
    c.drawString(inch + (inch * 0.2), defaultTop - (inch * 3.2), "Payment mode: " + order['payment-method'])
    headerstyle = getSampleStyleSheet()['Normal']
    headerstyle.leading = 24
    headerstyle.fontName = "Helvetica-Bold"
    headerstyle.alignment = TA_CENTER
    data = [['Quantity', 'Product Details', 'Price']]
    data.append([order['quantity-purchased'], Paragraph(order['product-name'], getSampleStyleSheet()['Normal']), 'Rs ' + str(order_reference['price'])])
    data.append(['', 'Tax Included', ''])
    data.append(['Total', 'Total', 'Rs ' + str(int(order['quantity-purchased']) * order_reference['price'])])
    tablewidth = width - inch - inch
    table = Table(data, colWidths=(80, tablewidth - 180, 100))
    table.setStyle(TableStyle([
        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
        ('BOX', (0,0), (-1,-1), 0.25, colors.black),
        ('ALIGN', (0,0),(-1,-1), 'CENTER'),
        ('ALIGN', (1,2),(1,2), 'RIGHT'),
        ('VALIGN', (0,1), (2,1), 'MIDDLE'),
        ('FONTNAME', (0,0),(2,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,3),(2,3), 'Helvetica-Bold'),
        ('SPAN', (0,3), (1,3)),
        ('TOPPADDING', (0,0), (2,1), 10),
        ('BOTTOMPADDING', (0,0), (2,1), 10),
        ('TOPPADDING', (0,3), (2,3), 5),
        ('BOTTOMPADDING', (0,3), (2,3), 5)
    ]))
    table.wrapOn(c, tablewidth, 400)
    table.drawOn(c, inch, defaultTop - (inch * 3.4) - 120)
    c.drawRightString(width - inch, defaultTop - (inch * 3.4) - 240, "authorized signatory")
    logo = ImageReader('Authorized_signature.jpg')
    
    c.drawImage(logo, width - 2.5*inch, defaultTop - (inch * 3.4) - 220,width=108,height=60, mask='auto')
    c.line(inch, defaultTop - (inch * 3.4) - 280, width - inch, defaultTop - (inch * 3.4) - 280)
    c.drawString(inch, defaultTop - (inch * 3.4) - 300, "Registered address:")
    c.drawString(inch, defaultTop - (inch * 3.4) - 315, "Bracevor India Pvt Ltd, 1st cross, Omkar nagar, Bangalore 560076 Karnataka, IN")
    c.showPage()
c.save()

# Write to output_pickup.txt
with open("output_pickup.txt", "w") as f:
    csvwriter = csv.writer(f, delimiter="\t")
    csvwriter.writerow(['order_id', 'invoice_id', 'package_weight', 'package_length', 'package_width', 'package_height', 'schedule_pickup_date', 'schedule_pickup_time', 'merchant_additional_identifier'])
    for index, order in enumerate(ORDERS_FORMATTED):
        try:
            order_reference = next((item for item in REFERS_FORMATTED if item["sku"] == order['sku']))
        except:
            continue
        dateformat = dateutil.parser.parse(PREFS['pickdate']).strftime("%Y-%m-%d")
        csvwriter.writerow([order['order-id'], str(int(PREFS['invoicestart']) + index), order_reference['weight'], order_reference['l'], order_reference['w'], order_reference['h'], dateformat, PREFS['picktime']])
        
#upload output file
br.open("https://sellercentral.amazon.in/gp/transactions/uploadSchedulePickup.html")
print ("uploading file...")
br.select_form(name="uploadForm")
br.form.add_file(open("output_pickup.txt"), 'text/plain', "output_pickup.txt")

br.submit()


#download pdf_invoices
list = br.open( 'https://sellercentral.amazon.in//gp/upload-download-utils/uploadStatusData.html?internalFileType=SchedulePickup&troubleshootErrorsHelpID=&maxResults=10&showReuploadLinks=0&marketplaceIDs=&nocache=' + str(math.floor(random.random() * 10000) ))
a = BeautifulSoup(list.read())
while(len(a.findAll("tr")[1].findAll("td")[-1].findAll("a"))!=2):
    print ("waiting for the pdf to be ready")
    time.sleep(10)
    list = br.open( 'https://sellercentral.amazon.in//gp/upload-download-utils/uploadStatusData.html?internalFileType=SchedulePickup&troubleshootErrorsHelpID=&maxResults=10&showReuploadLinks=0&marketplaceIDs=&nocache=' + str(math.floor(random.random() * 10000) ))
    a = BeautifulSoup(list.read())

link=str(a.findAll("tr")[1].findAll("td")[-1].findAll("a")[1].attrs[0][1])
open("labels.pdf","wb").write(br.open(link,"rb").read())
