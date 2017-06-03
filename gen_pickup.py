import os, sys, dateutil.parser, textwrap, csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A6
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER

ORDER_FILE = "input_orders.txt"
PREFS_FILE = "input_prefs.txt"
REFER_FILE = "input_reference.txt"

if not os.path.isfile(ORDER_FILE) or not os.path.isfile(PREFS_FILE) or not os.path.isfile(REFER_FILE):
    print("All 3 files are required. " + ORDER_FILE + ", " + PREFS_FILE + ", " + REFER_FILE)
    sys.exit()

# Parse orders file, remove duplicates
ORDERS_FORMATTED = []
with open(ORDER_FILE) as f:
    csvreader = csv.reader(f, delimiter="\t")
    for index, row in enumerate(csvreader):
        if index == 0:
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

# Parse reference file
REFERS_FORMATTED = []
with open(REFER_FILE) as f:
    csvreader = csv.reader(f, delimiter="\t")
    for index, row in enumerate(csvreader):
        if index == 0:
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
with open(PREFS_FILE) as f:
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
pagesize = (21*cm,18*cm)
c = canvas.Canvas("output_invoices.pdf", pagesize=pagesize)
width, height = pagesize
for index, order in enumerate(ORDERS_FORMATTED):
    defaultTop = height - (inch * 0.8)
    fromTop = 0.5
    try:
        order_reference = next((item for item in REFERS_FORMATTED if item["sku"] == order['sku']))
    except:
        with open("output_badsku.txt", "w") as f:
            f.write(order['sku'] + "\n")
        continue
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(width / 2.0, defaultTop, PREFS['title'])
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
        c.drawString(inch, defaultTop - (inch * fromTop), add)
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
        dateformat = dateutil.parser.parse(PREFS['pickdate']).strftime("%Y/%m/%d")
        csvwriter.writerow([order['order-id'], str(int(PREFS['invoicestart']) + index), order_reference['weight'], order_reference['l'], order_reference['w'], order_reference['h'], dateformat, PREFS['picktime']])
