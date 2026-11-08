# INSERT ANSWER FOR QUESTION 1 HERE
directory = f'{os.getcwd()}\\'

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}
descriptionsdf = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')
descriptionsdf.head()
for index, row in descriptionsdf.iterrows():
    description = row['Description']
    unitnum = row['Unit number']
    product_descriptions[unitnum] = description
    #print(f'{unitnum} is {description}')
pricesdf = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Prices')
for index, row in pricesdf.iterrows():
    price = row['Price per unit']
    unitnum = row['Unit number']
    product_prices[unitnum] = price
    #print(f'{unitnum} price is {price}')

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):
    invoice_data = openpyxl.load_workbook(directory + filename)
    sheet = invoice_data['OrderForm']
    orderdate = (sheet['C9'].value)
    month = orderdate.month
    year = orderdate.year
    day = orderdate.day
    orderdate = f'{month}/{day}/{year}'
    customeremail = sheet['C13'].value
    customername = sheet['C14'].value
    companyname = sheet['C15'].value
    customeraddress = sheet['C16'].value
    customercity = sheet['C17'].value
    customerstate = sheet['C18'].value
    customerzip = sheet['C19'].value
    ordertotal = 0
    order_data_dict = {}
    if sheet['B24'].value is not None:
        item1 = sheet['B24'].value
        item1quant = sheet['C24'].value
        price1 = product_prices[item1]
        total1 = price1*item1quant
        ordertotal = ordertotal + total1
        total1 = f'${total1:,.2f}'
        price1 = f'${price1:,.2f}'
        description1 = product_descriptions[item1]
        order_data_dict['Price0'] = price1
        order_data_dict['Quantity0'] = str(item1quant)
        order_data_dict['Amount0'] = total1
        order_data_dict['Description0'] = description1

    if sheet['B25'].value is not None:
        item2 = sheet['B25'].value
        item2quant = int(sheet['C25'].value)
        price2 = product_prices[item2]
        total2 = price2*item2quant
        ordertotal = ordertotal + total2
        total2 = f'${total2:,.2f}'
        price2 = f'${price2:,.2f}'
        description2 = product_descriptions[item2]
        order_data_dict['Price1'] = price2
        order_data_dict['Quantity1'] = str(item2quant)
        order_data_dict['Amount1'] = total2
        order_data_dict['Description1'] = description2

    if sheet['B26'].value is not None:
        item3 = sheet['B26'].value
        item3quant = sheet['C26'].value
        price3 = product_prices[item3]
        total3 = price3*item3quant
        ordertotal = ordertotal + total3
        total3 = f'${total3:,.2f}'
        price3 = f'${price3:,.2f}'
        description3 = product_descriptions[item3]
        order_data_dict['Price2'] = price3
        order_data_dict['Quantity2'] = str(item3quant)
        order_data_dict['Amount2'] = total3
        order_data_dict['Description2'] = description3

    if sheet['B27'].value is not None:
        item4 = sheet['B27'].value
        item4quant = sheet['C27'].value
        price4 = product_prices[item4]
        total4 = price4*item4quant
        ordertotal = ordertotal + total4
        total4 = f'${total4:,.2f}'
        price4 = f'${price4:,.2f}'
        description4 = product_descriptions[item4]
        order_data_dict['Price3'] = price4
        order_data_dict['Quantity3'] = str(item4quant)
        order_data_dict['Amount3'] = total4
        order_data_dict['Description3'] = description4

    order_data_dict['InvoiceNum'] = str(invoice_num)
    order_data_dict['Date'] = orderdate
    order_data_dict['Total'] = f'${ordertotal:,.2f}'
    order_data_dict['CustomerAddress1'] = customeraddress
    order_data_dict['CustomerName'] = customername
    order_data_dict['CustomerCompany'] = companyname
    order_data_dict['CustomerAddress2'] = f'{customercity}, {customerstate} {customerzip}'
        
    return order_data_dict, customername, customeremail, orderdate

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)
create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
gmail_user = "jgammon8@gmail.com"
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num, name, email, date, directory):
    recipients = [email]
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)
    message = add_attachment(directory+f'Invoice {invoice_num}.pdf', message)
    
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'jgammon8@gmail.com','11/13/2023',directory)

# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames = sorted(filenames)
invoice_num = 8515
for filename in filenames: 
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num,filename,directory)
        create_invoice(invoice_num,order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
    
