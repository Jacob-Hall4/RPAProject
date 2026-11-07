# INSERT ANSWER FOR QUESTION 1 HERE
while True:
    directory = single_input('Enter directory')
    if directory[-1] != '/':
        directory = directory + '/'

    list= os.listdir(directory)
    if 'ProductData.xlsx' in list and 'MasterInvoice.pdf' in list and 'Order1.xlsx' in list and 'Order2.xlsx' in list and 'Order3.xlsx' in list and 'Order4.xlsx' in list:
        break
    else:
        show_message('Error','Re-enter the directory')

product_descriptions = {}
descriptions = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')
for index, row in descriptions.iterrows():
    Unit = row['Unit number']
    desc = row['Description']
    product_descriptions[Unit] = desc
print(product_descriptions)

product_prices = {}
descriptions = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')
for index, row in descriptions.iterrows():
    Unit = row['Unit number']
    ppu = row['Price per unit']
    product_prices[Unit] = ppu
print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE
def get_order_data(invoice_num, filename, directory):  # processes the excel file to get information
    workbook = openpyxl.load_workbook(directory + filename)
    order_data = workbook['OrderForm']
    name = order_data['C14'].value
    email = order_data['C13'].value
    wrongdate = order_data['C9'].value
    month = wrongdate.month
    day = wrongdate.day
    year = wrongdate.year
    date = f'{month}/{day}/{year}'


    update_dict = {}
    update_dict['InvoiceNum'] = invoice_num
    update_dict['Date'] = date
    update_dict['CustomerAddress1'] = order_data['C16'].value
    update_dict['CustomerName'] = name
    update_dict['CustomerCompany'] = order_data['C15'].value
    update_dict['CustomerAddress2'] = str(order_data['C17'].value) + ', ' + str(order_data['C18'].value) + ', ' + str(order_data['C19'].value)
    update_dict['Description0'] = product_descriptions[order_data['B24'].value]
    update_dict['Price0'] = float(product_prices[order_data['B24'].value])
    update_dict['Quantity0'] = order_data['C24'].value
    update_dict['Amount0'] = round(update_dict['Price0'] * update_dict['Quantity0'], 2)
    update_dict['Total'] = update_dict['Amount0']
    
    if order_data['B25'].value is not None:
        update_dict['Price1'] = float(product_prices[order_data['B25'].value])
        update_dict['Quantity1'] = order_data['C25'].value
        update_dict['Amount1'] = round(update_dict['Price1'] * update_dict['Quantity1'], 2)
        update_dict['Description1'] = product_descriptions[order_data['B25'].value]
        update_dict['Total'] = update_dict['Total'] + update_dict['Amount1']
        
    if order_data['B26'].value is not None:
        update_dict['Price2'] = float(product_prices[order_data['B26'].value])
        update_dict['Quantity2'] = order_data['C26'].value
        update_dict['Amount2'] = round(update_dict['Price2'] * update_dict['Quantity2'], 2)
        update_dict['Description2'] = product_descriptions[order_data['B26'].value]
        update_dict['Total'] = update_dict['Total'] + update_dict['Amount2']
        
    if order_data['B27'].value is not None:
        update_dict['Price3'] = float(product_prices[order_data['B27'].value])
        update_dict['Quantity3'] = order_data['C27'].value
        update_dict['Amount3'] = round(update_dict['Price3'] * update_dict['Quantity3'], 2)
        update_dict['Description3'] = product_descriptions[order_data['B27'].value]
        update_dict['Total'] = update_dict['Total'] + update_dict['Amount3']
        
    if order_data['B28'].value is not None:
        update_dict['Price4'] = float(product_prices[order_data['B28'].value])
        update_dict['Quantity4'] = order_data['C28'].value
        update_dict['Amount4'] = round(update_dict['Price4'] * update_dict['Quantity4'], 2)
        update_dict['Description4'] = product_descriptions[order_data['B28'].value]
        update_dict['Total'] = update_dict['Total'] + update_dict['Amount4']

    # convert to string for everything (incl amounts) at the END after total is done
    update_dict['Total'] = f'${update_dict["Total"]:,.2f}'
    update_dict['Price0'] = f'${update_dict["Price0"]:,.2f}'
    update_dict['Amount0'] = f'${update_dict["Amount0"]:,.2f}'
    
    if order_data['B25'].value is not None:
        update_dict['Price1'] = f'${update_dict["Price1"]:,.2f}'
        update_dict['Amount1'] = f'${update_dict["Amount1"]:,.2f}'
        
    if order_data['B26'].value is not None:
        update_dict['Price2'] = f'${update_dict["Price2"]:,.2f}'
        update_dict['Amount2'] = f'${update_dict["Amount2"]:,.2f}'
        
    if order_data['B27'].value is not None:
        update_dict['Price3'] = f'${update_dict["Price3"]:,.2f}'
        update_dict['Amount3'] = f'${update_dict["Amount3"]:,.2f}'
        
    if order_data['B28'].value is not None:
        update_dict['Price4'] = f'${update_dict["Price4"]:,.2f}'
        update_dict['Amount4'] = f'${update_dict["Amount4"]:,.2f}'

    return update_dict, name, email, date

order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)

print(order_data)
print(name)
print(email)
print(date)

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f"Invoice {invoice_num}.pdf",order_data)
create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    gmail_user = 'jacobhhall4@gmail.com'
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    recipients = [email]

    subject = f"Your Order from PythonTech Co. - Invoice {invoice_num}"

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}
    
    Thank you for your business.
    
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = os.getcwd()+ f'/Invoice {invoice_num}.pdf' # MAC
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

directory = os.getcwd()
send_email(8515, 'Thomas Holt', 'jacobhhall4@gmail.com', '11/13/2023', directory)

filenames = os.listdir(directory)
directory = directory + '/'
filenames.sort()
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date =  get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
