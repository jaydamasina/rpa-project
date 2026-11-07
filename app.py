# question 1
needed_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:    
    directory = single_input("Please enter path.")
    if directory.endswith('/'):
        try: 
            filenames = os.listdir(directory)
            file_count=0
            for f, file in enumerate(needed_files):
                if file in filenames:
                    file_count=file_count+1
                else:
                    show_message('Incorrect',f'{file} is not in filenames.')
            if file_count==6:
                show_message("Correct", "path had '/' and this path works!")
                break
            else:
                show_message("Incorrect", "Path had '/' and this path did not work.")
        except: 
            show_message("Incorrect", "Path had '/' and this path did not work.")
    elif directory[-1]!='/':
       directory = directory.strip("'") + '/'
       try:
            filenames = os.listdir(directory)
            file_count=0
            for f, file in enumerate(needed_files):
                if file in filenames:
                    file_count=file_count+1
                else:
                    show_message('Incorrect',f'{file} is not in filenames.')
            if file_count==6:
                show_message("Correct", "Path did not have '/' and this path works!")
                break
            else:
                show_message("Incorrect", "Path did not have '/' and this path did not work.")    
       except: 
            show_message("Incorrect", "Path did not have '/' and this path did not work.")

# Question 2
product_descriptions={}
product_prices={}


product_data = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Descriptions")
for i,row in product_data.iterrows():
    unit_no = row['Unit number']
    description = row['Description']
    product_descriptions[unit_no] = description


prices_data = pd.read_excel(directory + "ProductData.xlsx", sheet_name="Prices")
for i,row in prices_data.iterrows():
    unit_price_no = row['Unit number']
    prices = row['Price per unit']
    product_prices[unit_price_no] = prices

# Question 3
def get_order_data(invoice_num,filename , directory):
    order_form=openpyxl.load_workbook(directory + filename)
    selected_sheet=order_form['OrderForm']
    order_data={}
 #date
    order_data['Date'] = str(selected_sheet['C9'].value.strftime('%m/%d/%Y'))
 #total
    item1=selected_sheet['B24'].value
    unit_price1=product_prices.get(item1, '0')
    total = float(selected_sheet['C24'].value * unit_price1)
    if pd.isnull(selected_sheet['B25'].value)==False:
        item2=selected_sheet['B25'].value
        unit_price2=product_prices.get(item2, '0')
        total = total + float(selected_sheet['C25'].value * unit_price2)
    if pd.isnull(selected_sheet['B26'].value)==False:
        item3=selected_sheet['B26'].value
        unit_price3=product_prices.get(item3, '0')
        total = total + float(selected_sheet['C26'].value * unit_price3)
    order_data['Total'] = f'${total:,.2f}'
#Address1
    address= selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state= selected_sheet['C18'].value
    zip= selected_sheet['C19'].value
    order_data['CustomerAddress1']= f'{address}, {city}, {state} {zip}'
#customer name
    order_data['CustomerName'] = selected_sheet['C14'].value
#company name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = selected_sheet['C13'].value #email address?
#order price, quantity, amount
    item0=selected_sheet['B24'].value
    order_data['Price0'] = product_prices.get(item0, '0')
    order_data['Quantity0'] = selected_sheet['C24'].value
    price0 = order_data['Price0']* order_data['Quantity0']
    order_data['Amount0']=f'${price0:,.2f}'
    order_data['Description0']=product_descriptions.get(item0,'0')
    if pd.isnull(selected_sheet['B25'].value)==False:
        item1=selected_sheet['B25'].value
        order_data['Price1'] = product_prices.get(item1, '0')
        order_data['Quantity1'] = selected_sheet['C25'].value
        price1 = order_data['Price1']* order_data['Quantity1']
        order_data['Amount1']=f'${price1:,.2f}'
        order_data['Description1']=product_descriptions.get(item1,'0')
    if pd.isnull(selected_sheet['B26'].value)==False:
        item2=selected_sheet['B26'].value
        order_data['Price2'] =product_prices.get(item2, '0')
        order_data['Quantity2'] =  selected_sheet['C26'].value
        price2 = order_data['Price2']* order_data['Quantity2']
        order_data['Amount2']=f'${price2:,.2f}'
        order_data['Description2']=product_descriptions.get(item2,'0')
    if pd.isnull(selected_sheet['B27'].value)==False:
        item3=selected_sheet['B27'].value
        order_data['Price3'] =product_prices.get(item3, '0')
        order_data['Quantity3'] =  selected_sheet['C27'].value
        price3 = order_data['Price3']* order_data['Quantity3']
        order_data['Amount3']=f'${price3:,.2f}'
        order_data['Description3']=product_descriptions.get(item3,'0')
    if pd.isnull(selected_sheet['B28'].value)==False:
        item4=selected_sheet['B28'].value
        order_data['Price4'] =product_prices.get(item4, '0')
        order_data['Quantity4'] =  selected_sheet['C28'].value
        price4 = order_data['Price4']* order_data['Quantity4']
        order_data['Amount4']=f'${price4:,.2f}'
        order_data['Description4']=product_descriptions.get(item4,'0')


    return order_data, order_data['CustomerName'], order_data['CustomerAddress2'], order_data['Date']

# Question 4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf',directory+f'Invoice {invoice_num}.pdf', order_data)

# Question 5
def send_email(invoice_num,name,email, date,directory):
    # Enter your gmail user name and password
    import smtplib
    from email.message import EmailMessage

    gmail_user = email
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    gmail_password = open(password_path).read()

    # List of recipients

    recipients = email

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.
    
    Thank you for your business.


    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    E-mail: sales@pythontechco.com"""
    

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)


    # Add attachment

    # path_to_file = os.getcwd()+'\\Invoice Template.xlsx'
    path_to_file = directory+ f'Invoice {invoice_num}.pdf' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


# Question 6
filenames= os.listdir(directory)
filenames.sort()
invoice_num=8515
for file in filenames:
    if "Order" in file and "xlsx" in file:
        order_data= get_order_data(invoice_num, file,  directory)
        
        name=order_data[1]
        email=order_data[2]
        date=order_data[3]
        create_invoice(invoice_num, order_data[0], directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num=invoice_num+1
