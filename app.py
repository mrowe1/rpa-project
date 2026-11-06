# Insert your python code here

# IMPORT PACKAGES IN THIS CELL

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# Question 1
# INSERT ANSWER FOR QUESTION 1 HERE
def get_dir():
    directory = select_file('Please select a file in the directory you would like to select.')
    split_directory = directory.split('/')
    del split_directory[-1]
    directory = '/'.join(split_directory)
    directory = directory + '/'
    return directory


while True:
    directory = get_dir()
    files = os.listdir(directory)
    if 'ProductData.xlsx' in files and 'MasterInvoice.pdf' in files and 'Order1.xlsx'in files and "Order2.xlsx" in files and 'Order3.xlsx' in files and 'Order4.xlsx' in files:
        break
    else:
        show_message('Error', 'Please Enter the directory where the order files, product data file, and master invoice pdf are found.')
            

print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE

product_descriptions = {}
product_prices = {}

product_data = directory + 'ProductData.xlsx'


def excel_to_dict(file, sheet, colname):
    df = pd.read_excel(file, sheet_name= sheet)
    new_dict = {}
    for _, row in df.iterrows():
        new_dict[row['Unit number']] = row[colname]
    return new_dict


product_descriptions = excel_to_dict(product_data, 'Descriptions', 'Description')
product_prices = excel_to_dict(product_data, 'Prices', 'Price per unit')


product_descriptions
product_prices

# INSERT ANSWER FOR QUESTION 3 HERE


def get_order_data(invoice_num, filename, directory= directory):
    # Initialize new dictionary to return order data
    updated_dict = {}

    # Naming new dictionary as the invoice number variable passed in
    updated_dict['InvoiceNum'] = str(invoice_num)

    # Open the order workbook and pull values based on cell positions, simultaneously update new dictionary
    wb = openpyxl.load_workbook(directory + filename, data_only=True)
    ws = wb['OrderForm']
    
    order_date = ws['C9'].value
    month = order_date.month
    day = order_date.day
    year = order_date.year
    updated_dict['Date'] = f'{month}/{day}/{year}'
    date = updated_dict['Date']
    
    email = ws['C13'].value

    cust_name = ws['C14'].value
    updated_dict['CustomerName'] = cust_name

    comp_name = ws['C15'].value
    updated_dict['CustomerCompany'] = comp_name

    street_address = ws['C16'].value
    updated_dict['CustomerAddress1'] = street_address

    city = ws['C17'].value
    state = ws['C18'].value
    zipcode = ws['C19'].value
    updated_dict['CustomerAddress2'] = f'{city}, {state} {zipcode}'

    # Initialize items needed to loop correctly
    start_row = 24
    curr_row = start_row
    invoice_total = 0
    while True:
        # using the column and the row number, grab the cell of the order info
        B_cell = 'B' + str(curr_row)
        C_cell = 'C' + str(curr_row)
        if ws[B_cell].value is None:
            break
            # If None, breaks out of the While True indicating that the row is empty
        else:
            unit_num = ws[B_cell].value
            qty = ws[C_cell].value
            qtystr = str(qty)
                
            # find the price for each item ordered and add it to the invoice total
            unit_price = product_prices[unit_num]
            unit_pricestr = f'${unit_price:,.2f}'
            unit_description = product_descriptions[unit_num]
            item_total = qty * unit_price
            item_totalstr = f'${item_total:,.2f}'
            invoice_total = invoice_total + item_total
            invoice_totalstr = f'${invoice_total:,.2f}'


            # Attempting to create a dictionary where I can easily return all the data related to each order
            item_id = curr_row - 24
            updated_dict[f'Description{item_id}'] = unit_description
            updated_dict[f'Price{item_id}'] = unit_pricestr
            updated_dict[f'Quantity{item_id}'] = qtystr
            updated_dict[f'Amount{item_id}'] = item_totalstr


            # numstr = f'${num:,.2f}
            # iterates through rows until it finds an empty row at which point it should break out of While True
            curr_row = curr_row + 1
    updated_dict['Total'] = invoice_totalstr
    return updated_dict, comp_name, email, date

# order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
# display(order_data)

# INSERT ANSWER FOR QUESTION 4 HERE
Master_invoice_path = directory + 'MasterInvoice.pdf'

def create_invoice(invoice_num, order_data, directory= directory):
    create_pdf(Master_invoice_path, directory + f'Invoice {invoice_num}.pdf', order_data)
    

# create_invoice(8516, 'Order4.xlsx')

# INSERT ANSWER FOR QUESTION 5 HERE


def send_email(invoice_num, name, email, date, directory= directory):
    
    gmail_user = 'mrowe4554@gmail.com'
    gmail_password = open("C:\\Users\\mrowe\\School\\2026-2027 (Y1 MAcc)\\Fall 2026\\Analytics (Python)\\Advanced Analytics\\gmail_password.txt").read()
    invoice_subject = f'Invoice {invoice_num}'

    # List of recipients
    recipients = email

    # Email subect
    subject = f'Your Order from PythonTech Co. - {invoice_subject}'

    # Body of the email
    body = f"""Dear {name},

    Attached is your invoice dated {date}.

    Thank you for your business.

    Python Tech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    
    message = EmailMessage()
    message['From'] = gmail_user # You can change the gmail_user to be any text string and it will alias the from (will still have email next to alias)
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


# send_email(8515, 'Thomas Holt', 'mrowe4554@gmail.com', '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
invoice_num = 8515
filenames = os.listdir(directory)
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
