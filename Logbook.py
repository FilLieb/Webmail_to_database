import imapclient
import email
from email.header import decode_header
import openpyxl
import pyautogui

#functions to receive information
def find_between(text, first, last ):
    try:
        start = text.index( first ) + len( first )
        end = text.index( last, start )
        return text[start:end]
    except ValueError:
        return ""

def substring_after(text, after):
    return text.partition(after)[2]

# create variables with info
def fetch_data(text):
    name = find_between(text, "Name ", "Date")
    date = find_between(text, "Date ", "Column")
    column = find_between(text, "used ", "#")
    runs = find_between(text, "runs ", "Buffer")
    pressure = find_between(text, "pressure ", "Buffer")
    flow = find_between(text, "rate ", "Column")
    clean = find_between(text, "(NaOH)? ", "Solution")
    solution = find_between(text, "equilibrated in? ", "Errors/Comments")
    comments = substring_after(text, "Errors/Comments ")
    new_data = [[name, date, column, runs, pressure, flow, clean, solution, comments]]
    return new_data

def append_sheet(email):
    new = fetch_data(email)
    # Open existing workbook and select the active worksheet
    workbook = openpyxl.load_workbook('ColumnLogbook_QR.xlsx')
    sheet = workbook.active
    # Append new data
    for row in new:
        sheet.append(row)
    # Save the workbook
    workbook.save('ColumnLogbook_QR.xlsx')



# Define the email server and login details
EMAIL_HOST = 'imap.uni-koeln.de'
EMAIL_USER = pyautogui.password(text='Enter username: ', title='username', default='', mask='')
EMAIL_PASS = pyautogui.password(text='Enter password: ', title='password', default='', mask='*')

# Connect to the server
server = imapclient.IMAPClient(EMAIL_HOST, ssl=True)
server.login(EMAIL_USER, EMAIL_PASS)

# Select the mailbox you want to search in
server.select_folder('INBOX')

# Search for emails
# You can customize the search criteria as per your requirements
messages = server.search(['FROM', 'nobody@uni-koeln.de'])

# Fetch and parse the emails
for uid in messages:
    raw_message = server.fetch(uid, ['BODY[]', 'FLAGS'])
    email_message = email.message_from_bytes(raw_message[uid][b'BODY[]'])

    # Get the email body
    body = ''
    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode(part.get_content_charset())
                break
    else:
        body = email_message.get_payload(decode=True).decode(email_message.get_content_charset())

    # Append the email data to the Excel sheet
    append_sheet(body)

# Logout from the server
server.logout()
