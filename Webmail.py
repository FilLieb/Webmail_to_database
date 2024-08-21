import imapclient
import email
from email.header import decode_header
import openpyxl
import pyautogui


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

# Create a new workbook and select the active worksheet
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.append(['Subject', 'From', 'Date', 'Body'])

# Fetch and parse the emails
for uid in messages:
    raw_message = server.fetch(uid, ['BODY[]', 'FLAGS'])
    email_message = email.message_from_bytes(raw_message[uid][b'BODY[]'])

    # Decode the email subject
    subject, encoding = decode_header(email_message['Subject'])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding if encoding else 'utf-8')

    # Decode the email sender
    sender, encoding = decode_header(email_message.get('From'))[0]
    if isinstance(sender, bytes):
        sender = sender.decode(encoding if encoding else 'utf-8')

    # Decode the email date
    date = email_message.get('Date')

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
    sheet.append([subject, sender, date, body])

# Save the workbook
workbook.save('emails.xlsx')

# Logout from the server
server.logout()
