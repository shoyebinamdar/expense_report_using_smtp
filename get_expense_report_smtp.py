#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jan 23 10:26:41 2019

@author: shoyeb
"""
import sys
import imaplib
import getpass
import email
import email.header
import datetime
import calendar
import re
from datetime import timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

EMAIL_ACCOUNT = "email-account-to-process"

EMAIL_FOLDER = "INBOX"

#HEADER_FILTER = ['paid', 'debited', 'transaction']
#Removing "paid" for nowas it complicates the conditions. Will add wallet debit in future verisons.
HEADER_FILTER = ['debited', 'transaction']
HEADER_EXCEPTIONS = ['credit transaction']

def send_mail(total_spendings, typeOfReport, timeVsDebitTransaction):
    gmail_user = 'sender-email'  
    gmail_password = 'sender-password'
    
    sent_from = gmail_user  
    to = "reciever-email" 
    subject = 'Daily Expense Summary'  
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = gmail_user
    msg['To'] = to
    
    text = "Hi!!!, \nYou can only view the expense report in html format.\nThanks,\nShofees Expense Tracker Bot"
    html = """<!DOCTYPE html>
            <html lang="en">
            <head>
              <title>Expense Tracker</title>
              <meta charset="utf-8">
              <meta name="viewport" content="width=device-width, initial-scale=1">
              <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/css/bootstrap.min.css">
              <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
              <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/js/bootstrap.min.js"></script>
              <style>
              table {
              font-family: arial, sans-serif;
              border-collapse: collapse;
              width: 40%;
            }
            
            th {
              background-color: #dddddd;
              border: 1px solid #000000;
              text-align: left;
              padding: 8px;
            }
            td {
              border: 1px solid #000000;
              text-align: left;
              padding: 8px;
            }
              </style>
            </head>
            <body>
            <div>
            <h3>Dear Shofee,</h3>
            </div>
            <h4>Your """ + typeOfReport + """ spending is : """ + str(total_spendings) + """</h4>
            <h5>Detailed analysis of """ + typeOfReport + """ expense is as below:</h5>
            <table>
            <thead >
            <tr><th>Time</th>
            <th>Amount</th>
            </tr>
            </thead>
            <tbody>"""
            
    for key, value in timeVsDebitTransaction.items():
        html += "<tr><td>" + key + "</td><td>" + value + "</td></tr>"
        
    html += "<tr><th>Total</th><th>Rs." + str(total_spendings) + "</th></tr>"
    html += """</tbody>
                </table>
                <br/>
                Thanks,
                <br/>
                Shofees Expense Tracker Bot
                </body>
                </html>
            """
    
    part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    
    msg.attach(part1)
    msg.attach(part2)
    
    try:  
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(sent_from, to, msg.as_string())
        server.quit()
    
        print('Email sent!')
    except Exception as e: 
        print('Something went wrong...')
        print(e)  
    
def is_valid_subject(subject):
    return any(find_word(subject.lower(), x) for x in HEADER_FILTER) and not any(find_word(subject.lower(), x) for x in HEADER_EXCEPTIONS)

def find_word(text, search):

   result = re.findall('\\b'+search+'\\b', text, flags=re.IGNORECASE)
   if len(result)>0:
      return True
   else:
      return False
  
def extract_amount_string(message):
    #matchObj = re.search(r'(₹|Rs|INR)(\s|\.)*[0-9]+,*[0-9]*\.*[0-9]*', re.UNICODE, message)
    matchObj = re.search(r'(&#x20B9;|Rs|INR)(\s|\.)*[0-9]+,*[0-9]*\.*[0-9]*\s+', message)
    if matchObj is not None and matchObj.group(0):
        return matchObj.group(0)
    else:
        return None
    
def extract_amount(message):
    matchObj = re.findall(r'([0-9]+,*[0-9]*\.*[0-9]*)', message)
    groupSize = len(matchObj)
    
    if matchObj is not None and matchObj[0]:
        if groupSize > 1:
            return matchObj[groupSize - 1]
        else:
            return matchObj[0]
    else:
        return None
    
def get_expense_report(M, dateSearchString, typeOfReport):
    rv, data = M.search(None, dateSearchString)
    if rv != 'OK':
        print("No messages found!")
        return
    
    timeVsDebitTransaction = dict()

    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return

        msg = email.message_from_bytes(data[0][1])
        if 'Subject' in msg:
            hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
            subject = str(hdr)
            entireMessage = subject
            amount = None
            print(subject)
            if is_valid_subject(subject):
                #print('Message %s: %s' % (num, subject))
                #print('Raw Date:', msg['Date'])
                if msg.is_multipart():
                    print("Multipart")
                    for payload in msg.get_payload():
                        
                        if type(payload.get_payload()) is list:
                            for payload_in in payload.get_payload():
                                entireMessage = entireMessage + payload_in.get_payload()
                        else:
                            entireMessage = entireMessage + payload.get_payload()
                        amount = extract_amount_string(entireMessage)
                else:
                    amount = extract_amount_string(msg.get_payload())
                print(msg.get_payload())
                # Now convert to local date-time
                date_tuple = email.utils.parsedate_tz(msg['Date'])
                if date_tuple:
                    local_date = datetime.datetime.fromtimestamp(
                        email.utils.mktime_tz(date_tuple))
                    #add data into dictionary
                    if amount is not None:
                        timeVsDebitTransaction[local_date.strftime("%a, %d %b %Y %H:%M:%S")] = amount
    
    total_spendings = 0.0
    for key, value in timeVsDebitTransaction.items():
        print(key + ", " + value)
        extracted_amount = extract_amount(value)
        print(extracted_amount)
        if extracted_amount is not None:
            total_spendings += float(extracted_amount.replace(",", ""))
    
    total_spendings = format(total_spendings, '.2f')
    print(total_spendings)
    send_mail(total_spendings, typeOfReport, timeVsDebitTransaction)
  
def process_mailbox(M):
    """
    Do something with emails messages in the folder.  
    For the sake of this example, print some headers.
    """

    # Daily report
    print("Generating daily expense report ...")
    todaysDate = datetime.datetime.now().date()
    yesterdaysDate = todaysDate - timedelta(1)
    dateSearchString = ('SINCE "%s"' % yesterdaysDate.strftime('%d-%b-%Y'))
    get_expense_report(M, dateSearchString, "Daily")    
    
    # Weekly report
    # SUNDAY is 6, so we generate weekly report on every Sunday.
    if datetime.datetime.today().weekday() == 6:
        print("Generating weekly expense report ...")
        weekStartDate = todaysDate - timedelta(6)
        dateSearchString = ('SINCE "%s"' % weekStartDate.strftime('%d-%b-%Y'))
        get_expense_report(M, dateSearchString, "Weekly")
    
    # Monthly report
    currentDayOfMonth = datetime.datetime.today().day
    
    if currentDayOfMonth == 1:
        print("Generating last months expense report ...")
        lastDayOfPreviousMonth = calendar.monthrange(datetime.datetime.today().year, datetime.datetime.today().month - 1)[1]
        startDateOfPreviousMonth = todaysDate - timedelta(lastDayOfPreviousMonth - 1)
        dateSearchString = ('SINCE "%s"' % startDateOfPreviousMonth.strftime('%d-%b-%Y'))
        get_expense_report(M, dateSearchString, "Monthly")

if len(sys.argv) != 2 :
    print("ERROR : Invalid arguments !!!")
    print("Usage: python get_expense_report.py <email-password>")
    sys.exit (1)
    
M = imaplib.IMAP4_SSL('imap.gmail.com')

try:
    rv, data = M.login(EMAIL_ACCOUNT, getpass.getpass())
    #rv, data = M.login(EMAIL_ACCOUNT, sys.argv[1])
except imaplib.IMAP4.error:
    print ("LOGIN FAILED!!! ")
    sys.exit(1)

print(rv, data)

rv, mailboxes = M.list()
if rv == 'OK':
    print("Mailboxes:")
    print(mailboxes)

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print("Processing mailbox...\n")
    process_mailbox(M)
    M.close()
else:
    print("ERROR: Unable to open mailbox ", rv)

M.logout()
