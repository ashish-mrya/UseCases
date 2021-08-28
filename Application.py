
# Get data from the google sheets
# We will be using gspred for accessing the Spredasheet.
# You can get the documentation at: https://docs.gspread.org/en/latest/

# Authentication
# For more details on authentication refer to the folloing page:
# https://docs.gspread.org/en/latest/oauth2.html

# importing the required libraries
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders
import json


#read the config file for the mail configuration
#opening JSON file
with open("C:/Ashish Maurya/Projects/Python/GitHub/UseCases/config.json") as f:
    config = json.load(f)
#retrun JSON object as a dictionary


#get senders details
sendersDetails = config["sendersEmailDetails"]
username = sendersDetails["username"]
password = sendersDetails["password"]

#get google Sheet Details
googleSheetDetails = config["googleSheetDetails"]
workbookName = googleSheetDetails["workbookName"]
sheetName = googleSheetDetails["sheetName"]

#get email Details
emailDetails = config["emailDetails"]
sent_from = emailDetails["from"]
sent_to= emailDetails['to']
subject = emailDetails["subject"]
attachmentsPath = emailDetails['attachmentsPath']

#gmail service account login user gamil account details
gmail_user = username
gmail_password = password


# define the scope
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# add credentials to the account
# creds = ServiceAccountCredentials.from_json_keyfile_name('C:/Ashish Maurya/Projects/Python/GitHub/UseCases/myKey.json',scope)

# authorize the clientsheet
# client = gspread.authorize(creds)
client = gspread.service_account()

# get the instance of the Spreadsheet
sheet = client.open('Send Email')

# get the first sheet of the Spreadsheet
sheet_instance = sheet.get_worksheet(0)

sheet1 = client.open('Send Email').sheet1

# Get the complete data
data = sheet1.get_all_values()


# Set the coloumn for comparing value and column to be updated
col_update = "Status"
col_check = "Name"
header = data[0]
col_update_no = header.index(col_update) + 1
col_check_no = header.index(col_check) + 1
print(header)



# Send Email
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(gmail_user, gmail_password)
print("Connection successful")

server.sendmail(sent_from, sent_to, "Test Mail")
server.close()
print("email succesfully sent")

print("Mail Sent")

# Send email getting the email from the Google sheets
# Get the maximum row count, later write it to a file to make it persistance

# Get the list of email
email_col_no = header.index("EmailID")
status_col_no = header.index("Status")

# Send email Id for each email id in the email list.
def send_mail(start=0):
    global old_row_count
    for row_id in range(start, len(data) - 1):
        row_no = row_id + 1  # row_id represent the index and not the row number
        email_id = data[row_no][email_col_no]
        status = data[row_no][status_col_no]
        print(email_id)
        print(status)
        ########################################################################
        print("starting")
        dir_path = "C:\Ashish Maurya\Projects\Python\GitHub\AttachmentFiles"
        files = ["file1.pdf", "file2.pdf"]

        msg = MIMEMultipart()
        msg['To'] = sent_to
        msg['From'] = sent_from
        msg['Subject'] = subject

        body = MIMEText('This is the test message for body')
        msg.attach(body)  # add message body (text or html)

        for f in attachmentsPath:  # add files to the message
            # file_path = os.path.join(dir_path, f)
            # attachment = MIMEApplication(open(file_path).read())
            # attachment.add_header('Content-Disposition','attachment', filename=f)
            # msg.attach(attachment)
            f = os.path.join(dir_path, f)
            part = MIMEBase('application', "octet-stream")
            part.set_payload(open(f, "rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
            msg.attach(part)
        # s = smtplib.SMTP()
        # s.connect(host=SMTP_SERVER)
        # s.sendmail(msg['From'], msg['To'], msg.as_string())
        # print 'done!'
        # s.close()

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        print("Connection successful")
        server.sendmail(msg['From'], email_id, msg.as_string())
        server.close()
        print("email succesfully sent to ")
        #############################################################################

        # server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        # server.ehlo()
        # server.login(gmail_user, gmail_password)
        # server.sendmail(sent_from, email_id, email_text)
        # server.close()
        # print(f"email succesfully sent to {email_id}")
        # update sheet after sendimg email
        sheet1.update_cell(row_no + 1, status_col_no + 1,
                            "Email sent again")  # +1 is coz list indez start from 0 but sheet starts from 1

        # update the last_index
        old_row_count = row_no + 1
        print(f'row count {old_row_count}')
    print(f'total no of after sending email is {old_row_count}')


#Initiate this funciton once at the begining of the Intial phase
send_mail()

while True:
    time.sleep(10)
    data = sheet1.get_all_values()
    updated_row_count = len(data)
    print(f'old count = {old_row_count}, new count = {updated_row_count}')

    if updated_row_count > old_row_count:
        send_mail(old_row_count - 1)
    else:
        continue



