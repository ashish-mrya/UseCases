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



gmail_user = username
gmail_password = password

    ########################################################################

def send_selenium_report():
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
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.close()
    print("email succesfully sent")


send_selenium_report()