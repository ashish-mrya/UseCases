# Get data from the google sheets

# We will be using gspred for accessing the Spredasheet.
# You can get the documentation at: https://docs.gspread.org/en/latest/

# Authentication
# For more details on authentication refer to the folloing page:
# https://docs.gspread.org/en/latest/oauth2.html

# importing the required libraries
import gspread
import smtplib
from oauth2client.service_account import ServiceAccountCredentials

# define the scope
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# add credentials to the account
# creds = ServiceAccountCredentials.from_json_keyfile_name('C:/Ashish Maurya/Projects/Python/GitHub/UseCases/myKey.json',scope)

# authorize the clientsheet
# client = gspread.authorize(creds)//To authorize by defining the key.json file
client = gspread.service_account()

# get the instance of the Spreadsheet
sheet = client.open('Send Email')

# get the first sheet of the Spreadsheet
sheet_instance = sheet.get_worksheet(0)
sheet1 = client.open('Send Email').sheet1

# Get the complete data
data = sheet1.get_all_values()

# Set the column for comparing value and column to be updated
col_update = "Status"
col_check = "Name"
header = data[0]
# print(header)
col_update_no = header.index(col_update) + 1

# Send Email
gmail_user = 'ashish.mrya@gmail.com'
gmail_password = ''  # later read it from file

sent_from = "ashish.mrya@gmail.com"
# to = ['ashish.mrya@gmail.com']
subject = 'OMG Super Important Message'
body = 'Hey, what\'s up?\n\n - You'
email_text = "This is a dummy email for testing"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()

'''
server.login(gmail_user, gmail_password)
print("Connection successful")
server.sendmail(sent_from, to, email_text)
server.close()
print("email successfully sent")
'''

# Send email getting the email from the Google sheets

# Get the list of email
email_col_no = header.index("EmailID")
status_col_no = header.index("Status")

# Send email for each email id in the EmailID column.
for row_id in range(len(data) - 1):
    row_no = row_id + 1
    email_id = data[row_no][email_col_no]
    status = data[row_no][status_col_no]
    print(email_id)
    print(status)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
    print("Connection successful")

    server.sendmail(sent_from, email_id, email_text)
    server.close()
    print(f"email succesfully sent to {email_id}")
    sheet1.update_cell(row_no + 1, status_col_no + 1,
                       "Email sent again")  # +1 is coz list indez start from 0 but sheet starts from 1