# We will be using gspred for accessing the Spredasheet.
# You can get the documentation at: https://docs.gspread.org/en/latest/

# Authentication
# For more details on authentication refer to the folloing page:
# https://docs.gspread.org/en/latest/oauth2.html

# importing the required libraries
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# define the scope
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('C:/Ashish Maurya/Projects/Python/GitHub/UseCases/myKey.json',
                                                         scope)

# authorize the clientsheet
#client = gspread.authorize(creds)
client = gspread.service_account()

# get the instance of the Spreadsheet
sheet = client.open('Send Email')

# get the first sheet of the Spreadsheet
sheet_instance = sheet.get_worksheet(0)

sheet1 = client.open('Send Email').sheet1

#Get the complete data
data = sheet1.get_all_values()

# Set the coloumn for comparing value and column to be updated
col_update = "Status"
col_check = "Name"
header = data[0]
col_update_no = header.index(col_update)+1
col_check_no = header.index(col_check)+1
print(f'')

for row_id in range(len(data)):
    if data[row_id][col_check_no] == "":
        sheet1.update_cell(row_id+1, col_update_no, "As Hell missing")
        print(f"{row_id},{col_update_no}")
        print("Missing")








