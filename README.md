# Fetch-Google-sheet-data-
Fetch Google sheet data and update in Microsoft Excel Sheet
# First, install the required libraries:
# !pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl pywin32

# Import the necessary libraries
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.oauth2 import service_account
from openpyxl import load_workbook
import win32com.client
from googleapiclient import errors
import xlsxwriter
SERVICE_ACCOUNT_FILE = 'myproject4086.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Authenticate and authorize the Python script to access the Google Sheets API
cred = None
cred = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Build the service object for interacting with the Google Sheets API
service = build('sheets', 'v4', credentials=cred)

# Update the Google Sheet with new data
sheet_id = '1rqsv6caQv-kveEz_FVaG8l2Dr-waJ9PRaJ-U0J-pN3w'  # Replace with the ID of your Google Sheet
range_name = 'data!A1:H16'  # The range of cells to update in the Google Sheet
values = [['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]] 
# The new values to insert into the Google Sheet
body = {
    'values': values
}

result = service.spreadsheets().values().update(
    spreadsheetId=sheet_id, range=range_name, valueInputOption='USER_ENTERED', body=body).execute()
print(f'{result["updatedCells"]} cells updated.')

# Fetch the updated data from the Google Sheet
result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=sheet_id, range="sheet1!A1:H16").execute()

values = result.get('values', [])
print(values)
# Open the Microsoft Excel sheet and update it with the data from the Google Sheet
with xlsxwriter.Workbook('new_sample4086.xlsx') as workbook:
    worksheet = workbook.add_worksheet()
    for row_num, data in enumerate(values):
        worksheet.write_row(row_num, 0, data)       
    
ws = wb['Sheet1']
for i, row in enumerate(values):
    for j, value in enumerate(row):
        ws.cell(row=i+1, column=j+1, value=value)
wb.save(excel_path)

# Automate Excel to refresh the data
excel = win32com.client.Dispatch("Excel.Application")
excel.Workbooks.Open(excel_path)
excel.Application.Run("RefreshAll")
excel.Application.Quit()
