import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
import os

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
my_file = os.path.join(THIS_FOLDER, 'client_secret.json')

creds = ServiceAccountCredentials.from_json_keyfile_name(my_file, scope)
client = gspread.authorize(creds)


sheet1 = client.open('PROPOSAL EMAIL LIST - MailChimp Synchronized').get_worksheet(0)

pp = pprint.PrettyPrinter()
result = sheet1.get_all_records()