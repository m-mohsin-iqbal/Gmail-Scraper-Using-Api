import csv
import os
import re
import gspread
from googleapiclient.discovery import build
from gspread.exceptions import SpreadsheetNotFound
from httplib2 import Http
import logging
from oauth2client import file, client, tools

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]


class GmailScraper():
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.sh = logging.StreamHandler()
        self.log_format = "%(levelname)s : %(asctime)s : %(message)s"
        self.formatter = logging.Formatter(self.log_format)
        logging.basicConfig()
        self.sh.setFormatter(self.formatter)
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(self.sh)

    
    def insert_data_into_google_sheet(self, values, google_sheet_file):
        gc = gspread.oauth(
            credentials_filename='credentials.json',
            authorized_user_filename='token.json'
        )
        # Verifying Existing Sheet
        try:
            worksheet = gc.open(google_sheet_file).sheet1
            worksheet.append_rows(values)
        # Sheet Not Found Trying to Create New Sheet
        except SpreadsheetNotFound as e:
            choice = input("Your give Sheet is not available. Do you want to create a new Sheet?\n"
                  "If Yes then Press Y\t")
            if choice.lower() == 'y':
                # Create A New Spreadsheet
                sh = gc.create(google_sheet_file)
                worksheet = sh.sheet1
                worksheet.append_row(['Employee_id', 'date', 'body'])
                worksheet.append_rows(values)
            else:
                self.logger.info("As you have not Entered Y so quiting...")
                return

    def insert_data_into_csv(self, item):
        #Data insertion in CSV
        filename = 'Data.csv'
        file_exists = os.path.isfile(filename)

        with open(filename, 'a', encoding="utf-8") as csvfile:
            headers = item.keys()
            writer = csv.DictWriter(csvfile, delimiter=',',
                                    lineterminator='\n',
                                    fieldnames=headers)
            if not file_exists:
                writer.writeheader()  # file doesn't exist yet, write a header
            writer.writerow(item)

    def parse_emails(self):
        self.logger.info("program started")
        values = []
        store = file.Storage('token.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
            creds = tools.run_flow(flow, store)
        service = build('gmail', 'v1', http=creds.authorize(Http()))

        # Call the Gmail API to fetch INBOX
        results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
        messages = results.get('messages', [])
        if not messages:
            self.logger.info("No messages found.")
        else:
            google_sheet_file = input("Please Enter the Name of Sheet you want to Populate\t")
            self.logger.info("Message snippets:")
            for message in messages:
                msg = service.users().messages().get(userId='me', id=message['id']).execute()
                employee_id = ''
                date = ''
                body = ''
                for header in msg['payload']['headers']:
                    if 'Subject' in header['name']:
                        item = dict()
                        subject = header['value']
                        match = re.search(r"[a-zA-Z]+[ -]?(\d[\d-]+)", subject)
                        if match:
                            employee_id = match.group(1)
                    if 'Date' in header['name']:
                        date = header['value']
                if employee_id:
                    body = msg['snippet']
                    item = dict(
                        empployee_id=employee_id,
                        date=date,
                        body=body
                    )
                    # insert data into csv file
                    self.insert_data_into_csv(item)
                    # insert data into  Google Sheet
                    values.append([employee_id, date, body])
            self.insert_data_into_google_sheet(values, google_sheet_file)


gs = GmailScraper()
gs.parse_emails()
