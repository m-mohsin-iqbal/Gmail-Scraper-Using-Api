import gspread
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import re
import os
import csv

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

class GmailScraper():
    def insert_data_into_google_sheet(self, values):
        gc = gspread.oauth(
            credentials_filename='credentials.json',
            authorized_user_filename='token.json'
        )

        # Open a sheet from a spreadsheet in one go
        wks = gc.open("data").sheet1
        wks.append_row(values)

    def insert_data_into_csv(self, item):
        filename = 'data.csv'
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
            print("No messages found.")
        else:
            print("Message snippets:")
            for message in messages:
                msg = service.users().messages().get(userId='me', id=message['id']).execute()
                employee_id = ''
                date = ''
                body = ''
                for header in msg['payload']['headers']:
                    if 'Subject' in header['name']:
                        item = dict()
                        subject = header['value']
                        match = re.search(r"[a-zA-Z]+[ -]?(\d[\d-]+)",subject)
                        if match:
                            employee_id = match.group(1)
                    if 'Date' in header['name']:
                        date = header['value']
                if employee_id:
                    body = msg['snippet']
                    item = dict(
                        employee_id=employee_id,
                        date=date,
                        body=body
                    )
                    # insert data into csv file
                    self.insert_data_into_csv(item)
                    # insert data into  Google Sheet
                    values = [employee_id, date, body]
                    self.insert_data_into_google_sheet(values)


gs = GmailScraper()
gs.parse_emails()
