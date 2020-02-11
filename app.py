import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from gdoctableapppy import gdoctableapp

from flask import Flask, render_template, request

app = Flask(__name__)

DOCUMENT_SCOPES = ['https://www.googleapis.com/auth/documents']
SPREADSHEET_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

DOCUMENT_ID = '18xQzPIpvjLaB2fiDExlAmNH_sSjabeyD1Iq6e56BnkU'
SPREADSHEET_ID = '1UTJXmS_sBIY9X-GaiIqS_IdtlLO4flHP32nQIsiQ6Hk'
RAW_RANGE = 'Raw Data!A:C'
RESULT_RANGE = 'Result!A:A'


@app.route("/", methods=['GET', 'POST'])
def index():
    message = ''
    if request.method == 'POST':
        if request.form.get('sheet') == 'Process Google Sheet':

            """Shows basic usage of the Sheets API.
            Prints values from a sample spreadsheet.
            """
            creds = None
            # The file token.pickle stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first
            # time.
            if os.path.exists('sheet.pickle'):
                with open('sheet.pickle', 'rb') as token:
                    creds = pickle.load(token)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'credentials.json', SPREADSHEET_SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('sheet.pickle', 'wb') as token:
                    pickle.dump(creds, token)

            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                        range=RAW_RANGE).execute()

            values = result.get('values', [])
            sheets, docs, columns = [], [], []

            if values:
                for row in range(len(values)):

                    if values[row]:
                        rows = ''
                        for i in range(len(values[row])):
                            # If customer does not enter anything in the first field it will write "Customer" there in it's place.
                            if i == 0:
                                if not values[row][0]:
                                    rows += 'Customer'
                                else:
                                    rows += '%s' % values[row][i].strip()

                            else:
                                rows += '\n%s' % values[row][i].strip()

                        sheets.append([rows])

                        # TODO: You need convert list to numpy array and then reshape:

                        if (row+1) % 3 == 0:
                            columns.append(rows)
                            docs.append(columns)
                            columns = []
                        else:
                            columns.append(rows)
                            if len(values) == (row+1):
                                # docs.append(columns)

            resource = {
                "majorDimension": "ROWS",
                "values": sheets
            }

            # Reset Sheet
            service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID,
                                                  range=RESULT_RANGE).execute()

            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=RESULT_RANGE,
                body=resource,
                valueInputOption="USER_ENTERED"
            ).execute()

        # elif request.form.get('docs') == 'Process Google Docs':

            """Shows basic usage of the Docs API.
            Prints the title of a sample document.
            """
            creds = None
            # The file token.pickle stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first
            # time.
            if os.path.exists('docs.pickle'):
                with open('docs.pickle', 'rb') as token:
                    creds = pickle.load(token)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        'credentials.json', DOCUMENT_SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('docs.pickle', 'wb') as token:
                    pickle.dump(creds, token)

            service = build('docs', 'v1', credentials=creds)

            resource = {
                "oauth2": creds,
                "documentId": DOCUMENT_ID
            }
            # You can see the retrieved values like this.
            document = gdoctableapp.GetTables(resource)

            # Reset Table
            if document['tables']:
                resource = {
                    "oauth2": creds,
                    "documentId": DOCUMENT_ID,
                    "tableIndex": 0
                }
                gdoctableapp.DeleteTable(resource)

            resource = {
                "oauth2": creds,
                "documentId": DOCUMENT_ID,
                "rows": len(docs),
                "columns": 3,
                "createIndex": 1,
                "values": docs
            }
            gdoctableapp.CreateTable(resource)
            message = 'Google Sheet and Docs Processed. Please see the links below'
        else:
            # pass # unknown
            return render_template("index.html")

    return render_template("index.html", message=message)


if __name__ == '__main__':
    app.run()
