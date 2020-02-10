import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from flask import Flask, render_template, request

app = Flask(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = '1UTJXmS_sBIY9X-GaiIqS_IdtlLO4flHP32nQIsiQ6Hk'
RAW_RANGE = 'Raw Data!A:C'
RESULT_RANGE = 'Result!A:A'


@app.route("/", methods=['GET', 'POST'])
def index():
    print(request.method)
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
                        'credentials.json', SCOPES)
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
            results = []
            if not values:
                print('No data found.')
            else:
                for row in values:

                    if row:
                        rows = ''
                        for i in range(len(row)):
                            # If customer does not enter anything in the first field it will write "Customer" there in it's place.
                            if i == 0:
                                if not row[0]:
                                    rows += 'Customer'
                                else:
                                    rows += '%s' % row[i].strip()

                            else:
                                rows += '\n%s' % row[i].strip()

                        results.append([rows])

            # print(results)
            resource = {
                "majorDimension": "ROWS",
                "values": results
            }

            service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID,
                                                  range=RESULT_RANGE).execute()

            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=RESULT_RANGE,
                body=resource,
                valueInputOption="USER_ENTERED"
            ).execute()

        elif request.form.get('docs') == 'Process Google Docs':
            # pass # do something else
            print("docs")
        else:
            # pass # unknown
            return render_template("index.html")
    elif request.method == 'GET':
        # return render_template("index.html")
        print("No Post Back Call")
    return render_template("index.html")


if __name__ == '__main__':
    app.run()
