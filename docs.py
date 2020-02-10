from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents']

# The ID of a sample document.
DOCUMENT_ID = '18xQzPIpvjLaB2fiDExlAmNH_sSjabeyD1Iq6e56BnkU'


def main():
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('docs.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)

    # Retrieve the documents contents from the Docs service.
    document = service.documents().get(documentId=DOCUMENT_ID).execute()

    print('The title of the document is: {}'.format(document.get('title')))

    if 2 < len(document['body']['content']):

        table = document['body']['content'][2]
        requests = [{
            'deleteContentRange': {
                'range': {
                    'segmentId': '',
                    'startIndex': table['startIndex'],
                    'endIndex':   table['endIndex']
                }
            },
        }
        ]

        result = service.documents().batchUpdate(documentId=DOCUMENT_ID,
                                                 body={'requests': requests}).execute()

    requests = [{
        'insertTable': {
            'rows': 10,
            'columns': 3,
            'location': {
                "segmentId": '',
                "index": 1
            }
        }
    }, {
        'insertText': {
            'location': {
                'index': 5
            },
            'text': 'Hello'
        }
    }, {
        'insertText': {
            'location': {
                'index': 10
            },
            'text': 'World'
        }
    }
    ]

    result = service.documents().batchUpdate(documentId=DOCUMENT_ID,
                                             body={'requests': requests}).execute()


if __name__ == '__main__':
    main()
