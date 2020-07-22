from __future__ import print_function
import pickle
import os.path
import csv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import time

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS = '' #Path to OAuth Credentials File
READ_CSV = '' #Path to CSV from where values are to be read
WRITE_CSV = '' #Path to CSV where logs are to be written

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    requests = []
    requests.append({
        #Sample Request to update data format of cells on a sheet
        "repeatCell": {
            "cell": {
            "userEnteredFormat": {
                "numberFormat": {
                "type": "DATE",
                "pattern": "mm/dd/yyyy"
                }
            }
            },
            "range": {
            "sheetId": 0,
            "endColumnIndex": 7,
            "endRowIndex": 441,
            "startColumnIndex": 5,
            "startRowIndex": 31
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    })

    body = {
        'requests': requests
    }

    with open(READ_CSV, newline='') as csvrfile:
        reader = csv.reader(csvrfile)
        with open(WRITE_CSV, 'w', newline='') as csvwfile:
            writer = csv.writer(csvwfile)
            row = list(reader)
            for r in row:
                time.sleep(4) #Delay to be added to keep the number of writes below the Google's permittedd limits
                spreadsheetId = r[1] #Column in the CSV which has the required Spreadsheet ID
                try:
                    response = service.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheetId,
                        body=body).execute()
                    print(response)
                    #Sample Successful Log
                    writer.writerow([r[0], 'Updated'])
                except Exception as e:
                    print(e)
                    #Sample Error Log
                    writer.writerow([r[0], 'Not Updated', e])


if __name__ == "__main__":
    main()