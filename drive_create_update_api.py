from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import csv

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']
CREDENTIALS = '' #Path to OAuth Credentials File
READ_CSV = '' #Path to CSV from where values are to be read
WRITE_CSV = '' #Path to CSV where logs are to be written
ORIGINAL_FILE = ''

def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    
    emailMessage = """
        Sample Email Message to be sent to the user.
    """

    with open(READ_CSV, newline='') as csvrfile:
        reader = csv.reader(csvrfile)
        with open(WRITE_CSV, 'w', newline='') as csvwfile:
            writer = csv.writer(csvwfile)
            row = list(reader)
            for r in row:
                file_name = r[0] #Column name from CSV to whic file is to be renamed.
                user1 = r[1].lower() #1st user to whom permission are to be allocated
                user2 = r[2].lower() #1st user to whom permission are to be allocated
                
                try:
                    #Create copy of the origin file. copy_file stores the response.
                    copy_file = service.files().copy(
                        fileId=ORIGINAL_FILE).execute()

                    #Rename the new file
                    updateMetaData = {
                        'name': name
                    }
                    updateName = service.files().update(
                        body=updateMetaData, fileId=copy_file['id']).execute()
                    
                    #Create Permission instance
                    batch = service.new_batch_http_request(callback=callback)
                    user1_permission = {
                        'type': 'user', #assign type
                        'role': 'writer', #assign role
                        'emailAddress': user1
                    }
                    batch.add(service.permissions().create(
                            fileId=copy_file['id'],
                            body=user1_permission,
                            fields='id',
                            emailMessage=emailMessage
                    ))
                    user2_permission = {
                        'type': 'user', #assign type
                        'role': 'writer', #assign role
                        'emailAddress': user2
                    }
                    batch.add(service.permissions().create(
                            fileId=copy_file['id'],
                            body=user2_permission,
                            fields='id',
                            emailMessage=emailMessage
                    ))
                    
                    #Execute batch permission update
                    batch.execute()
                    
                    #Sample successful log
                    writer.writerow([name, 'File copied & updated',copy_file['id']])
                except Exception as e:
                    writer.writerow([name, 'Error occured.', e])

if __name__ == '__main__':
    main()