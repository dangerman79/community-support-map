from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1kG2WSS2O7iVHnBbmKBonsLXu4yxMSVQbm2WG5GZxLFo'
SAMPLE_RANGE_NAME = 'Form Responses 1!A5:J'

def main():
    setup()
    supportRequests = getSupportRequests()
    debugPrintData(supportRequests)
    
def setup():
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    global service
    service = build('sheets', 'v4', credentials=creds)

def getSupportRequests():


    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    data = []
    # print (values)
    for row in values:
        #if some rows don't have data at the end so errors occur about the array being out of bounds
        row = resize(row, 10)
        supportRequest = SupportRequest()
        
        # print('%s, %s' % (row[0], row[4]))
        supportRequest.timeStamp = row[0]
        supportRequest.firstName = row[1]
        supportRequest.telephone = row[2]
        supportRequest.postcode = row[4]
        supportRequest.requestText = row[5]
        supportRequest.paymentType = row[6]
        supportRequest.status = row[9]
        
        data.append(supportRequest)
        # Print columns A and E, which correspond to indices 0 and 4.
        # print('%s, %s' % (row[0], row[4]))
    return data    

# why does the class def need to be in the func? how do i globally define it?
class SupportRequest:
    timeStamp = ""
    firstName = ""
    telephone = ""
    postcode = ""
    requestText = ""
    paymentType = ""
    status = ""
    chat = ""
def resize(arr, size):
    if len(arr) >= size:
        return arr
    for x in range(len(arr), size):
        arr.append("")
    return arr

  
def debugPrintData(supportRequests):
    print('timeStamp, firstName, telephone, postcode, requestText, paymentType, status, chat')
    for sr in supportRequests:
       print('%s, %s, %s, %s, %s, %s, %s, %s' % (sr.timeStamp, sr.firstName, sr.telephone, sr.postcode, sr.requestText, sr.paymentType, sr.status, sr.chat)) 

  
if __name__ == '__main__':
    main()