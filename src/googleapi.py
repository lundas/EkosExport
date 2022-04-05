#!/usr/bin/env python

import csv
import logging
import os.path

from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
# Handler
log_path = '' # path to log file
fh = logging.FileHandler('{}deliveries.log'.format(log_path))
fh.setLevel(logging.INFO)
# Formatter
formatter = logging.Formatter(
    '%(asctime)s : %(name)s : %(levelname)s : %(message)s'
)
fh.setFormatter(formatter)
logger.addHandler(fh)

class SheetsAPI:

    def __init__(
        self,
        scopes,
        spreadsheet_id
    ):
        '''Class of functions for interacting with the Google Sheest API. Initialized
        using the required API scopes and well as with the spreadsheet_id for the 
        Google Sheets Spreadsheet Resource to be modified

        PARAMS
        ------------
        scopes : OAuth 2.0 scopes that determine level of access to the API. Best 
        practice is to provide the minimum level of access necessary to perform
        the necessary tasks.

            'https://www.googleapis.com/auth/drive' : See, edit, create, and delete
                all of your Google Drive files

            'https://www.googleapis.com/auth/drive.file' : View and manage Google
                Drive files and folders that you have opened or created with this app

            'https://www.googleapis.com/auth/drive.readonly' : See and download all
                your Google Drive files

            'https://www.googleapis.com/auth/spreadsheets' : See, edit, create, and 
                delete your spreadsheets in Google Drive

            'https://www.googleapis.com/auth/spreadsheets.readonly' : View your Google Spreadsheets

        spreadsheet_id : Value containing letters, numbers, hyphers, or underscores
        that is used to identify a Spreadsheet resourece; it can be found in the 
        Google Sheets URL

            e.g. https://docs.google.com/spreadsheets/d/spreadsheet_id/edit#gid=0
        '''
        self.scopes = scopes
        self.spreadsheet_id = spreadsheet_id

    def get_credentials(self, cred_path, token_path):
        '''Runs OAuth 2.0 flow to obtain credentials for using the Google API
        Visit https://developers.google.com/workspace/guides/create-credentials#desktop-app
        for more information on creating the necessary access credentials

        PARAMS
        ------------
        cred_path : PATH to your secret credenetials generated using the google
        cloud console. These will be used to complete the oauth flow and generate
        access and refresh tokens if they do not yet exist

        token_path : PATH to your access and refresh tokens. If the tokens exists,
        the function will look for them at this location, otherwise the tokens will 
        be created and saved at this location
        '''
        creds = None
        if os.path.exists(token_path):
            creds = Credentials.from_authorized_user_file(token_path, self.scopes)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    cred_path, self.scopes)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(token_path, 'w') as token:
                token.write(creds.to_json())

        return creds

    def get_service(self, credentials):
        '''Creates service of Google Sheets API using the credentials created
        with get_credentials function.

        Sheets API v4

        PARAMS
        ------------
        credentials : OAuth 2.0 credentials created using get_credentials function
        '''
        try:
            service = build('sheets', 'v4', credentials=credentials)

        except HttpError as err:
            logging.exception(err)
        return service

    def import_data(
        self,
        service,
        data,
        sheet_range,
        major_dimension='ROWS',
        value_input_option='USER_ENTERED',
        clear=True
    ):
        '''Import data to Google Sheet using Google Sheet API. Uses servie created
        with get_service function

        PARAMS
        ---------------
        service : Google Sheets service created using get_service function

        data : data to be imported to Google Sheet

        sheet_range : range of cells insert data into within the Google Sheet,
        provided in A1 notation

        major_dimension : Dimension used to feed data records from csv into sheet.

            'ROWS' will feed data records in as rows 

            'COLUMNS' will feed them in as columns

        value_input_option : Determines how values are treated after they are written
        to cells

            'RAW' == the input is not parsed and is simply inserted as a string

                e.g. the input "=1+2" is entered as the string "=1+2" rather than as
                as formula

            'USER_ENTERED' == the input is parsed exactly as if it were entered into the
            Google Sheets UI

                e.g. "Mar 1 2016" becomes a date and "=1+2" becomes a formula

        clear : If true, clears cells in sheet_range before writing new values to
        cells. Ensures that no old data persists in the sheet when writing new data
        to the same sheet_range
        '''
        # Open and read csv data as list
        data = list(csv.reader(open(data)))
        # Populate Google Sheet with csv data
        body = {
            'values' : data,
            'majorDimension' : major_dimension
        }
        try:
            if clear == True: # clear values in sheet if cleer == True
                request = service.spreadsheets().values().clear(
                    spreadsheetId = self.spreadsheet_id,
                    range = sheet_range,
                    body = {}
                )
                result = request.execute()

            request = service.spreadsheets().values().update(
                spreadsheetId = self.spreadsheet_id,
                range = sheet_range,
                valueInputOption = value_input_option,
                body = body
            )
            result = request.execute()

        except HttpError as err:
            logger.exception(err)

        return

    def last_updated(self, service, sheet_range):
        '''Enters the current datetime into a provided sheet_range to allow
        users to quickly determine when the Google Sheet was last updated
        by the program

        PARAMS
        -----------
        service : Google Sheets service created using get_service function

        sheet_range : range of cells insert data into within the Google Sheet,
        provided in A1 notation
        '''
        today = [[str(datetime.today())]] # list of lists required by sheets API
        body = {
            'values' : today,
            'majorDimension' : 'ROWS'
        }

        request = service.spreadsheets().values().clear(
            spreadsheetId = self.spreadsheet_id,
            range = sheet_range,
            body = {}
        )
        result = request.execute()

        request = service.spreadsheets().values().update(
            spreadsheetId = self.spreadsheet_id,
            range = sheet_range,
            valueInputOption = 'USER_ENTERED',
            body = body
        )
        result = request.execute()

        return

if __name__ == '__main__':
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    # Credentials
    cred_path = ''
    token_path = ''

    # The ID and range of a spreadsheet.
    SPREADSHEET_ID = ''
    DATA_RANGE_NAME = ''
    INFO_RANGE_NAME = ''

    # Ekos Data
    report = ''

    sheets_api = SheetsAPI(
        scopes = SCOPES, 
        spreadsheet_id = SPREADSHEET_ID
    )

    credentials = sheets_api.get_credentials(cred_path,token_path)
    service = sheets_api.get_service(credentials)
    sheets_api.import_data(
        service = service,
        data = report,
        sheet_range = DATA_RANGE_NAME
    )
    sheets_api.last_updated(service = service, sheet_range = INFO_RANGE_NAME)

