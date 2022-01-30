#!/usr/bin/env python

import csv
import os.path

from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class SheetsAPI:

    def __init__(
        self,
        scopes,
        spreadsheet_id
    ):
        self.scopes = scopes
        self.spreadsheet_id = spreadsheet_id

    def get_credentials(self, cred_path, token_path):
        '''Runs oauth flow to obtain credentials for using the Google API
        Visit https://developers.google.com/workspace/guides/create-credentials#desktop-app
        for more information on creating the necessary access credentials

        PARAMS
        ------------
        cred_path : path to your secret credenetials generated using the google
        cloud console. These will be used to complete the oauth flow and generate
        access and refresh tokens if they do not yet exist

        token_path : path to your access and refresh tokens. If the tokens exits,
        the function will look for them at this path, otherwise the tokens will 
        be created and saved at this path
        '''
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
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
        '''comment string for function

        '''

        try:
            service = build('sheets', 'v4', credentials=credentials)

        except HttpError as err:
            print(err)
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
        '''Comment string for function

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
            print(err)

        return

    def last_updated(self, service, sheet_range):
        '''function comment string

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

