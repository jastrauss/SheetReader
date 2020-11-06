"""
    Based off the Google Sheets Python Quickstart
    https://developers.google.com/sheets/api/quickstart/python
"""

# from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from six import iteritems

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

class GoogleSheetsConnector:

    # If modifying these scopes, delete your previously saved credentials
    # at .credentials/sheets.googleapis.com-python-quickstart.json
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    CLIENT_SECRET_FILE = 'credentials.json'
    APPLICATION_NAME = 'Google Sheets API Python Quickstart'

    def __init__(self, spreadsheet_id, sheet_name):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.set_service()

    def get_credentials(self):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        current_dir = os.getcwd()
        credential_dir = os.path.join(current_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'sheets.googleapis.com-python-quickstart.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self.CLIENT_SECRET_FILE, self.SCOPES)
            flow.user_agent = self.APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def set_service(self):
        credentials = self.get_credentials()
        http = credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?' 'version=v4')
        self.service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)

    def read_range(self, start_cell, end_cell):
        rangeName = '%s!%s:%s' % (self.sheet_name, start_cell, end_cell)
        result = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id, range=rangeName).execute()
        values = result.get('values', [])
        return values

    def write_range(self, start_cell, end_cell, values=[[]]):
        rangeName = '%s!%s:%s' % (self.sheet_name, start_cell, end_cell)
        # Write the new values
        body = {
            'values': values
        }
        result = self.service.spreadsheets().values().update(
            spreadsheetId = self.spreadsheet_id,
            range = rangeName,
            # valueInputOption = "RAW",
            valueInputOption = "USER_ENTERED",
            body = body).execute()

    def bulk_write_range(self, cell_index_to_value_map):
        updates = []
        for cell_index, value in iteritems(cell_index_to_value_map):
            range_name = '%s!%s:%s' % (self.sheet_name, cell_index, cell_index)
            value_formatted = [[value]]
            updates.append({
                'range': range_name,
                'values': value_formatted
            })

        body = {
            'value_input_option': 'USER_ENTERED',
            'data': updates,
        }
        result = self.service.spreadsheets().values().batchUpdate(
            spreadsheetId = self.spreadsheet_id,
            body = body).execute()

