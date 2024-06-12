import httplib2
import googleapiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


class GoogleService:
    """Класс для работы с Google Cloud"""

    auth_request = None
    excel_service = None
    drive_service = None
    list_name = 'Лист1'

    def __init__(self, path_to_google_credits):
        google_credits = ServiceAccountCredentials.from_json_keyfile_name(path_to_google_credits,
            ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        self.auth_request = google_credits.authorize(httplib2.Http())
        self.excel_service = googleapiclient.discovery.build('sheets', 'v4', http=self.auth_request)
        self.drive_service = googleapiclient.discovery.build('drive', 'v3', http=self.auth_request)

    def create_excel_file(self, file_name, title):
        spreadsheet = self.excel_service.spreadsheets().create(body={
            'properties': {'title': file_name, 'locale': 'ru_RU'},
            'sheets': [{'properties': {'sheetType': 'GRID',
                'sheetId': 0,
                'title': title,
                'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
        }).execute()
        file_id = spreadsheet['spreadsheetId']
        return file_id

    def grant_access_to_file(self, file_id, gmail):
        access = self.drive_service.permissions().create(
            fileId=file_id,
            body={'type': 'user', 'role': 'writer', 'emailAddress': gmail},
            fields='id'
        ).execute()
        return access

    def get_index_of_first_empty_line(self, file_id):
        ranges = [self.list_name + '!' +"B1:B1000"]
        query_results = self.excel_service.spreadsheets().values().batchGet(spreadsheetId=file_id, ranges=ranges).execute()
        sheet_values = query_results['valueRanges'][0]['values']
        return len(sheet_values) + 1

    def check_condition_in_line(self, file_id, row, value):
        ranges = [self.list_name + '!' + row + "1:" + row + "1000"]
        query_results = self.excel_service.spreadsheets().values().batchGet(spreadsheetId=file_id, ranges=ranges).execute()
        sheet_values = query_results['valueRanges'][0]['values']
        for current_value in sheet_values:
            if value in current_value:
                return True
        return False

    def print_in_excel_file(self, file_id, start_line, start_row, values):
        end_row = chr(ord(start_row) + len(values) - 1)
        results = self.excel_service.spreadsheets().values().batchUpdate(spreadsheetId=file_id, body={
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {"range": self.list_name + '!' + start_row + str(start_line) + ':' + end_row + str(start_line),
                     "majorDimension": "ROWS",     # Сначала заполнять строки, затем столбцы
                     "values": [
                                values
                               ]
                    }
                ]
        }).execute()
        return results
