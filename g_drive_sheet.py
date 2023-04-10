import time
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from dataSet import JSON_DATA, GOOGLE_SHEET_ADDR, GOOGLE_DRIVE_ADDR


class GoogleSheet:
    def __init__(self, googleApiDataSet: str, table_name: str, sheet_name: str):
        print("Start connecting...")
        self.googleApiDataSet = googleApiDataSet
        self.readGoogleKeys = ServiceAccountCredentials.from_json_keyfile_name(
            self.googleApiDataSet,
            [GOOGLE_SHEET_ADDR, GOOGLE_DRIVE_ADDR]
        )
        time.sleep(0.5)
        print("Authorization process...")
        self.authorization = self.readGoogleKeys.authorize(httplib2.Http())
        time.sleep(0.5)
        print("Choices work with table...")
        self.workService = apiclient.discovery.build('sheets', 'v4', http=self.authorization)
        time.sleep(0.5)
        # Дальше идет создание документа, тут стоп, необходимо это упаковать в метод
        print('Creating new table...')
        self.table_name = table_name
        self.sheet_name = sheet_name

        spreadsheet = self.workService.spreadsheets().create(
            body={
                'properties': {'title': self.table_name,
                               'locale': 'ru_RU'},
                'sheets': [{'properties': {
                    'sheetType': 'GRID',
                    'sheetId': 0,
                    'title': self.sheet_name,
                    'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
            }).execute()
        self.new_document = spreadsheet['spreadsheetId']
        time.sleep(0.5)
        print("Create your file")

        drive_service = apiclient.discovery.build('drive', 'v3',
                                                 http=self.authorization)
        access = drive_service.permissions().create(
            fileId=self.new_document,
            body={'type': 'user', 'role': 'writer', 'emailAddress': 'y.keizarov@gmail.com'},
            fields='id'
        ).execute()
        time.sleep(0.5)
        print(f'Work table: "{table_name}" is created')

    def write_excel_sheet(self, title: str, row_count: int, column_count: int):
        """Create new sheet in excel document"""
        results = self.workService.spreadsheets().batchUpdate(
            spreadsheetId=self.new_document,
            body=
            {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": title,
                                "gridProperties": {
                                    "rowCount": row_count,
                                    "columnCount": column_count
                                }
                            }
                        }
                    }
                ]
            }).execute()
        
    def get_sheet(self):
        """Get all sheets in excel document"""
        exel_document = self.workService.spreadsheets().get(spreadsheetId=self.new_document).execute()
        sheet_list = exel_document.get("sheets")
        all_lists = []
        for sheet in sheet_list:
            all_lists.append((f"Sheet ID: {sheet['properties']['sheetId']}", f"Sheet name: {sheet['properties']['title']}"))
        return all_lists

