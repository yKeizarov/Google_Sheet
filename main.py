from dataSet import JSON_DATA
from g_drive_sheet import GoogleSheet


new_doc = GoogleSheet(googleApiDataSet=JSON_DATA, table_name="Моя новая таблица тест", sheet_name="первый лист")
print(new_doc.get_sheet())
new_doc2 = GoogleSheet(googleApiDataSet=JSON_DATA, table_name="Еще один документ", sheet_name="превый листочек")
print(new_doc2.get_sheet())