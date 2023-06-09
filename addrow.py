import pandas as pd
import openpyxl
from pandas.io.excel import ExcelWriter
# data = pd.read_excel(io='data.xlsx', engine='openpyxl')
#
#
# df1 = pd.DataFrame([['a', 'b'], ['c', 'd']],
#                    index=['row 1', 'row 2'],
#                    columns=['col 1', 'col 2'])
# df1.to_excel("output.xlsx")
#
my_data = pd.DataFrame([["Кейзеров Евгений Сергеевич", "10 сентября 1990", "Начальник СГ и СМ", 0, 0],
                        ["Мазура Олег Геннадьевич", "24 августа 1973", "Начальник склада (авиа)", 1, 3],
                        ["Поскура Юрий Александрович", "13 ноябр 1980", "Начальник склада (авто)", 1, 2]],
                       columns=["ФИО", "Дата рождения", "Должность", "Холост/Женат", "Количество детей"])
my_data.to_excel("test.xlsx")

data = pd.read_excel('test.xlsx', engine='openpyxl')
print(data)

new_data = pd.DataFrame({"ФИО": "Плавский Сергей Викторович", "Дата рождения": "01 января 2000",
                         "Должность": "Начальник ЦЗТ", "Холост/Женат": 1, "Количество детей": 2})


with ExcelWriter('test.xlsx', mode="a") as writer:
    new_data.sample(10).to_excel(writer, sheet_name="Лист 3")

# data.append({"ФИО": "Плавский Сергей Викторович", "Дата рождения": "01 января 2000", "Должность": "Начальник ЦЗТ",
#             "Холост/Женат": 1, "Количество детей": 2}, ignore_index=True)

