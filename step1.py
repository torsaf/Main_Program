import pandas as pd
import pathlib
from pathlib import Path
import xlrd
import os


def all_to_csv():
    class Vendor:
        def __init__(self, name, shortname, start, art, model, nal, opt, rrc, ext='xls', sheet_name=0):
            self.name = name
            self.shortname = shortname
            self.start = start
            self.art = art
            self.model = model
            self.nal = nal
            self.opt = opt
            self.rrc = rrc
            self.ext = ext
            self.sheet_name = sheet_name

        def set_name(self, column, from_opt=0, margin=0):
            self.column = column
            self.from_opt = from_opt
            self.margin = margin

        def delit_CSV():
            """Удаляем старые файлы, что бы заполнить их заново"""
            if os.path.exists(Path(pathlib.Path.cwd(), "CSV", "!BD.csv")):
                os.remove(Path(pathlib.Path.cwd(), "CSV", "!BD.csv"))
            if os.path.exists(Path(pathlib.Path.cwd(), "CSV", "!Name.csv")):
                os.remove(Path(pathlib.Path.cwd(), "CSV", "!Name.csv"))

        def open_csv_file(self):
            """открывает CSV-файл"""
            df = pd.read_csv(Path(pathlib.Path.cwd(), "Prices", f"{self.name}.{self.ext}"), encoding='windows-1251', dtype=object, sep=';', header=None).loc[self.start:, :]
            self.re_name(df)

        def open_xls_file(self):
            """открывает XLS-файл"""
            if self.name == 'Invask':
                file_path = f"./Prices/{self.name}.{self.ext}"
                # откройте файл с помощью xlrd и прочитайте файл с помощью Pandas
                df = pd.read_excel(xlrd.open_workbook(file_path, encoding_override="windows-1251"), header=None).fillna(0).loc[self.start:, :]
            else:
                df = pd.read_excel(Path(pathlib.Path.cwd(), "Prices", f"{self.name}.{self.ext}"), sheet_name=self.sheet_name, header=None).loc[self.start:, :]
                df = df.fillna(0)
            # если название прайсов совпадает с этими тремя, то переходим в функцию, которая создает столбец с Оптовой ценой
            if self.name in ('Arispro', 'Proaudio', 'Roland'):
                self.create_column(df)
            else:
                self.re_name(df)

        def create_column(self, df):
            """Создаёт столбец ОПТ для 'Arispro', 'Proaudio', 'Roland'"""
            df[self.column] = (df[self.from_opt] * self.margin).astype(int)
            self.re_name(df)

        def re_name(self, df):
            """Переименовывает столбцы в нужный вид"""
            df.rename(columns={self.art: 'Артикул', self.model: 'Модель', self.nal: 'Наличие', self.opt: 'ОПТ', self.rrc: 'РРЦ'}, inplace=True)
            df['Поставщик'] = self.name
            df = df[['Поставщик', 'Артикул', 'Модель', 'Наличие', 'ОПТ', 'РРЦ']]
            df = df.drop_duplicates('Артикул')  # убирает дубликаты артикулов, т.к. такое встречается в прайсах.
            self.change_words(df)

        def change_words(self, df):
            """Убираем строки в ячейках которых, всякий мусор"""
            # Определяем списки значений, которые требуется исключить
            exclude_values = ['Уточняйте ', 'Уточняйте', 'не для продажи в РФ', 'Çâîíèòå', '0', 'Звоните', 'витрина']
            # Фильтруем DataFrame, используя значение включающего оператора "и" (&)
            df = df.loc[~df['РРЦ'].isin(exclude_values) & ~df['ОПТ'].isin(exclude_values) & (df['РРЦ'] != 0) & (df['ОПТ'] != 0)]
            df = df.loc[(df['Наличие'] != 'витрина') & (df['ОПТ'] != 'витрина')]
            # Удаляем дубликаты
            df = df.drop_duplicates('Артикул')
            self.change_type(df)

        def change_type(self, df):
            """Выравниваем типы данных во всех столбцах"""
            df.fillna(0, inplace=True)
            df[['ОПТ', 'РРЦ']] = df[['ОПТ', 'РРЦ']].astype(float).astype(int)
            # убираем в двух столбцах пробелы с двух сторон
            df['Артикул'] = df['Артикул'].apply(lambda x: str(x).strip())
            df['Модель'] = df['Модель'].apply(lambda x: str(x).strip())
            self.tocsv(df)
            self.create_BD_file(df)

        def tocsv(self, df):
            """Сохраняет в CSV Файл всех поставщиков после того, как мы прибрались в столбцах"""
            df.to_csv(Path(pathlib.Path.cwd(), "CSV", f"{self.name}.csv"), sep=';', mode='w', index=False)

        def create_BD_file(self, df):
            """Создаёт файл !BD"""
            df['Поставщик'] = self.shortname
            df.to_csv(Path(pathlib.Path.cwd(), "CSV", "!BD.csv"), sep=';', mode='a', index=False)

        def create_Name_file():
            """Берёт из главного файла название и все артикулы и сохраняет в файл !Name"""
            df1 = pd.read_excel(Path(pathlib.Path.cwd(), "!Товары.xlsm"), sheet_name='General', header=None).iloc[2:, 10:31]
            df1 = df1[[10, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29]]
            df1 = df1.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            df1.to_csv(Path(pathlib.Path.cwd(), "CSV", "!Name.csv"), index=False)

        def create_general_file():
            """Из главного файла берет большинство столбцов и сохраняет в файл General, для второй части кода"""
            df = pd.read_excel(r"!Товары.xlsm", usecols=[0, 10, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30], header=None, skiprows=1)  # Считывает содержимое файла Excel "!Товары.xlsm" в объект pandas.DataFrame, выбирая только определенные столбцы и пропуская первую строку с заголовками столбцов в качестве шапки
            df = df.applymap(str)  # Преобразует все значения в DataFrame в строковый формат
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)  # Убираем пробелы справа и слева во всех столбцах DataFrame
            df.rename(columns={0: 'Номер в Avito - Id', 10: 'Название объявления - Title', 14: 'Склад', 15: 'Цена склада', 16: 'Attrade', 17: 'Slami', 18: 'Invask', 19: 'Proaudio', 20: 'Arispro',
                               21: 'Artimusic', 22: 'Pop', 23: 'Roland', 24: 'Okno', 25: 'Grand', 26: 'Lutner', 27: 'Neva', 28: 'Gewa', 29: 'United', 30: 'Итоговая цена'}, inplace=True)
            # Переименовывает выбранные столбцы DataFrame по определенным меткам
            df['Итоговая цена'] = 0  # Создает новый столбец в DataFrame с именем Итоговая цена и заполняет его нулями
            df.to_csv(r"CSV\General.csv", sep=';', mode='w', index=False)  # Сохраняет объект DataFrame в формате CSV в файл CSV\General.csv с разделителем ';' и без индекса строк


    Artimusic = Vendor('Artimusic', 'ART', 4, 1, 0, 5, 2, 4)
    Attrade = Vendor('Attrade', 'ATT', 21, 1, 4, 7, 18, 11)
    Gewa = Vendor('Gewa', 'GEW', 16, 0, 2, 3, 5, 6, 'csv')
    Grand = Vendor('Grand', 'GDM', 10, 1, 2, 3, 7, 9)
    Invask = Vendor('Invask', 'INV', 14, 0, 2, 8, 7, 6)
    Lutner = Vendor('Lutner', 'LTN', 1, 1, 2, 3, 4, 6, 'csv')
    Neva = Vendor('Neva', 'NVS', 9, 0, 4, 16, 15, 14)
    Okno = Vendor('Okno', 'OKN', 8, 0, 1, 5, 3, 4, 'xls', 'Основной прайс-лист')
    Pop = Vendor('Pop', 'POP', 8, 13, 0, 5, 9, 10)
    Slami = Vendor('Slami', 'SLM', 5, 2, 4, 8, 7, 5)
    United = Vendor('United', 'UNT', 9, 1, 2, 5, 11, 9)
    Arispro = Vendor('Arispro', 'ARP', 4, 2, 3, 5, 7, 4)
    Arispro.set_name(column=7, from_opt=4, margin=0.7)
    Proaudio = Vendor('Proaudio', 'PRO', 7, 0, 2, 4, 8, 5, 'xlsx', 'Прайс-Лист')
    Proaudio.set_name(column=8, from_opt=5, margin=0.75)
    Roland = Vendor('Roland', 'ROL', 14, 2, 1, 9, 7, 5)
    Roland.set_name(column=9)

    Vendor.delit_CSV()
    Artimusic.open_xls_file()
    Attrade.open_xls_file()
    Gewa.open_csv_file()
    Grand.open_xls_file()
    Invask.open_xls_file()
    Lutner.open_csv_file()
    Neva.open_xls_file()
    Okno.open_xls_file()
    Pop.open_xls_file()
    Slami.open_xls_file()
    United.open_xls_file()
    Arispro.open_xls_file()
    Proaudio.open_xls_file()
    Roland.open_xls_file()
    Vendor.create_Name_file()
    Vendor.create_general_file()
