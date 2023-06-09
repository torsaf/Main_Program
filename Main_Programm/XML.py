import pandas as pd
import pathlib
import openpyxl
import sys


def to_xml_file():
    """Создаёт XML файл"""
    # Получаем текущую директорию
    dir_path = pathlib.Path.cwd()

    # Составляем путь до файла
    file_path = dir_path / "!Товары.xlsm"

    # Проверяем наличие файла и выводим сообщение
    if file_path.is_file() == False:
        print("Файл '!Товары.xlsm' не найден")
        input()
        sys.exit([0])

    # Указываем пути к файлам
    path = pathlib.Path.cwd() / "!Товары.xlsm"
    path_xml = pathlib.Path.cwd() / "XML.xml"

    # Читаем XLSX-файл и выбираем нужные столбцы
    df = pd.read_excel(path, header=None, sheet_name='General').iloc[1:, [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 30]]
    # Переименуем столбцы
    df.columns = ['Id', 'Category', 'GoodsType', 'AdType', 'Address', 'AllowEmail',
                  'ContactPhone', 'Condition', 'ManagerName', 'Title', 'Description', 'Images', 'VideoURL', 'Price']
    # Удаляем строки, где цена равна 0
    df = df[df['Price'] != 0]
    # Заменяем цены, равные 1, на пропущенные значения
    df.loc[df['Price'] == 1, 'Price'] = pd.NA
    # Заполняем пропущенные значения в столбцах цены и ссылок на видео
    df[['Price', 'VideoURL']] = df[['Price', 'VideoURL']].fillna('')
    # Заменяем символ '&' на строку '&amp;' в столбцах Title, Description, Images и VideoURL
    df[['Title', 'Description', 'Images', 'VideoURL']] = df[['Title', 'Description', 'Images', 'VideoURL']].replace('&', '&amp;', regex=True)

    # Функция, которая преобразует строки DataFrame в XML-строки
    def to_xml(df, filename=None, mode='w'):
        # Вложенная функция, которая преобразует одну строку DataFrame в одну XML-строку
        def row_to_xml(row):
            xml = ['<Ad>']
            for i, col_name in enumerate(row.index):
                xml.append(f'  <{col_name}>{row.iloc[i]}</{col_name}>')
            xml.append('</Ad>')
            return '\n'.join(xml)

        # Преобразуем все строки DataFrame в XML-строки
        res = '\n'.join(df.apply(row_to_xml, axis=1))

        # Если не указано имя файла, то возвращаем XML-строки
        if filename is None:
            return res
        # Иначе записываем полученные данные в указанный файл
        with open(filename, mode, encoding="utf-8") as f:
            f.write(res)

    # Добавляем метод to_xml в модуль DataFrame
    pd.DataFrame.to_xml = to_xml

    # Строки, которые будут записаны в начале и конце XML-файла
    a = '<?xml version="1.0" encoding="UTF-8"?>\n<Ads formatVersion="3" target="Avito.ru">\n'
    b = '\n</Ads>'

    # Открываем файл и записываем в него собранные данные в XML-формате
    with open(path_xml, "w", encoding="utf-8") as h:
        h.write(''.join([a, df.to_xml(), b]))
