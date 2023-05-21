import pandas as pd
import numpy as np
import pathlib
from pathlib import Path


def generate_new_prices():
    class Vendors:
        def __init__(self, name, short_name):
            self.name = name
            self.short_name = short_name

        def create_data_frame():
            """Создаём фрейм из файла General"""
            df = pd.read_csv(Path(pathlib.Path.cwd(), "CSV", "General.csv"), sep=';')
            return df

        def create_vendors(self, df):
            """Создаём переменную поставщика и оставляем только нужные 4 столбца"""
            self.short_name = pd.read_csv(Path(pathlib.Path.cwd(), "CSV", f"{self.name}.csv"), sep=';')[['Артикул', 'Наличие', 'ОПТ', 'РРЦ']].astype(str)
            # Берем товары, которые в наличии и пихаем в новый столбец
            df = pd.merge(df, self.short_name, left_on=self.name, right_on='Артикул', how='left')
            # Переименовываем столбцы
            new_columns = {'ОПТ': f'ОПТ {self.name}', 'РРЦ': f'РРЦ {self.name}', 'Артикул': f'Арт {self.name}', 'Наличие': f'Нал {self.name}'}
            df.rename(columns=new_columns, inplace=True)
            # Проверяем, если в столбце стоит значение 0 или оно пустое, то значит товара нет и в столбце 'ОПТ' оставляем пустое значение
            na_values = [' ', '0', None, 'нет', 'витрина', 'резерв']
            df.loc[df[f'Нал {self.name}'].isin(na_values), f'ОПТ {self.name}'] = pd.NA
            # для 'РРЦ {self.name}'
            df.loc[df[f'Нал {self.name}'].isin(na_values), f'РРЦ {self.name}'] = pd.NA
            return df

        def generate_price(df):
            """Процентовка, наценка, округление"""
            df["Итоговая цена"] = df[["ОПТ Attrade", "ОПТ Slami", "ОПТ Invask", "ОПТ Proaudio", "ОПТ Arispro", "ОПТ Artimusic", "ОПТ Pop", "ОПТ Roland", "ОПТ Okno", "ОПТ Grand", "ОПТ Lutner", "ОПТ Neva", "ОПТ Gewa", "ОПТ United"]].apply(pd.to_numeric, errors='coerce').astype('Int64').min(axis=1, numeric_only=True).fillna(0)
            df["Min_РРЦ"] = df[["РРЦ Attrade", "РРЦ Slami", "РРЦ Invask", "РРЦ Proaudio", "РРЦ Arispro", "РРЦ Artimusic", "РРЦ Pop", "РРЦ Roland", "РРЦ Okno", "РРЦ Grand", "РРЦ Lutner", "РРЦ Neva", "РРЦ Gewa", "РРЦ United"]].apply(pd.to_numeric, errors='coerce').astype('Int64').min(axis=1, numeric_only=True).fillna(0)
            edges = [0, 10000, 20000, 50000, 100000, np.inf]  # интервалы от и до. Cправа бесконечность
            pcts = iter([22, 15, 15, 15, 11])  # проценты. количество должно совпадать с количеством интервалов
            df['Итоговая цена'] = df.groupby(pd.cut(df["Итоговая цена"], bins=edges, right=False))["Итоговая цена"].apply(lambda x: x + x * next(pcts) / 100)  # преобразование цены по диапазонам
            df['Итоговая цена'] = df["Итоговая цена"].apply(lambda x: np.round(x * 2, -3) // 2)  # округление
            return df

        def forclients(df):
            """Генерируем файл forclients"""
            tofile = df.copy()
            tofile.loc[tofile['Склад'].isin([1, '1']) & (~tofile['Цена склада'].isna()), 'Итоговая цена'] = tofile['Цена склада']
            tofile = tofile.loc[tofile['Итоговая цена'].fillna(0).astype(int) != 0, ['Название объявления - Title', 'Склад', 'Цена склада', 'Итоговая цена', 'Min_РРЦ']]
            # tofile['РРЦ'] = tofile['РРЦ'].fillna(0).astype(int)
            tofile.loc[tofile['Склад'].isin([3, '3']) & (tofile['Итоговая цена'] != 0), 'Итоговая цена'] = tofile['Min_РРЦ']
            tofile['Итоговая цена'] = tofile['Итоговая цена'].astype(int)
            tofile = tofile[['Название объявления - Title', 'Итоговая цена']]
            with pathlib.Path("CSV", "!Forclients.csv").open("w", newline="", encoding="utf-8") as f:
                tofile.to_csv(f, sep=";", index=False)

        def one_zero(df):
            """Округляем столбец Min_РРЦ и Вносим изменения в столбец Итоговая цена"""
            df['Min_РРЦ'] = df['Min_РРЦ'].replace([np.inf, -np.inf], np.nan).fillna(0).astype(float).apply(lambda x: int(x // 100 * 100))
            # Просле процентовки необходимо поставить заменить цены на 1, там где нельзя ставить цены и проставить 1 и его цену, если товар есть на нашем складе.
            df.loc[df['Склад'].isin([1, '1']) & (df['Цена склада'].isna()), 'Итоговая цена'] = 1
            df.loc[df['Склад'].isin([1, '1']) & (~df['Цена склада'].isna()), 'Итоговая цена'] = df['Цена склада']
            df.loc[df['Склад'].isin([2, '2']) & (df['Итоговая цена'] != 0), 'Итоговая цена'] = 1
            df.loc[df['Склад'].isin([3, '3']) & (df['Итоговая цена'] != 0), 'Итоговая цена'] = df['Min_РРЦ']
            # Для Слами Roland
            # df.loc[df['Склад'].isin([3, '3']) & (df['ОПТ Slami'] == 'Уточняйте'), 'Итоговая цена'] = df['Min_РРЦ']
            # Удаление последнего столбца РРЦ, из которго мы берем РРЦ цену, Если в поле СКЛАД стоит значение 3.
            df.drop(columns=['Min_РРЦ'], inplace=True)
            return df

        def printtofile(df):
            """Сохраняем результат в Сводная таблица"""
            writer = pd.ExcelWriter(Path(pathlib.Path.cwd(), "CSV", "Сводная таблица.xlsx"), engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            writer.save()

    Attrade = Vendors('Attrade', 'att')
    Slami = Vendors('Slami', 'slm')
    Invask = Vendors('Invask', 'inv')
    Proaudio = Vendors('Proaudio', 'pro')
    Arispro = Vendors('Arispro', 'ari')
    Artimusic = Vendors('Artimusic', 'art')
    Pop = Vendors('Pop', 'pop')
    Roland = Vendors('Roland', 'rol')
    Okno = Vendors('Okno', 'okn')
    Grand = Vendors('Grand', 'grn')
    Lutner = Vendors('Lutner', 'lut')
    Neva = Vendors('Neva', 'nev')
    Gewa = Vendors('Gewa', 'gew')
    United = Vendors('United', 'uni')

    df = Vendors.create_data_frame()
    df = Attrade.create_vendors(df)
    df = Slami.create_vendors(df)
    df = Invask.create_vendors(df)
    df = Proaudio.create_vendors(df)
    df = Arispro.create_vendors(df)
    df = Artimusic.create_vendors(df)
    df = Pop.create_vendors(df)
    df = Roland.create_vendors(df)
    df = Okno.create_vendors(df)
    df = Grand.create_vendors(df)
    df = Lutner.create_vendors(df)
    df = Neva.create_vendors(df)
    df = Gewa.create_vendors(df)
    df = United.create_vendors(df)

    Vendors.generate_price(df)
    Vendors.forclients(df)
    Vendors.one_zero(df)
    Vendors.printtofile(df)
