import os
import sys
import pathlib
from pathlib import Path
from tkinter import messagebox

def check_prices():
    """Проверяем наличие прайсов. Если их нет или ошибка в наименовании или расширении, то выводим ошибку."""
    directory = Path(pathlib.Path.cwd(), "Prices")
    files_to_check = ['Arispro.xls', 'Artimusic.xls', 'Attrade.xls', 'Gewa.csv', 'Grand.xls', 'Invask.xls', 'Lutner.csv', 'Neva.xls', 'Okno.xls', 'Pop.xls', 'Proaudio.xlsx', 'Roland.xls', 'Slami.xls',
                      'United.xls']

    missing_files = []

    for file_name in files_to_check:
        file_path = os.path.join(directory, file_name)
        if not os.path.exists(file_path):
            missing_files.append(file_name)

    if missing_files:
        messagebox.showerror('Ошибка', f"Проблема с {', '.join(missing_files)}!\nПРОГРАММА НЕ ВЫПОЛНЕНА!")


def check_folder_csv():
    """
    Проверяет наличие папки CSV в корневом каталоге.
    Если папки нет, то создает ее.
    """
    folder_name = "CSV"
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
