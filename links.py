import os
import urllib.parse
import glob
import os.path
import pandas as pd
from pathlib import Path


def create_links(adress):
    def create_link_list(folder_name):
        """Создает список ссылок на фото в папке"""
        link_list = []
        folder_name_encoded = urllib.parse.quote(folder_name)
        n = len(glob.glob1(f"{adress}/{folder_name}", "*.jpg"))  # считает количество фото в папке
        for j in range(1, n + 1):  # выводит ссылки на фотографии от 1 до n включительно
            link = f'<Image url="http://6449.ru/{folder_name_encoded}/{j}.jpg" />'
            link_list.append(link)
        return link_list

    folder_names = [name for name in os.listdir(adress) if os.path.isdir(os.path.join(adress, name))]  # список названий папок

    data = {'Folder': [], 'Content': []}

    for folder_name in folder_names:
        if folder_name == 'desktop.ini':
            continue
        link_list = create_link_list(folder_name)
        data['Folder'].append(folder_name)
        data['Content'].append('\n'.join(link_list))

    df = pd.DataFrame(data)

    desktop_path = str(Path.home() / "Desktop")
    file_path = os.path.join(desktop_path, "Ссылки на фото.xlsx")
    df.to_excel(file_path, index=False)