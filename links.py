import os
import urllib.parse
import glob
import os.path



def create_links(adress):
    def create_file(n):
        """Создает файл и записывает туда название папки и ссылки"""
        with open(f'Ссылки на фото.txt', 'a', encoding='utf-8') as file:
            file.write(n + '\n')

    # adress = input('Введите ссылку на папку: ')  # указываем путь к папкам
    folder_names = os.listdir(adress)  # создаем список названий папок

    # Удаляем старый файл, если он есть
    if os.path.exists("Ссылки на фото.txt"):
        os.remove("Ссылки на фото.txt")

    for folder_name in folder_names:  # проходим циклом по каждой папке
        create_file(str(folder_name))  # вызов функции создания файла
        n = len(glob.glob1(f"{adress}/{folder_name}", "*.jpg"))  # считает количество фото в папке
        folder_name_encoded = urllib.parse.quote(folder_name)
        for j in range(1, n + 1):  # выводит ссылки на фотографии от 1 до n включительно
            link = f'<Image url="http://6449.ru/{folder_name_encoded}/{j}.jpg" />'
            create_file(str(link))
        create_file('')  # Создает пробел между моделями
