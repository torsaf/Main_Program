import os
import ftplib
from datetime import datetime
import shutil


def to_reserve_file():
    """Резервная копия главного файла !Товары"""
    # получение текущей даты и времени
    now = datetime.now()
    date_time = now.strftime("%d-%m-%y - %H-%M")

    # задание деталей FTP-сервера
    server = 'ftp.6449.ru'
    user = 'f0429526'
    password = 'jTtQx0w9x89'
    destination_folder = '/domains/6449.ru/Reserve/'

    # создание копии файла
    original_file = '!Товары.xlsm'
    copy_file = 'copy_' + original_file
    shutil.copy(original_file, copy_file)

    # переименование файла с текущей датой и временем
    new_file = original_file.split('.')[0] + '-' + date_time + '.' + original_file.split('.')[1]
    os.rename(copy_file, new_file)

    # подключение и передача файла на FTP-сервер
    ftp = ftplib.FTP(server)
    ftp.login(user, password)
    with open(new_file, 'rb') as file:
        ftp.cwd(destination_folder)
        ftp.storbinary('STOR ' + new_file, file)

    # удаление переименованного файла
    os.remove(new_file)

    # отключение от FTP-сервера
    ftp.quit()

