from ftplib import FTP


def to_ftp():
    """Закидывает файлы на FTP для Ботов"""
    # Настройки FTP сервера
    ftp_server_address = 'ftp.6449.ru'
    ftp_directory = '/domains/f0429526.xsph.ru/my_bot/'
    ftp_forclients_directory = '/domains/f0429526.xsph.ru/forclients/'
    ftp_username = 'f0429526'
    ftp_password = 'jTtQx0w9x89'

    # Список файлов, которые нужно отправить
    file_names = ['!BD.csv', '!Forclients.csv', '!Name.csv']

    # Создаем соединение с FTP сервером
    ftp = FTP(ftp_server_address)
    ftp.login(user=ftp_username, passwd=ftp_password)

    # Переходим в нужную директорию на сервере
    ftp.cwd(ftp_directory)

    # Отправляем каждый файл в список
    for file_name in file_names:
        file_path = 'CSV/' + file_name  # путь к файлу
        with open(file_path, 'rb') as file:
            ftp.storbinary(f'STOR {file_name}', file)  # отправляем файл в директорию my_bot на сервере
        if file_name == '!Forclients.csv':
            ftp.cwd(ftp_forclients_directory)
            # Отправляем файл !Forclients.csv также в директорию forclients
            with open(file_path, 'rb') as file:
                ftp.storbinary(f'STOR {file_name}', file)
        # Возвращаемся в директорию my_bot до отправки следующего файла
        ftp.cwd(ftp_directory)
    # Закрываем соединение
    ftp.quit()
