import tkinter as tk
from tkinter import messagebox
import pathlib
import os
import check
import step1
import step2
import push
import XML
import FTP
import reserve

# Создаем главное окно
window = tk.Tk()
window.title("Program for Avito")

# Рассчитываем позиции X и Y для расположения окна в центре экрана
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = (screen_width // 2) - (500 // 2)  # 300 это ширина окна
y = (screen_height // 2) - (300 // 2)  # 400 это высота окна

# Устанавливаем размеры окна и его позицию в центре экрана
window.geometry("550x400+{}+{}".format(x, y))

# запретим пользователю изменять размеры окна
window.resizable(False, False)


# Создаем функцию-обработчик кнопки
def button_handler1():
    try:
        push.download_prices()
        messagebox.showinfo("Все ок", "Успешно выполнено!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка : {e}")

    os.startfile(f'{pathlib.Path.cwd()}/Prices/United.xls')

def button_handler2():
    try:
        check.check_prices()
        check.check_folder_csv()
        step1.all_to_csv()
        step2.generate_new_prices()
        messagebox.showinfo("Все ок", "Успешно выполнено!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка : {e}")


def button_handler4():
    try:
        XML.to_xml_file()
        FTP.to_ftp()
        reserve.to_reserve_file()
        messagebox.showinfo("Все ок", "Успешно выполнено!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка : {e}")


# Создаем LabelFrame и добавляем в них каждую группу кнопок и соответствующие подписи
frame1 = tk.LabelFrame(window, text="1. Скачать прайсы", font=('Segoe UI', 13, 'bold'))
button1 = tk.Button(frame1, text='Выполнить', command=button_handler1)
label1 = tk.Label(frame1, text="Обновляем прайсы в папке Prices. \nПомним пересохранить United!", borderwidth=3, highlightthickness=6)
label1.config(width=40)

frame2 = tk.LabelFrame(window, text="2. Из xls в csv", font=('Segoe UI', 13, 'bold'))
button2 = tk.Button(frame2, text='Выполнить', command=button_handler2)
label2 = tk.Label(frame2, text="Преобразовываем всё в CSV формат \n и генерируем сводную таблицу", borderwidth=3, highlightthickness=6)
label2.config(width=40)

frame3 = tk.LabelFrame(window, text="3. Обновить цены", font=('Segoe UI', 13, 'bold'))
button3 = tk.Button(frame3, text='   Открыть   ', command=lambda: os.startfile('!Товары.xlsm'))
label3 = tk.Label(frame3, text="Необходимо зайти в файл !Товары и \n нажать кнопку 'Обновить цены'", borderwidth=3, highlightthickness=6)
label3.config(width=40)

frame4 = tk.LabelFrame(window, text="4. Сохранить", font=('Segoe UI', 13, 'bold'))
button4 = tk.Button(frame4, text='Выполнить', command=button_handler4)
label4 = tk.Label(frame4, text="Создаём XML файл и делаем бэкап файла \n!Товары и отправляем на FTP для бота", borderwidth=3, highlightthickness=6)
label4.config(width=40)

# Размещение кнопок и подписей в каждом фрейме
button1.pack(side=tk.LEFT, padx=10)
label1.pack(side=tk.LEFT, padx=10)

button2.pack(side=tk.LEFT, padx=10)
label2.pack(side=tk.LEFT, padx=10)

button3.pack(side=tk.LEFT, padx=10)
label3.pack(side=tk.LEFT, padx=10)

button4.pack(side=tk.LEFT, padx=10)
label4.pack(side=tk.LEFT, padx=10)

frame1.place(relx=0.1, rely=0.05, anchor=tk.NW)
frame2.place(relx=0.1, rely=0.28, anchor=tk.NW)
frame3.place(relx=0.1, rely=0.51, anchor=tk.NW)
frame4.place(relx=0.1, rely=0.74, anchor=tk.NW)

# Запускаем главный цикл обработки событий
window.mainloop()
