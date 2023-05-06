import tkinter
import customtkinter
from PIL import Image
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

customtkinter.set_appearance_mode("light")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

app = customtkinter.CTk()  # create CTk window like you do with the Tk window
app.geometry("600x400")
app.resizable(False, False)

frame_1 = customtkinter.CTkFrame(master=app)
frame_1.pack(pady=20, padx=50, fill="both", expand=True)

app.title("Program for Avito")
# Определите размеры экрана
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()

# Определите позицию окна на экране
x = (screen_width - app.winfo_reqwidth()) / 2
y = (screen_height - app.winfo_reqheight()) / 2

# Установите позицию окна приложения в середине экрана
app.geometry("+%d+%d" % (x, y))


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


font = ("Helvetica", 14, "bold")
font1 = ("Calibri", 14)


# Кнопка 1
button1 = customtkinter.CTkButton(master=frame_1, text="Скачать прайсы", command=button_handler1, font=font, height=55, width=135)
button1.grid(row=3, column=0, padx=20, pady=10)
button1.place(relx=0.22, rely=0.15, anchor=tkinter.CENTER)
# Описание 1
label = customtkinter.CTkLabel(master=frame_1, text="Обновляем прайсы в папке Prices \nПомним пересохранить United!", font=font1, justify='left')
label.place(relx=0.65, rely=0.15, anchor=tkinter.CENTER)

# Кнопка 2
button2 = customtkinter.CTkButton(master=frame_1, text="Из xls в csv", command=button_handler2, font=font, height=55, width=135)
button2.grid(row=3, column=0, padx=20, pady=10)
button2.place(relx=0.22, rely=0.38, anchor=tkinter.CENTER)
# Описание 2
label2 = customtkinter.CTkLabel(master=frame_1, text="Преобразовываем в CSV формат \nГенерируем сводную таблицу", font=font1, justify='left')
label2.place(relx=0.65, rely=0.38, anchor=tkinter.CENTER)

# Кнопка 3
button3 = customtkinter.CTkButton(master=frame_1, text="Обновить цены", command=lambda: os.startfile('!Товары.xlsm'), font=font, height=55, width=135)
button3.grid(row=3, column=0, padx=20, pady=10)
button3.place(relx=0.22, rely=0.61, anchor=tkinter.CENTER)
# Описание 3
label3 = customtkinter.CTkLabel(master=frame_1, text="     Заходим в файл !Товары \n     Нажимаем кнопку 'Обновить цены'", font=font1, justify='left')
label3.place(relx=0.65, rely=0.61, anchor=tkinter.CENTER)

# Кнопка 4
button3 = customtkinter.CTkButton(master=frame_1, text="Сохранить", command=button_handler4, font=font, height=55, width=135)
button3.grid(row=3, column=0, padx=20, pady=10)
button3.place(relx=0.22, rely=0.84, anchor=tkinter.CENTER)
# Описание 4
label4 = customtkinter.CTkLabel(master=frame_1, text="      Создаём XML файл и делаем бэкап \n      Отправляем на FTP для бота", font=font1, justify='left')
label4.place(relx=0.65, rely=0.84, anchor=tkinter.CENTER)

app.mainloop()
