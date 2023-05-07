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
import links


def handle_paste(event):
    content = event.widget.selection_get(selection="CLIPBOARD")
    event.widget.delete("sel.first", "sel.last")
    event.widget.insert("insert", content)

def button_handler0(event=None):
    adress = frame_1.entry.get()
    try:
        links.create_links(adress)
        messagebox.showinfo("Все ок", "Успешно выполнено!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка : {e}")


customtkinter.set_appearance_mode("light")  # Режимы: system (по умолчанию), light, dark
customtkinter.set_default_color_theme("blue")  # Цветовые темы: blue (по умолчанию), dark-blue, green

app = customtkinter.CTk()  # Создать окно CTk как обычно с использованием окна Tk
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


frame_1.entry = customtkinter.CTkEntry(frame_1, placeholder_text="Вставить: CTRL+V на ENG раскладке")
frame_1.entry.grid(row=3, column=2, columnspan=3, padx=(42, 180), pady=(315, 20), sticky="nsew")
frame_1.main_button_1 = customtkinter.CTkButton(master=frame_1, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), text="Создать файл", command=button_handler0)
frame_1.main_button_1.grid(row=3, column=3, padx=(310, 20), pady=(315, 20), sticky="nsew")
frame_1.entry.bind("<Return>", button_handler0)
frame_1.bind("<Control-v>", handle_paste)

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
button1.place(relx=0.22, rely=0.13, anchor=tkinter.CENTER)
# Описание 1
label = customtkinter.CTkLabel(master=frame_1, text="Обновляем прайсы в папке Prices \nПомним пересохранить United!", font=font1, justify='left')
label.place(relx=0.65, rely=0.13, anchor=tkinter.CENTER)

# Кнопка 2
button2 = customtkinter.CTkButton(master=frame_1, text="Из xls в csv", command=button_handler2, font=font, height=55, width=135)
button2.grid(row=3, column=0, padx=20, pady=10)
button2.place(relx=0.22, rely=0.33, anchor=tkinter.CENTER)
# Описание 2
label2 = customtkinter.CTkLabel(master=frame_1, text="Преобразовываем в CSV формат \nГенерируем сводную таблицу", font=font1, justify='left')
label2.place(relx=0.65, rely=0.33, anchor=tkinter.CENTER)

# Кнопка 3
button3 = customtkinter.CTkButton(master=frame_1, text="Обновить цены", command=lambda: os.startfile('!Товары.xlsm'), font=font, height=55, width=135)
button3.grid(row=3, column=0, padx=20, pady=10)
button3.place(relx=0.22, rely=0.53, anchor=tkinter.CENTER)
# Описание 3
label3 = customtkinter.CTkLabel(master=frame_1, text="     Заходим в файл !Товары \n     Нажимаем кнопку 'Обновить цены'", font=font1, justify='left')
label3.place(relx=0.65, rely=0.53, anchor=tkinter.CENTER)

# Кнопка 4
button3 = customtkinter.CTkButton(master=frame_1, text="Сохранить", command=button_handler4, font=font, height=55, width=135)
button3.grid(row=3, column=0, padx=20, pady=10)
button3.place(relx=0.22, rely=0.73, anchor=tkinter.CENTER)
# Описание 4
label4 = customtkinter.CTkLabel(master=frame_1, text="      Создаём XML файл и делаем бэкап \n      Отправляем на FTP для бота", font=font1, justify='left')
label4.place(relx=0.65, rely=0.73, anchor=tkinter.CENTER)

app.mainloop()
