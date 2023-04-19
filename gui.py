import tkinter as tk
import pickle, shelve

from email_sender import sender_controller
from email_reader import read_inbox
from email_reader import outlook_mail_list
from broken_links import br_ln
from tkinter.messagebox import *
from collections import namedtuple
from tkinter import *

def date_convert(date):
    day, month, year = date.split('.')
    months = {
        "01": "Jan",
        "02": "Feb",
        "03": "Mar",
        "04": "Apr",
        "05": "May",
        "06": "Jun",
        "07": "Jul",
        "08": "Aug",
        "09": "Sep",
        "10": "Oct",
        "11": "Nov",
        "12": "Dec"
    }
    return f"{day}-{months[month]}-{year}"

def transf(intgr):
    column = intgr.lower().rstrip()
    a = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s"]
    b = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]
    for i in range(len(a)):
        if column == a[i]:
            return b[i]
            break

class UserForm(tk.Toplevel):
    def __init__(self, parent, user_type):
        super().__init__(parent)
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.gmail = tk.StringVar()

        label = tk.Label(self, text="Настройки доступа к почте\n", font="Verdana 16", foreground="#7e4570",)
        entry_name = tk.Entry(self, textvariable=self.username, width=40)
        entry_pass = tk.Entry(self, textvariable=self.password,width=40)
        entry_gmail = tk.Entry(self, textvariable=self.gmail, width=40)
        btn = tk.Button(self, text="Сохранить", font="Verdana 14", foreground="#fff",background="#7e4570",padx="30", command=self.destroy)

        label.grid(row=0, columnspan=2)
        tk.Label(self, text="Введите свой hotmail адрес:", font="Verdana 12", foreground="#7e4570",).grid(row=1, column=0)
        tk.Label(self, text="Введите пароль приложения:", font="Verdana 12", foreground="#7e4570",).grid(row=2, column=0)
        tk.Label(self, text="Введите ваш gmail адрес", font="Verdana 12", foreground="#7e4570",).grid(row=3, column=0)
        entry_name.grid(row=1, column=1, padx=10)
        entry_pass.grid(row=2, column=1, padx=10)
        entry_gmail.grid(row=3, column=1, padx=10)
        btn.grid(row=4, columnspan=2, pady=10)

    def open(self):
        self.grab_set()
        self.wait_window()
        username = self.username.get()
        password = self.password.get()
        gmail = self.gmail.get()
        datafile=open("works_file/login_data.dat", "wb")
        pickle.dump(username, datafile)
        pickle.dump(password, datafile)
        pickle.dump(gmail, datafile)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.user_type = tk.StringVar()

        m = Menu(self)
        self.config(menu=m)
        m.add_cascade(label="Авторизация", command=self.open_window)

        name_root = Label(self, text="Настройки поиска ссылок\n", font="Verdana 16", foreground="#7e4570",)
        lb_gmail_column = Label(self, text="Столбик почты (буква):", font="Verdana 12", foreground="#7e4570",)
        self.gmail_column = Entry(self, width=30)
        self.gmail_column.insert(0, "G")
        lb_link_column = Label(self, text="Столбик ссылок (буква):", font="Verdana 12", foreground="#7e4570",)
        self.link_column = Entry(self, width=30)
        self.link_column.insert(0, "I")
        lb_start_line = Label(self, text="Начать со строчки (число):", font="Verdana 12", foreground="#7e4570",)
        self.start_line = Entry(self, width=30)
        self.start_line.insert(0, "3")
        lb_nm_of_cells = Label(self, text="Сколько ссылок обработать (число):", font="Verdana 12", foreground="#7e4570",)
        self.nm_of_cells = Entry(self, width=30)
        self.nm_of_cells.insert(0, "1000")
        lb_nm_of_flows = Label(self, text="Сколько потоков запустить (число):", font="Verdana 12", foreground="#7e4570",)
        self.nm_of_flows = Entry(self, width=30)
        self.nm_of_flows.insert(0, "300")
        lb_file_name = Label(self, text="Имя файла бд (имя без .xlsx):", font="Verdana 12", foreground="#7e4570",)
        self.file_name = Entry(self, width=30)
        self.file_name.insert(0, "FP Master Dedupe 2")
        but = Button(self, text='Начать', font="Verdana 14", foreground="#fff",background="#7e4570",padx="30",command=self.ch_bd,)
        name_root.grid(row=0, columnspan = 2,padx=50,)
        lb_gmail_column.grid(row=1, column = 0, padx=15, sticky = 'w')
        self.gmail_column.grid(row=1, column = 1, pady=5, padx=15)
        lb_link_column.grid(row=2, column = 0, padx=15, sticky = 'w')
        self.link_column.grid(row=2, column = 1, pady=5, padx=15)
        lb_start_line.grid(row=3, column = 0, padx=15, sticky = 'w')
        self.start_line.grid(row=3, column = 1, pady=5, padx=15)
        lb_nm_of_cells.grid(row=4, column = 0, padx=15, sticky = 'w')
        self.nm_of_cells.grid(row=4, column = 1, pady=5, padx=15)
        lb_nm_of_flows.grid(row=5, column = 0, padx=15, sticky = 'w')
        self.nm_of_flows.grid(row=5, column = 1, pady=5, padx=15)
        lb_file_name.grid(row=6, column = 0, padx=15, sticky = 'w')
        self.file_name.grid(row=6, column = 1, pady=5, padx=15)
        but.grid(row=8, column = 1, pady=10)

        name_root_1 = Label(self, text="Отправка писем нульовок\n", font="Verdana 16", foreground="#7e4570",)
        lb_delay_time = Label(self, text="Задержка при отправке (секунды):", font="Verdana 12", foreground="#7e4570",)
        self.delay_time = Entry(self, width=30)
        self.delay_time.insert(0, "2")
        but_1 = Button(self, text='Начать', font="Verdana 14", foreground="#fff",background="#7e4570",padx="30",command=self.nl_send,)
        name_root_1.grid(row=9, columnspan = 2,padx=50,)
        lb_delay_time.grid(row=10, column = 0,padx=15, sticky = 'w')
        self.delay_time.grid(row=10, column = 1, pady=5, padx=15, sticky = 'w')
        but_1.grid(row=13, column = 1, pady=10)

        name_root_2 = Label(self, text="Просмотр 100 последних сообщений\n", font="Verdana 16", foreground="#7e4570",)
        inf = Label(self, text="*Без параметров*", font="Verdana 12", foreground="#7e4570",)
        but_2 = Button(self, text='Начать',font="Verdana 14", foreground="#fff",background="#7e4570",padx="30", command=self.index_read,)
        name_root_2.grid(row=14, columnspan = 2,padx=50,)
        inf.grid(row=15, column = 0,padx=15, sticky = 'w')
        but_2.grid(row=15, column = 1, pady=10)

        name_root_3 = Label(self, text="Просмотр сообщений в ящике (по дате)\n", font="Verdana 16", foreground="#7e4570",)
        lb_start_dt = Label(self, text="Начать просмотр с (В формате: д.м.г) :", font="Verdana 12", foreground="#7e4570",)
        self.start_dt = Entry(self, width=30)
        self.start_dt.insert(0, "30.03.2023")
        lb_end_dt = Label(self, text="Закончить просмотр (В формате: д.м.г): ", font="Verdana 12", foreground="#7e4570",)
        self.end_dt = Entry(self, width=30)
        self.end_dt.insert(0, "31.03.2023")
        but_3 = Button(self, text='Начать', font="Verdana 14", foreground="#fff",background="#7e4570",padx="30", command=self.date_read)
        name_root_3.grid(row=16, columnspan = 2,padx=50,)
        lb_start_dt.grid(row=17, column = 0, pady=5, padx=15, sticky = 'w')
        self.start_dt.grid(row=17, column = 1)
        lb_end_dt.grid(row=18, column = 0, pady=5, padx=15, sticky = 'w')
        self.end_dt.grid(row=18, column = 1)
        but_3.grid(row=19, column = 1, pady=10)

    def ch_bd(self):
        self.withdraw()
        br_ln(transf(self.gmail_column.get()), transf(self.link_column.get()), int(self.start_line.get()), int(self.nm_of_cells.get()), int(self.nm_of_flows.get()), self.file_name.get().rstrip())
        showinfo("Обработка базы данных", "Обработка завершина\nРезультаты можно посмотреть в файле new_db.xlsx")
        app = App()
        app.mainloop()

    def nl_send(self):
        sender_controller(float(self.delay_time.get()))

    def date_read(self):
        outlook_mail_list(date_convert(self.start_dt.get()), date_convert(self.end_dt.get()))
        showinfo("Проверка почтового ящика (100 писем)", f"Проверка завершина\nРезультаты можно посмотреть в паке sending_letter или в файле new_db.xlsx")

    def index_read(self):
        read_inbox()
        showinfo("Проверка почтового ящика (по дате)", f"Проверка завершина\nРезультаты можно посмотреть в паке sending_letter или в файле new_db.xlsx")

    def open_window(self):
        window = UserForm(self, self.user_type.get())
        user = window.open()

if __name__ == "__main__":
    app = App()
    app.mainloop()
