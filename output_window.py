import openpyxl
from tkinter import *
from tkinter import ttk
import messages as mes
import get_requests_window
import variables as var
import service
import input_window


def out_xlsx(win):
    # очистка фрейма, содержащего таблицу для её обновления
    def clear_frame():
        for widget in fr4.winfo_children():
            widget.destroy()

    def open_custom_fill():
        input_window.open_custom_fill(sh)
        sh.withdraw()

    def get_request_from_mail():
        get_requests_window.requests_win(sh)
        sh.withdraw()

    def sh_closing():
        sh.destroy()
        win.deiconify()

    def get_nc():
        print(var.path_table, 111)
        if var.path_table == '':
            mes.error("Таблица с данными", "Выберите таблицу для обработки!")
        else:
            wb = openpyxl.load_workbook(var.path_table)
            sheet = wb["Заявление"]

            start = 'B2'
            end = sheet['P'][-1].coordinate
            lst = []
            test = []
            rows = 0

            try:
                for row in sheet[f'{start}':f'{end}']:
                    rows += 1
                    for cellObj in row:
                        if cellObj.value == None or cellObj.value == '':
                            continue

                        test.append(str(cellObj.value))
                x = 0
                n = 1
                while x < len(test):
                    lst.append((str(n), test[x + 14], test[x + 7]))
                    n += 1
                    x += 15
                wb.close()
                return lst
            except:
                mes.error("Обработка входных данных", "Ошибка обработки данных!")
                wb.close()

    def clear_eout():
        eout_seria.delete(0, END)
        eout_numb.delete(0, END)
        eout_date.delete(0, END)
        eout_code.delete(0, END)
        eout_date_birth.delete(0, END)
        eout_place_of_birth.delete(0, END)
        eout_gender.delete(0, END)
        eout_fam.delete(0, END)
        eout_name.delete(0, END)
        eout_otch.delete(0, END)
        eout_inn.delete(0, END)
        eout_snils.delete(0, END)
        eout_job.delete(0, END)
        eout_email.delete(0, END)
        eout_city.delete(0, END)

    def show():
        def item_selected(event):
            for selected_item in table.selection():
                item = table.item(selected_item)
                record = item['values']
                show_data(record[2])

        def insert_data(data):
            clear_eout()  # очищаем поля

            eout_seria.insert(0, data[0])
            eout_numb.insert(0, data[1])
            eout_date.insert(0, data[2])
            eout_code.insert(0, data[3])
            eout_date_birth.insert(0, data[4])
            eout_place_of_birth.insert(0, data[5])
            eout_gender.insert(0, data[6])
            eout_fam.insert(0, data[7])
            eout_name.insert(0, data[8])
            eout_otch.insert(0, data[9])
            eout_inn.insert(0, data[10])
            eout_snils.insert(0, data[11])
            eout_job.insert(0, data[12])
            eout_email.insert(0, data[13])
            eout_city.insert(0, data[14])

        def show_data(name):
            print(var.path_table, 222)
            if var.path_table == '':
                mes.error("Таблица с данными", "Выберите таблицу для обработки!")
            else:
                wb = openpyxl.load_workbook(var.path_table)
                sheet = wb["Заявление"]

                start = 'B2'
                end = sheet['P'][-1].coordinate
                data = []
                try:
                    for row in sheet[f'{start}':f'{end}']:
                        if (row[7].value) == name:
                            for cellObj in row:
                                data.append(cellObj.value)
                        else:
                            continue
                    wb.close()
                    insert_data(data)
                except:
                    mes.error("Обработка входных данных", "Ошибка обработки данных!")
                    wb.close()
        print(var.path_table, 333)
        if var.path_table == '':
            mes.error("Таблица с данными", "Выберите таблицу для обработки!")
        else:
            clear_frame()

            heads = ['№', 'Город', 'Фамилия']
            lst = get_nc()

            l1 = ttk.Label(fr4, text="Для заполнения полей выберите значение в таблице нажатием мыши!")
            l1.pack(side=TOP, padx=5, pady=5)

            table = ttk.Treeview(fr4, show='headings')
            table['columns'] = heads

            for header in heads:
                table.heading(header, text=header, anchor=CENTER)
                table.column(header, anchor=CENTER, width=1)
            for row in lst:
                table.insert('', END, values=row)

            scroll = ttk.Scrollbar(fr4, command=table.yview)
            table.configure(yscrollcommand=scroll.set)
            scroll.pack(side=RIGHT, fill=Y)

            table.bind('<<TreeviewSelect>>', item_selected)

            table.pack(expand=True, side=TOP, padx=5, fill=BOTH)

    def get_table():
        if service.open_table():
            l0['foreground'] = 'green'
        else:
            l0['foreground'] = 'red'

    def full_exit():
        service.full_exit(sh)

    sh = Toplevel()
    sh.geometry('450x610+100+50')
    sh.title('Вывод значений')
    sh.grab_set()
    win.withdraw()
    sh.protocol("WM_DELETE_WINDOW", sh_closing)
    sh.resizable(False, False)
    sh.config(bg="#F1EEE9")

    fr5 = Frame(sh, bg="#F1EEE9")
    fr5.pack(fill=X, padx=5)

    fr0 = Frame(sh, bg="#F1EEE9")
    fr0.pack(fill=X, padx=5, pady=10)

    fr_two = Frame(sh, bg="#F1EEE9")
    fr_two.pack(fill=X, padx=5)

    fr1 = Frame(fr_two, bg="#F1EEE9")
    fr1.pack(side=LEFT, fill=X, padx=5)

    fr2 = Frame(fr_two, bg="#F1EEE9")
    fr2.pack(side=LEFT, fill=X, padx=5)

    fr3 = Frame(fr_two, bg="#F1EEE9")
    fr3.pack(side=LEFT, fill=X, padx=5)

    fr4 = Frame(sh, bg="#F1EEE9")
    fr4.pack(fill=X, padx=5)

    l0 = ttk.Label(fr0, text="Таблица с данными", foreground='red')
    l0.pack(side=LEFT, padx=5)

    if var.path_table != '':
        l0['foreground'] = 'green'
    else:
        l0['foreground'] = 'red'

    btn_clear = ttk.Button(fr0, text="Очистить поля", command=clear_eout)
    btn_clear.pack(side=RIGHT, padx=5)

    btn_show = ttk.Button(fr0, text="Показать таблицу", command=show)
    btn_show.pack(side=RIGHT, padx=5)

    eout_seria = ttk.Entry(fr1)
    eout_seria.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_numb = ttk.Entry(fr1)
    eout_numb.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_date = ttk.Entry(fr1)
    eout_date.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_code = ttk.Entry(fr1)
    eout_code.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_date_birth = ttk.Entry(fr1)
    eout_date_birth.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_place_of_birth = ttk.Entry(fr1)
    eout_place_of_birth.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_gender = ttk.Entry(fr1)
    eout_gender.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_fam = ttk.Entry(fr2)
    eout_fam.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_name = ttk.Entry(fr2)
    eout_name.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_otch = ttk.Entry(fr2)
    eout_otch.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_inn = ttk.Entry(fr2)
    eout_inn.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_snils = ttk.Entry(fr3)
    eout_snils.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_job = ttk.Entry(fr3)
    eout_job.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_email = ttk.Entry(fr3)
    eout_email.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    eout_city = ttk.Entry(fr3)
    eout_city.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    btn_full_exit = ttk.Button(fr5, text="Выйти из программы", command=full_exit)
    btn_full_exit.pack(side=RIGHT, padx=5, pady=15)

    m_menu = Menu(sh)
    sh.config(menu=m_menu)

    # Выбрать таблицу
    in_menu = Menu(m_menu, tearoff=0)
    in_menu.add_command(label="Указать таблицу c данными", command=get_table)
    m_menu.add_cascade(label="Таблица", menu=in_menu, )

    # Запросы с почты
    mail_menu = Menu(m_menu, tearoff=0)
    mail_menu.add_command(label="Получить запросы с почты", command=get_request_from_mail)
    mail_menu.add_command(label="Добавить", command=open_custom_fill)
    m_menu.add_cascade(label="Данные", menu=mail_menu, )