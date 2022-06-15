from docxtpl import DocxTemplate
import openpyxl
from tkinter import *
from tkinter import messagebox
import os
from pathlib import Path
from tkinter import filedialog
from tkinter import ttk
import re
import faker
from random import randrange

path_sample = ''
path_table = ''
user_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents', 'Шаблоны')


def error(title, text):
    messagebox.showerror(title=title, message=text)


def info(title, text):
    messagebox.showinfo(title=title, message=text)


def open_sample_doc():
    global path_sample
    path_sample = filedialog.askopenfilename(title="Выбор шаблона для заполнения", initialdir=user_path,
                                             filetypes=(("Документы (*.docx)",
                                                         "*.docx"),
                                                        ("Все файлы", "*.*")))
    if path_sample != '':
        label3['foreground'] = 'green'
    else:
        label3['foreground'] = 'red'


def open_table():
    global path_table
    path_table = filedialog.askopenfilename(title="Выбор таблицы для заполнения", initialdir=user_path,
                                            filetypes=(("Таблицы (*.xlsx)",
                                                        "*.xlsx"),
                                                       ("Все файлы", "*.*")))
    if path_table != '':
        label4['foreground'] = 'green'
        wb = openpyxl.load_workbook(path_table)
        sheet = wb.get_sheet_by_name('Доверенность')
        e_end.delete(0, END)
        e_end.insert(0, sheet['M'][-1].coordinate)
        wb.close()
    else:
        label4['foreground'] = 'red'


def get_dict():
    start = e_start.get().upper()
    end = e_end.get().upper()
    test = []
    wb = openpyxl.load_workbook(f'{path_table}')
    sheet = wb.get_sheet_by_name('Доверенность')

    rows = 0

    try:
        for row in sheet[f'{start}':f'{end}']:
            rows += 1
            for cellObj in row:
                if cellObj.value == None or cellObj.value == '':
                    continue
                test.append(cellObj.value)

        info('Обработка входных данных', "Данные из таблицы успешно загружены!\nПереходим к "
                                         "обработке шаблона.\n"
                                         f"Было обработано строк: {rows}")
        wb.close()
        return test
    except ValueError:
        error("Обработка входных данных", "Введите корректные координаты!")
        wb.close()
    except:
        error("Обработка входных данных", "Ошибка обработки данных!")
        wb.close()


def start(test):
    x = 0
    files = 0
    path = Path(user_path + '/Доверенности')

    if not path.exists():
        os.mkdir(path)

    try:
        while x < len(test):
            doc = DocxTemplate(f'{path_sample}')
            context = {'director': test[x + 2], 'doc': test[x + 3], 'job': test[x + 4], 'name_of_subject': test[x + 5],
                       'seria': test[x + 6], 'num': test[x + 7], 'date': test[x + 8], 'passport_from_1': test[x + 9],
                       'job_of_director': test[x + 10], 'name_of_director': test[x + 11]}
            doc.render(context)
            path_dirs = Path(f'{path}/{test[x]}')
            if not path_dirs.exists():
                os.mkdir(path_dirs)
            doc.save(f'{path}/{test[x]}/Доверенность ' + test[x + 1] + '.docx')
            x += 12  # прыжок на следующую строку
            files += 1
        messagebox.showinfo(title="Обработка входных данных", message="Шаблон успешно обработан.\n"
                                                                      f"Создано файлов: {files}.")
    except:
        error("Обработка входных данных", "Ошибка заполнения шаблона!")


def check_text():
    if e_start.get().isalnum() and e_end.get().isalnum():
        return True
    else:
        return False


def fill_files():
    if check_text():
        if path_sample != '' and path_table != '':
            start(get_dict())
            e_start.delete(0, END)
            e_end.delete(0, END)
        else:
            error("Ошибка входных данных", "Выберите шаблон и таблицу с данными!")
    else:
        error("Ошибка входных данных", "Введите начальную и конечную точку таблицы!")


def open_dir():
    path = Path(user_path + '/Доверенности')

    if path.exists():
        os.startfile(path)
    else:
        error("Открытие папки", "Каталог не существует!")


def open_custom_fill():
    path_xlsx = ''

    def check_entry():
        if gender_data != '' and e_fio.get() and e_director.get() and e_doc.get() and \
                e_job.get() and e_name.get() and e_seria.get() and e_number.get() and e_date.get() and \
                e_point_of_get.get() and e_directors_job.get() and e_directors_name.get() and \
                e_code.get() and e_birth.get() and e_place_of_birth.get() and gender.get() and e_full_fio.get() and \
                e_inn.get() and e_snils.get() and e_mail.get():
            return True
        else:
            return False

    def open_xlsx():
        nonlocal path_xlsx
        path_xlsx = filedialog.askopenfilename(title="Выберите вашу таблицу с данными", initialdir=user_path,
                                               filetypes=(("Таблицы (*.xlsx)",
                                                           "*.xlsx"),
                                                          ("Все файлы", "*.*")))
        if path_xlsx != '':
            label13['foreground'] = 'green'
        else:
            label13['foreground'] = 'red'

    def open_xlsx_sample():
        path = Path(user_path)

        if path.exists():
            os.startfile(path)
        else:
            error("Открытие папки", "Каталог не существует!")

    def create_xlsx():
        book = openpyxl.Workbook()
        book.remove(book.active)
        sheet_1 = book.create_sheet("Доверенность")
        sheet_2 = book.create_sheet("Заявление")

        sheet_1.insert_rows(0)
        col_names = ['№', 'Город', 'На кого', 'Руководитель', 'Основание', 'Должность', 'ФИО', 'Серия паспорта',
                     'Номер паспорта', 'Дата выдачи', 'Кем выдан', 'Должность руководителя', 'ФИО руководителя']
        l = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
             'V', 'W', 'X', 'Y', 'Z']
        n = 0
        for i in col_names:
            sheet_1[f'{l[n]}1'].value = f'{col_names[n]}'
            sheet_1.column_dimensions[l[n]].width = 20
            n += 1

        sheet_2.insert_rows(0)
        col2_names = ['№', 'Серия', 'Номер', 'Дата выдачи', 'Код', 'Дата рождения', 'Место рождения', 'Пол',
                      'Фамилия', 'Имя', 'Отчество', 'ИНН', 'СНИЛС', 'Должность', 'Почта', 'Населенный пункт']
        k = 0
        for i in col2_names:
            sheet_2[f'{l[k]}1'].value = f'{col2_names[k]}'
            sheet_2.column_dimensions[l[k]].width = 20
            k += 1

        path = Path(user_path)
        if not path.exists():
            os.mkdir(path)

        sheet_1.auto_filter.ref = "B1:M999"
        sheet_2.auto_filter.ref = "B1:P999"

        book.save(f"{path}/Переменные.xlsx")
        info("Создание таблицы", 'Таблица успешно создана!')

    def clear_xlsx():
        if path_xlsx != '':
            question = messagebox.askokcancel(title="Очистка таблицы", message="Вы уверены, что хотите очистить "
                                                                               "таблицу?\nВсе "
                                                                               "данные в таблице будут безвозвратно удалены!")
            if question:
                book = openpyxl.load_workbook(path_xlsx)
                sheet = book["Доверенность"]
                sheet2 = book["Заявление"]
                last_row = len(list(sheet.rows))
                first_row = 1
                sheet.delete_rows(first_row, last_row)
                last_row2 = len(list(sheet2.rows))
                first_row2 = 1
                sheet2.delete_rows(first_row2, last_row2)
                del1 = last_row - first_row
                del2 = last_row2 - first_row2
                if del1 < 0:
                    del1 = 0
                if del2 < 0:
                    del2 = 0
                book.save(path_xlsx)
                info("Очистка таблицы", f'Успешно удалено строк: {del1} в листе "Доверенность" и '
                                        f'{del2} в листе "Заявление"')


        else:
            error("Ошибка входных данных", "Выберите таблицу с данными!")

    gender_data = ''

    def insert_data():
        clear_for_fill()
        e_director.insert(0, 'Главного врача Филиной Натальи Григорьевны')
        e_doc.insert(0, 'Устава, утвержденного Приказом от 28.12.2021 №2586-орг')
        e_directors_job.insert(0, 'Главный врач')
        e_directors_name.insert(0, 'Н.Г. Филина')
        e_mail.insert(0, 'ikdomashenko@kkck.ru')

    def get_data():
        city = combo_city.get().rstrip()
        data = [[city, e_fio.get(), e_director.get(), e_doc.get(), e_job.get(), e_name.get(), e_seria.get(),
                 e_number.get(), e_date.get(), e_point_of_get.get(), e_directors_job.get(), e_directors_name.get()]]
        return data

    def get_data2():
        temp = e_full_fio.get()
        temp2 = temp.split()
        if len(temp2) < 3:
            error("Ошибка входных данных", "Укажите верное ФИО (полное)!")
        else:
            fam = temp2[0]
            name = temp2[1]
            otch = temp2[2]
            temp = e_snils.get()
            snils = "".join(c for c in temp if c.isalnum())
            city = combo_city.get().rstrip()
            data2 = [
                [
                    e_seria.get(), e_number.get(), e_date.get(), e_code.get(), e_birth.get(), e_place_of_birth.get(),
                    gender_data, fam, name, otch, e_inn.get(), snils, e_job.get(), e_mail.get(), city
                ]
            ]
            return data2

    def confirm_data():
        if path_xlsx == '':
            error("Ошибка входных данных", "Укажите таблицу данных!")
        else:
            if check_entry():
                wb = openpyxl.load_workbook(path_xlsx)
                sheet = wb["Доверенность"]
                temp = get_data()
                try:
                    number = int(sheet["A"][-1].value) + 1
                except:
                    number = 1
                data = temp
                data[0].insert(0, number)
                for row in data:
                    sheet.append(row)
                # wb1.save(path_xlsx)

                # wb2 = openpyxl.load_workbook(path_xlsx)
                sheet2 = wb["Заявление"]
                temp2 = get_data2()
                try:
                    number = int(sheet2["A"][-1].value) + 1
                except:
                    number = 1
                data2 = temp2
                data2[0].insert(0, number)
                for row in data2:
                    sheet2.append(row)
                wb.save(path_xlsx)
                info('Обработка данных', f'Файл успешно сохранен! Добавлено строк: {len(data)}')
                clear_data()
            else:
                error("Ошибка входных данных", "Введите требуемые значения!")

    def del_city():
        que = messagebox.askokcancel(title="Удалить город из списка",
                                     message=f"Вы уверены, что хотите удалить {combo_city.get()} из списка?")
        if que:
            path_dir_txt = Path(user_path + '/Города')
            file_path = Path(rf"{path_dir_txt}\cities.txt")
            with open(file_path) as f:
                lines = f.readlines()

            str = combo_city.get()
            pattern = re.compile(re.escape(str))
            with open(file_path, 'w') as f:
                for line in lines:
                    result = pattern.search(line)
                    if result is None:
                        f.write(line)

            upd_cities = r_cities_txt()
            combo_city.config(values=upd_cities)
            combo_city.current(2)

    def clear_for_fill():
        e_director.delete(0, END)
        e_doc.delete(0, END)
        e_directors_job.delete(0, END)
        e_directors_name.delete(0, END)
        e_mail.delete(0, END)

    def clear_data():
        e_fio.delete(0, END)
        e_director.delete(0, END)
        e_doc.delete(0, END)
        e_job.delete(0, END)
        e_name.delete(0, END)
        e_seria.delete(0, END)
        e_number.delete(0, END)
        e_point_of_get.delete(0, END)
        e_directors_job.delete(0, END)
        e_directors_name.delete(0, END)
        e_date.delete(0, END)

        e_code.delete(0, END)
        e_birth.delete(0, END)
        e_place_of_birth.delete(0, END)

        e_full_fio.delete(0, END)
        e_inn.delete(0, END)
        e_snils.delete(0, END)
        e_mail.delete(0, END)

    def select():
        nonlocal gender_data
        l = gender.get()
        if l == 1:
            gender_data = 'М'
        elif l == 2:
            gender_data = 'Ж'

    def on_closing():
        win.destroy()
        root.deiconify()

    def close_ad(city):
        w_cities_txt(city)
        upd_cities = r_cities_txt()
        combo_city.config(values=upd_cities)
        combo_city.current(len(upd_cities) - 1)

    def r_cities_txt():
        path_dir_txt = Path(user_path + '/Города')
        if not path_dir_txt.exists():
            os.mkdir(path_dir_txt)

        file_path = Path(rf"{path_dir_txt}\cities.txt")
        if not file_path.exists():
            my_file = open(rf"{path_dir_txt}\cities.txt", "w+")
            for city in cities:
                my_file.write(city + f'\n')
            my_file.close()

        data = []
        r_file = open(rf"{path_dir_txt}\cities.txt", "r").readlines()
        for line in r_file:
            data.append(line)
        return data

    def w_cities_txt(temp):
        path_dir_txt = Path(user_path + '/Города')
        if not path_dir_txt.exists():
            os.mkdir(path_dir_txt)

        file_path = Path(rf"{path_dir_txt}\cities.txt")
        if not file_path.exists():
            my_file = open(rf"{path_dir_txt}\cities.txt", "w+")
            for city in cities:
                my_file.write(city + f'\n')
            my_file.write(temp + f'\n')
            my_file.close()
        else:
            my_file = open(rf"{path_dir_txt}\cities.txt", "a+")
            my_file.write(temp + f'\n')
            my_file.close()

    def update_cities():
        def close_win_ad():
            ad.destroy()
            win.deiconify()

        def add():
            city = e_new_city.get()
            if city != '':
                close_ad(city)
                ad.destroy()
                win.deiconify()
            else:
                error("Ошибка ввода", "Вы ввели некорректное значение!")

        ad = Toplevel()
        ad.geometry('400x100+100+50')
        ad.title('Введите новый город')
        ad.grab_set()
        win.withdraw()
        ad.protocol("WM_DELETE_WINDOW", close_win_ad)
        ad.resizable(False, False)
        ad.config(bg="#F1EEE9")

        f0 = Frame(ad, bg="#F1EEE9")
        f0.pack(fill=BOTH, padx=10, pady=5)

        e_new_city = Entry(f0)
        e_new_city.pack(fill=X, padx=10, ipady=2, expand=True)

        btn_add_city = ttk.Button(f0, text="Принять", command=add)
        btn_add_city.pack(side=BOTTOM, padx=5)

    def out_xlsx():
        pth_xlsx = ''

        def open_xlsx_sh():
            nonlocal pth_xlsx
            pth_xlsx = filedialog.askopenfilename(title="Выберите вашу таблицу с данными", initialdir=user_path,
                                                  filetypes=(("Таблицы (*.xlsx)",
                                                              "*.xlsx"),
                                                             ("Все файлы", "*.*")))
            if pth_xlsx != '':
                l0['text'] = 'Таблица с данными выбрана!'
                l0['foreground'] = 'green'
            else:
                l0['text'] = 'Таблица с данными не выбрана!'
                l0['foreground'] = 'red'

        def sh_closing():
            sh.destroy()
            win.deiconify()

        def get_nc():
            if pth_xlsx == '':
                error("Таблица с данными", "Выберите таблицу для обработки!")
            else:
                wb = openpyxl.load_workbook(pth_xlsx)
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
                    error("Обработка входных данных", "Ошибка обработки данных!")
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
                clear_eout() # очищаем поля

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
                if pth_xlsx == '':
                    error("Таблица с данными", "Выберите таблицу для обработки!")
                else:
                    wb = openpyxl.load_workbook(pth_xlsx)
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
                        error("Обработка входных данных", "Ошибка обработки данных!")
                        wb.close()

            if pth_xlsx == '':
                error("Таблица с данными", "Выберите таблицу для обработки!")
            else:
                for widget in fr4.winfo_children():
                    widget.destroy()

                heads = ['№', 'Город', 'Фамилия']
                lst = get_nc()

                l1 = ttk.Label(fr4, text="Для заполнения полей выберите значение в таблице нажатием мыши!")
                l1.pack(side=TOP, padx=5, pady=5)

                table = ttk.Treeview(fr4, show='headings')
                table['columns'] = heads

                for header in heads:
                    table.heading(header, text=header, anchor=CENTER)
                    table.column(header, anchor=CENTER)
                for row in lst:
                    table.insert('', END, values=row)

                scroll = ttk.Scrollbar(fr4, command=table.yview)
                table.configure(yscrollcommand=scroll.set)
                scroll.pack(side=RIGHT, fill=Y)

                table.bind('<<TreeviewSelect>>', item_selected)

                table.pack(expand=True, side=TOP, padx=5, fill=BOTH)

        sh = Toplevel()
        sh.geometry('450x500+100+50')
        sh.title('Вывод значений')
        sh.grab_set()
        win.withdraw()
        sh.protocol("WM_DELETE_WINDOW", sh_closing)
        sh.resizable(False, False)
        sh.config(bg="#F1EEE9")

        fr0 = Frame(sh, bg="#F1EEE9")
        fr0.pack(fill=X, padx=5)

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

        l0 = ttk.Label(fr0, text="Таблица данных не выбрана!", foreground='red')
        l0.pack(side=LEFT, padx=5)

        btn_show = ttk.Button(fr0, text="Показать таблицу", command=show)
        btn_show.pack(side=LEFT, padx=5)

        btn_clear = ttk.Button(fr0, text="Очистить поля", command=clear_eout)
        btn_clear.pack(side=LEFT, padx=5)

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

        m_menu = Menu(sh)
        sh.config(menu=m_menu)

        # Выбрать таблицу
        in_menu = Menu(m_menu, tearoff=0)
        in_menu.add_command(label="Указать таблицу c данными", command=open_xlsx_sh)
        m_menu.add_cascade(label="Таблица", menu=in_menu, )

    win = Toplevel()
    win.geometry('1500x450+100+50')
    win.title('Внесение данных в таблицу')
    win.grab_set()
    root.withdraw()
    win.protocol("WM_DELETE_WINDOW", on_closing)
    win.resizable(True, True)
    win.minsize(1300, 450)
    win.maxsize(1800, 550)
    win.config(bg="#F1EEE9")
    win.columnconfigure(index=0, minsize=550, weight=550, pad=2)
    win.columnconfigure(index=1, minsize=600, weight=550, pad=2)
    win.columnconfigure(index=2, weight=50, pad=1)

    f0 = Frame(win, bg="#F1EEE9")
    # f0.pack(fill=X, padx=10, pady=5)
    f0.grid(row=0, column=0, sticky=W + E, padx=5)

    f1 = Frame(win, bg="#F1EEE9")
    # f1.pack(fill=X, padx=10, pady=5)
    f1.grid(row=1, column=0, sticky=W + E, padx=5)

    f2 = Frame(win, bg="#F1EEE9")
    # f2.pack(fill=X, padx=10, pady=5)
    f2.grid(row=2, column=0, sticky=W + E, padx=5)

    f3 = Frame(win, bg="#F1EEE9")
    # f3.pack(fill=X, padx=10, pady=5)
    f3.grid(row=3, column=0, sticky=W + E, padx=5)

    f3_1 = Frame(win, bg="#F1EEE9")
    # f3_1.pack(fill=X, padx=10, pady=1)
    f3_1.grid(row=6, column=0, sticky=W + E, padx=5)

    f4 = Frame(win, bg="#F1EEE9")
    # f4.pack(fill=X, padx=10, pady=5)
    f4.grid(row=7, column=0, sticky=W + E, padx=5)

    f5 = Frame(win, bg="#F1EEE9")
    # f5.pack(fill=X, padx=10, pady=5)
    f5.grid(row=8, column=0, sticky=W + E, padx=5)

    f6 = Frame(win, bg="#F1EEE9")
    # f6.pack(fill=X, padx=10, pady=5)
    f6.grid(row=9, column=0, sticky=W + E, padx=5)

    f7 = Frame(win, bg="#F1EEE9")
    # f7.pack(fill=X, padx=10, pady=5)
    f7.grid(row=10, column=0, sticky=W + E, padx=5)

    f8 = Frame(win, bg="#F1EEE9")
    # f8.pack(fill=X, padx=10, pady=5)
    f8.grid(row=1, column=1, sticky=W + E, padx=5)

    f9 = Frame(win, bg="#F1EEE9")
    # f9.pack(fill=X, padx=10, pady=5)
    f9.grid(row=2, column=1, sticky=W + E, padx=5)

    f10 = Frame(win, bg="#F1EEE9")
    # f10.pack(fill=X, padx=10, pady=5)
    f10.grid(row=3, column=1, sticky=W + E, padx=5)

    f11 = Frame(win, bg="#F1EEE9")
    # f11.pack(fill=X, padx=10, pady=5)
    f11.grid(row=4, column=1, sticky=W + E, padx=5)

    f12 = Frame(win, bg="#F1EEE9")
    # f12.pack(fill=X, padx=10, pady=5)
    f12.grid(row=5, column=1, sticky=W + E, padx=5)

    f13 = Frame(win, bg="#F1EEE9")
    # f13.pack(fill=X, padx=10, pady=5)
    f13.grid(row=0, column=1, sticky=W + E, padx=5)

    f14 = Frame(win, bg="#F1EEE9")
    # f14.pack(fill=X, padx=10, pady=5)
    f14.grid(row=6, column=1, sticky=W + E, padx=5)

    f15 = Frame(win, bg="#F1EEE9")
    # f15.pack(fill=X, padx=10, pady=5)
    f15.grid(row=7, column=1, sticky=W + E, padx=5)

    f16 = Frame(win, bg="#F1EEE9")
    # f16.pack(fill=X, padx=10, pady=5)
    f16.grid(row=8, column=1, sticky=W + E, padx=5)

    f17 = Frame(win, bg="#F1EEE9")
    # f17.pack(fill=X, padx=10, pady=5)
    f17.grid(row=9, column=1, sticky=W + E, padx=5)

    f18 = Frame(win, bg="#F1EEE9")
    # f18.pack(fill=X, padx=10, pady=5)
    f18.grid(row=10, column=1, sticky=W + E, padx=5)

    f19 = Frame(win, bg="#F1EEE9")
    # f19.pack(fill=X, padx=10, pady=5)
    f19.grid(row=11, column=1, sticky=W + E, padx=5)

    f20 = Frame(win, bg="#F1EEE9")
    # f20.pack(fill=X, padx=10, pady=5)
    f20.grid(row=11, column=0, sticky=W + E, padx=5)

    f21 = Frame(win, bg="#F1EEE9")
    # f21.pack(fill=X, padx=10, pady=5)
    f21.grid(row=12, column=0, sticky=W + E, padx=5)

    label_0 = Label(f0, width=25, text='Город:')
    label_0.pack(side=LEFT, pady=5)

    btn_add_city = ttk.Button(f0, text="Другой", command=update_cities)
    btn_add_city.pack(side=LEFT, padx=5)

    cities = ['Ачинск', 'Канск', 'Красноярск', 'Лесосибирск', 'Минусинск', 'Норильск']
    actual_cities = r_cities_txt()
    combo_city = ttk.Combobox(f0, values=actual_cities)
    combo_city.current(2)
    combo_city.pack(side=LEFT, fill=X, padx=10, pady=5)

    btn_del_city = ttk.Button(f0, text="<- Удалить", command=del_city)
    btn_del_city.pack(side=LEFT, padx=2, pady=5)

    label_1 = Label(f1, width=25, text='Фамилия И.О.:')
    label_1.pack(side=LEFT, pady=5)

    e_fio = Entry(f1)
    e_fio.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_2 = Label(f2, width=25, text='В лице руководителя:')
    label_2.pack(side=LEFT, pady=5)

    e_director = Entry(f2)
    e_director.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_3 = Label(f3, width=25, text='На основании:')
    label_3.pack(side=LEFT, pady=5)

    e_doc = Entry(f3)
    e_doc.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_3_1 = Label(f3_1, width=25, text='Уполномочивает', background='grey', foreground='yellow')
    label_3_1.pack(side=TOP, pady=5)

    label_4 = Label(f4, width=25, text='Должность:')
    label_4.pack(side=LEFT, pady=5)

    e_job = Entry(f4)
    e_job.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_5 = Label(f5, width=25, text='Сотрудника (полн.ФИО):')
    label_5.pack(side=LEFT, pady=5)

    e_name = Entry(f5)
    e_name.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_17 = Label(f6, width=25, text='ИНН:')
    label_17.pack(side=LEFT, pady=5)

    e_inn = Entry(f6)
    e_inn.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_18 = Label(f7, width=25, text='СНИЛС:')
    label_18.pack(side=LEFT, pady=5)

    e_snils = Entry(f7)
    e_snils.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_6 = Label(f8, width=25, text='Серия паспорта:')
    label_6.pack(side=LEFT, pady=5)

    e_seria = Entry(f8)
    e_seria.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_7 = Label(f9, width=25, text='Номер паспорта:')
    label_7.pack(side=LEFT, pady=5)

    e_number = Entry(f9)
    e_number.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_9 = Label(f10, width=25, text='Кем выдан:')
    label_9.pack(side=LEFT, pady=5)

    e_point_of_get = Entry(f10)
    e_point_of_get.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_8 = Label(f11, width=25, text='Дата выдачи паспорта:')
    label_8.pack(side=LEFT, pady=5)

    e_date = Entry(f11)
    e_date.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_12 = Label(f12, width=25, text='Код места выдачи:')
    label_12.pack(side=LEFT, pady=5)

    e_code = Entry(f12)
    e_code.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_16 = Label(f13, width=25, text='ФИО (полное):')
    label_16.pack(side=LEFT, pady=5)

    e_full_fio = Entry(f13)
    e_full_fio.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_15 = Label(f14, width=25, text='Пол:')
    label_15.pack(side=LEFT, pady=5)

    gender = IntVar()

    male_checkbutton = Radiobutton(f14, text="М", value=1, variable=gender, padx=10, command=select)
    male_checkbutton.pack(side=LEFT, pady=5)

    female_checkbutton = Radiobutton(f14, text="Ж", value=2, variable=gender, padx=10, command=select)
    female_checkbutton.pack(side=LEFT, pady=5)

    label_13 = Label(f15, width=25, text='Дата рождения:')
    label_13.pack(side=LEFT, pady=5)

    e_birth = Entry(f15)
    e_birth.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_14 = Label(f16, width=25, text='Место рождения:')
    label_14.pack(side=LEFT, pady=5)

    e_place_of_birth = Entry(f16)
    e_place_of_birth.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_10 = Label(f17, width=25, text='Должность руководителя:')
    label_10.pack(side=LEFT, pady=5)

    e_directors_job = Entry(f17)
    e_directors_job.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_11 = Label(f18, width=25, text='И.О.Фамилия руководителя:')
    label_11.pack(side=LEFT, pady=5)

    e_directors_name = Entry(f18)
    e_directors_name.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    label_19 = Label(f19, width=25, text='Почта:')
    label_19.pack(side=LEFT, pady=5)

    e_mail = Entry(f19)
    e_mail.pack(fill=X, padx=10, ipady=2, expand=True, pady=5)

    btn_confirm = ttk.Button(f20, text="Принять", command=confirm_data)
    btn_confirm.pack(side=RIGHT, pady=5, padx=10)

    btn_clear_e = ttk.Button(f20, text="Очистить", command=clear_data)
    btn_clear_e.pack(side=LEFT, padx=5, pady=5)

    btn_fill_e = ttk.Button(f20, text="Заполнить", command=insert_data)
    btn_fill_e.pack(side=LEFT, padx=5, pady=5)

    btn_fill_e = ttk.Button(f20, text="Шаблон", command=open_xlsx_sample)
    btn_fill_e.pack(side=LEFT, padx=17, pady=5)

    label12 = ttk.Label(f21, text="Статус загрузки файлов:")
    label12.pack(side=LEFT, padx=10, pady=5)

    label13 = ttk.Label(f21, text="Таблица с данными", foreground='red', justify=LEFT)
    label13.pack(side=RIGHT, padx=10, pady=5)

    main_menu = Menu(win)
    win.config(menu=main_menu)

    # Выбрать таблицу
    file_menu = Menu(main_menu, tearoff=0)
    file_menu.add_command(label="Выбрать таблицу c данными", command=open_xlsx)
    file_menu.add_command(label="Очистить таблицу", command=clear_xlsx)
    file_menu.add_command(label="Создать таблицу", command=create_xlsx)
    main_menu.add_cascade(label="Таблица", menu=file_menu, )

    # Вывод значений
    out_menu = Menu(main_menu, tearoff=0)
    out_menu.add_command(label="Вывести значения таблицы", command=out_xlsx)
    main_menu.add_cascade(label="Вывод", menu=out_menu, )

    insert_data()


root = Tk()
root.title('Доверенность')
root.geometry('385x150+100+50')
root.resizable(False, False)

main_menu = Menu(root)
root.config(menu=main_menu)

# Автоматический ввод данных
file_menu = Menu(main_menu, tearoff=0)
file_menu.add_command(label="Выбрать таблицу c данными", command=open_table)
file_menu.add_command(label="Выбрать шаблон", command=open_sample_doc)
main_menu.add_cascade(label="Шаблон", menu=file_menu, )

# Рцчной ввод данных
fill_menu = Menu(main_menu, tearoff=0)
fill_menu.add_command(label="Заполнить вручную", command=open_custom_fill)
main_menu.add_cascade(label="Данные", menu=fill_menu, )

f1 = Frame(root)
f1.pack(fill=X, padx=10, pady=5)

f2 = Frame(root)
f2.pack(fill=X, padx=10, pady=5)

f3 = Frame(root)
f3.pack(fill=X, padx=10, pady=5)

f4 = Frame(root)
f4.pack(fill=X, padx=10, pady=5)

label1 = ttk.Label(f1, text="Начальная точка:", width=17)
label1.pack(side=LEFT, padx=10)

e_start = ttk.Entry(f1)
e_start.pack(side=RIGHT, fill=X, expand=True)

label2 = ttk.Label(f2, text="Конечная точка:", width=17)
label2.pack(side=LEFT, padx=10)

e_end = ttk.Entry(f2)
e_end.pack(side=RIGHT, fill=X, expand=True)

btn_show_table = ttk.Button(f3, text="Заполнить", command=fill_files)
btn_show_table.pack(side=RIGHT)

btn_show_dir = ttk.Button(f3, text="Результат", command=open_dir)
btn_show_dir.pack(side=LEFT, padx=10)

label3_1 = ttk.Label(f4, text="Статус загрузки файлов: ")
label3_1.pack(side=LEFT, padx=10)

label3 = ttk.Label(f4, text="Шаблоны", foreground='red', justify=LEFT)
label3.pack(side=BOTTOM, padx=10)

label4 = ttk.Label(f4, text="Таблица с данными", foreground='red', justify=LEFT)
label4.pack(side=BOTTOM, padx=10)

e_start.delete(0, END)
e_start.insert(0, 'B2')

root.mainloop()
