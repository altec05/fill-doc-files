from docxtpl import DocxTemplate
import openpyxl
from tkinter import *
from tkinter import messagebox
import os
from pathlib import Path
from tkinter import filedialog
from tkinter import ttk

path_sample = ''
path_table = ''


def open_sample_doc():
    global path_sample
    path_sample = filedialog.askopenfilename(title="Выбор шаблона для заполнения",
                                             filetypes=(("Документы (*.docx)",
                                                         "*.docx"),
                                                        ("Все файлы", "*.*")))


def open_table():
    global path_table
    path_table = filedialog.askopenfilename(title="Выбор таблицы для заполнения",
                                            filetypes=(("Таблицы (*.xlsx)",
                                                        "*.xlsx"),
                                                       ("Все файлы", "*.*")))


def get_dict():
    start = e_start.get().upper()
    end = e_end.get().upper()
    test = []
    wb = openpyxl.load_workbook(f'{path_table}')
    sheet = wb.active

    rows = 0

    try:
        for row in sheet[f'{start}':f'{end}']:
            rows += 1
            for cellObj in row:
                if cellObj.value == None or cellObj.value == "":
                    continue
                test.append(cellObj.value)
        messagebox.showinfo(title="Обработка входных данных",
                            message="Данные из таблицы успешно загружены!\nПереходим к "
                                    "обработке шаблона.\n"
                                    f"Было обработано строк: {rows}")
        return test
    except ValueError:
        messagebox.showerror(title="Обработка входных данных", message="Введите корректные коордиаты!")
    except:
        messagebox.showerror(title="Обработка входных данных", message="Ошибка обработки данных!")


def start(test):
    x = 0
    files = 0

    path = Path('Доверенности')
    if not path.exists():
        os.mkdir(path)

    try:
        while x < len(test):
            doc = DocxTemplate(f'{path_sample}')
            context = {'director': test[x + 1], 'doc': test[x + 2], 'job': test[x + 3], 'name_of_subject': test[x + 4],
                       'seria': test[x + 5], 'num': test[x + 6], 'date': test[x + 7], 'passport_from': test[x + 8],
                       'job_of_director': test[x + 9], 'name_of_director': test[x + 10]}
            doc.render(context)
            doc.save('Доверенности/Доверенность ' + test[x] + '.docx')
            x += 11  # прыжок на следующую строку
            files += 1
        messagebox.showinfo(title="Обработка входных данных", message="Шаблон успешно обработан.\n"
                                                                      f"Создано файлов: {files}.")
    except:
        messagebox.showerror(title="Обработка входных данных", message="Ошибка заполнения шаблона!")


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
            messagebox.showerror(title="Ошибка входных данных", message="Выберите шаблон и таблицу с данными!")

    else:
        messagebox.showerror(title="Ошибка входных данных", message="Введите начальную и конечную точку таблицы!")


def open_dir():
    path = Path('Доверенности')
    if path.exists():
        os.startfile(path)
    else:
        messagebox.showerror(title="Открытие папки", message="Каталог не существует!")


root = Tk()
root.title('Доверенность')
root.geometry('300x100+500+250')
root.resizable(False, False)

main_menu = Menu(root)
root.config(menu=main_menu)

# Ввод данных
file_menu = Menu(main_menu, tearoff=0)
file_menu.add_command(label="Выбрать таблицу c данными", command=open_table)
file_menu.add_command(label="Выбрать шаблон", command=open_sample_doc)
main_menu.add_cascade(label="Шаблон", menu=file_menu, )

f1 = Frame(root)
f1.pack(fill=X, padx=10, pady=5)

f2 = Frame(root)
f2.pack(fill=X, padx=10, pady=5)

f3 = Frame(root)
f3.pack(fill=X, padx=10, pady=5)

label1 = ttk.Label(f1, text="Начальная точка:")
label1.pack(side=LEFT, padx=10)

e_start = ttk.Entry(f1)
e_start.pack(side=RIGHT, fill=X)

label2 = ttk.Label(f2, text="Конечная точка:")
label2.pack(side=LEFT, padx=10)

e_end = ttk.Entry(f2)
e_end.pack(side=RIGHT, fill=X)

btn_show_table = ttk.Button(f3, text="Заполнить", command=fill_files)
btn_show_table.pack(side=RIGHT)

btn_show_dir = ttk.Button(f3, text="Результат", command=open_dir)
btn_show_dir.pack(side=LEFT, padx=10)

root.mainloop()
