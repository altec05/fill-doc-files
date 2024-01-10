import os
from distutils.dir_util import copy_tree
from variables import path_to_backup, path_from_backup
from check_funcs import check_path, empty_or_not
import messages as mes
from tkinter import filedialog
import variables as var
import check_funcs as ch_foo


def full_exit(root):
    root.quit()


def open_sample_doc():
    filetypes = [("Документы (*.docx)", ".docx")]

    if var.last_path == '':
        var.path_sample = filedialog.askopenfilename(title="Укажите docx шаблон для заполнения", initialdir=var.user_docs_path, filetypes=filetypes)
        if var.path_sample != '':
            var.last_path = var.path_sample.replace(os.path.basename(var.path_sample), '')
    else:
        var.path_sample = filedialog.askopenfilename(title="Укажите docx шаблон для заполнения", initialdir=var.last_path, filetypes=filetypes)
        if var.path_sample != '':
            var.last_path = var.path_sample.replace(os.path.basename(var.path_sample), '')

    if var.path_sample != '':
        if ch_foo.check_path(var.path_sample):
            mes.info('Указание шаблона',
                     f'Успешно выбран шаблон: {os.path.basename(var.path_sample)}')
            return True
        else:
            mes.error('Указание шаблона',
                      f'Путь {var.path_sample} не существует!')
            return False
    else:
        mes.error('Указание шаблона',
                 f'Шаблон не был выбран!')
        return False


def open_table():
    filetypes = [("Таблицы (*.xlsx)", ".xlsx")]
    print(var.last_path)

    if var.last_path == '':
        var.path_table = filedialog.askopenfilename(title="Укажите xlsx таблицу с данными",
                                                     initialdir=var.user_docs_path, filetypes=filetypes)
        if var.path_table != '':
            var.last_path = var.path_table.replace(os.path.basename(var.path_table), '')
            print(var.last_path)
    else:
        var.path_table = filedialog.askopenfilename(title="Укажите xlsx таблицу с данными",
                                                     initialdir=var.last_path, filetypes=filetypes)
        if var.path_table != '':
            var.last_path = var.path_table.replace(os.path.basename(var.path_table), '')
            print(var.last_path)
    if var.path_table != '':
        if ch_foo.check_path(var.path_table):
            mes.info('Указание таблицы с данными',
                      f'Успешно выбрана таблица с данными: {os.path.basename(var.path_table)}')
            return True
        else:
            mes.error('Указание таблицы с данными',
                      f'Путь {var.path_table} не существует!')
            return False
    else:
        mes.error('Указание таблицы с данными',
                 f'Таблица с данными не выбрана!')
        return False


def clear_folder(path):
    import os, shutil
    folder = path
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
                print(f'Удалил {file_path}')
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
                print(f'Удалил {file_path}')
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


from_directory = rf'{path_from_backup}'
to_directory = rf'{path_to_backup}'


# Резервное копирование по заданному пути
# на данном этапе изменение пути для РК не реализовано, но предусмотрено с необходимыми проверками
def backup_bd(extra_path):
    from datetime import datetime
    now_date = datetime.now().date().strftime("%d.%m.%Y")
    final_path = ''

    if extra_path == '':
        final_path = os.path.join(to_directory, now_date)
    else:
        final_path = os.path.join(extra_path, now_date)

    try:
        if check_path(from_directory):
            if check_path(final_path):
                if empty_or_not(final_path) is not None:
                    if mes.ask('Проверка пути для копирования',
                               f'Внимание! Конечная папка содержит файлы. Очистить её и продолжить копирование?\n\n{final_path}'):
                        clear_folder(final_path)
                        result = copy_tree(from_directory, final_path)
                        mes.info('Резервное копирование файлов',
                                 f'Успешно скопировано файлов: {len(result)}.\n\nСкопированы в: "{final_path}".')
                    else:
                        mes.error('Резервное копирование файлов', f'Отмена операции пользователем!')
                else:
                    result = copy_tree(from_directory, final_path)
                    mes.info('Резервное копирование файлов',
                             f'Успешно скопировано файлов: {len(result)}.\n\nСкопированы в: "{final_path}".')
            else:
                os.makedirs(final_path, exist_ok=True)
                result = copy_tree(from_directory, final_path)
                mes.info('Резервное копирование файлов',
                         f'Успешно скопировано файлов: {len(result)}.\n\nСкопированы в: "{final_path}".')
        else:
            mes.error('Резервное копирование файлов', f'Ошибка: путь с шаблонами не существует!\n\n{from_directory}')
    except:
        mes.error('Резервное копирование БД', 'Ошибка: повторное копирование вызывает ошибку пути!')
