import os
from pathlib import Path
import variables as var
import messages as mes


# Проверка запрашиваемого пути
def check_path(path):
    if os.path.exists(path):
        return True
    else:
        return False


# Проверка на пустоту каталога
def empty_or_not(path):
    return next(os.scandir(path), None)


def open_xlsx_sample():
    path = Path(var.sample_folder_path)

    if path.exists():
        os.startfile(path)
    else:
        mes.error("Открытие папки", "Каталог не существует!")


def check_table():
    if check_path(var.table_xlsx_file_path):
        return True
    else:
        return False
