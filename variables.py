import os

sample_folder_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents', 'ПО Казначейство', 'Шаблоны')
cities_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents', 'ПО Казначейство', 'Города')
doverennost_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents', 'ПО Казначейство', 'Доверенности')
user_docs_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')

# Путь для рез копирования
path_to_backup = r"\\192.168.15.4\Soft\Программирование\Py\Казначейство\Резервная копия файлов\ПО Казначейство"
path_from_backup = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents', 'ПО Казначейство')

table_xlsx_file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents', 'ПО Казначейство', 'Шаблоны', 'Переменные.xlsx')

last_path = ''

path_sample = ''
path_table = ''

login = ''
password = ''
mail_code = '6dJfe98debiLgyatb0eE'
mail_imap = 'imap.mail.ru'

temp_sort = 0
