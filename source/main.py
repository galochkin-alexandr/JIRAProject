from tkinter.filedialog import *

from Project import *
from GISMUException import *
import pandas as pd
from settings import get_settings

"""Достаём конфиги"""
settings = get_settings()
server_name = settings.server_name
project_name = settings.project_name
token_auth = settings.token_auth

"""Выбираем папки"""
path = askdirectory()
array_of_dir = []
main_file = ''

try:
    gismu3lp_project = Project(server_name, project_name, token_auth)

    """main_file - файл с обращениями, array_of_dir - массив папок с вложениями"""
    with os.scandir(path) as all_files:
        for current_file in all_files:
            if current_file.is_file():
                main_file = current_file
            else:
                array_of_dir.append(current_file)

    """result - DataFrame с обработанными обращениями"""
    result = pd.DataFrame({'Имя': [], 'SD': [], 'Метки': [], 'Регион': [], 'Воспроизводится': [],
        'Название': [], 'Описание': [], 'Действие': []})

    all_issue = pd.read_excel(current_file, sheet_name=0).values.tolist()
    for current_issue in all_issue:
        sd = current_issue[0]
        try:
            if isinstance(sd, int) or pd.isnull(sd):
                sd = 'SD' + str(sd)
            new_issue = gismu3lp_project.find_issue(sd)

            """Если обращения с таким sd нет - создаём новое, иначе - добавляем комментарий"""
            if new_issue is None:
                new_issue = gismu3lp_project.create_issue(name=current_issue[4], description=current_issue[5],
                    sd=sd, labels=current_issue[1], reproduce_type=current_issue[3],
                    region=current_issue[2])
                result = pd.concat([result, gismu3lp_project.issue_to_pd(new_issue, 'Новая')])
            else:
                gismu3lp_project.add_comment(new_issue.key, current_issue[5])
                result = pd.concat([result, gismu3lp_project.issue_to_pd(new_issue, 'Комментарий')])
        except Exception as exception:
            issue_except = GISMUException(["Ошибка при обработке обращения " + sd, exception])
            issue_except.print_to_file(path + '\\Ошибки.txt')

    """Если есть вложения - прикладываем их"""
    if array_of_dir.__len__() > 0:
        for current_dir in array_of_dir:
            try:
                print(gismu3lp_project.add_attachments(current_dir))
            except Exception as exception:
                attachment_except = GISMUException("Ошибка при обработке вложения " + current_dir.name, exception)
                attachment_except.print_to_file(path + '\\Ошибки.txt')

    """Создаём файл с результатами"""
    result.to_excel(path + '\\Результат.xlsx', index=False)

except Exception as exception:
    main_except = GISMUException("Ошибка в main файле ", exception)
    main_except.print_to_file(path + '\\Ошибки.txt')
