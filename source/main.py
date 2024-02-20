from tkinter.filedialog import *
from Project import *
import pandas as pd
from settings import get_settings

"""Достаём конфиги"""
settings = get_settings()
server_name = settings.server_name
project_name = settings.project_name
token_auth = settings.token_auth

gismu3lp_project = Project(server_name, project_name, token_auth)

"""Выбираем папки"""
path = askdirectory()
array_of_dir = []
main_file = ''

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

"""Если есть вложения - прикладываем их"""
if array_of_dir.__len__() > 0:
    for current_dir in array_of_dir:
        print(gismu3lp_project.add_attachments(current_dir))

"""Создаём файл с результатами"""
result.to_excel(path + '\\Результат.xlsx', index=False)
