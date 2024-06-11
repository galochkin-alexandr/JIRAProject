from pathlib import Path

from Project import *
from GISMUException import *
import pandas as pd
from settings import get_settings

"""Достаём конфиги"""
settings = get_settings()
server_name = settings.server_name
project_name = settings.project_name
token_auth = settings.token_auth

gismu3lp_project = Project(server_name, project_name, token_auth)

# iss = gismu3lp_project.get_issue('GISMU3LP-23420')

new_issues = gismu3lp_project.get_current_day_issues()
# new_issues = gismu3lp_project.get_prev_day_issues()

result = pd.DataFrame({'Номер Обращения': [], 'Дата регистрации': [],
         'Плановая дата выполнения (для обращений "в работе")': [],
         'ФИО Заявителя': [],
         'Подразделение заявителя': [], 'Подробное описание запроса': [],
         'В рамках чьего поручения (если указано в тексте обращения)': [], 'Приоритет': [],
         'Категория (2 или 3)': [],
         'Статус выполнения (одновременно не может быть более одной задачи в статусе "в Работе")': [],
         'Комментарий Сервисной организации': [],
         'В случае изменения приоритета, указать по согласованию с Кем и когда выполнено изменение приоритета': [],
         'Ссылка на обращение': []})

for current_key in new_issues:
    print(current_key.key)
    issue = gismu3lp_project.get_issue(current_key.key)
    sd = issue.fields.customfield_23497
    registration_dt = issue.fields.customfield_26999
    expected_end_dt = issue.fields.customfield_23515
    department = issue.fields.customfield_26998
    description = issue.fields.description
    fio = gismu3lp_project.get_fio(description)
    comment = gismu3lp_project.get_last_comment(issue).body
    status = issue.fields.status.name
    issue_pd = pd.DataFrame(
        {'Номер Обращения': [sd], 'Дата регистрации': [registration_dt],
         'Плановая дата выполнения (для обращений "в работе")': [expected_end_dt],
         'ФИО Заявителя': [fio],
         'Подразделение заявителя': [department], 'Подробное описание запроса': [description],
         'В рамках чьего поручения (если указано в тексте обращения)': [''], 'Приоритет': [''],
         'Категория (2 или 3)': [''],
         'Статус выполнения (одновременно не может быть более одной задачи в статусе "в Работе")': [status],
         'Комментарий Сервисной организации': [comment],
         'В случае изменения приоритета, указать по согласованию с Кем и когда выполнено изменение приоритета': [''],
         'Ссылка на обращение': ['https://jira.phoenixit.ru/browse/' + current_key.key]})
    result = pd.concat([result, issue_pd])

path = str(Path.home())

"""Создаём файл с результатами"""
with pd.ExcelWriter(path + '\\Desktop\\Обращения за текущий день.xlsx') as writer:
    result.to_excel(writer, index=False, sheet_name='Лист1')

print('Путь до файла: ' + path + '\\Desktop\\Обращения за текущий день.xlsx')


