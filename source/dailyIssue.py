from pathlib import Path

from Project import *
import pandas as pd
from settings import get_settings
from GoogleService import *

"""Достаём конфиги"""
settings = get_settings()
server_name = settings.server_name
project_name = settings.project_name
token_auth = settings.token_auth
file_id = settings.excel_file_id

gismu3lp_project = Project(server_name, project_name, token_auth)

# google_service = GoogleService(path_to_google_credits)
new_issues = gismu3lp_project.get_current_day_issues()
# new_issues = gismu3lp_project.get_prev_day_issues()

result = pd.DataFrame({'SD': [], 'Задача в джире': [], 'Дата рег обращения в асуэ': [],
    'Ожидаемая дата решения': [],
    'Подразделение': [],
    'ФИО': [],
    'Текст обращения': [],
    'Последний комент в джире': [],
    'Приоритет': [],
    'Категория': [],
    'Статус': [],
    'Кто изменил приоритет': [],
    'Примечание': []})

start_line = 0
for current_key in new_issues:
    print(current_key.key)
    # if start_line == 0:
    #     start_line = google_service.get_index_of_first_empty_line(file_id)
    issue = gismu3lp_project.get_issue(current_key.key)
    sd = issue.fields.customfield_23497
    registration_dt = gismu3lp_project.change_date_format(issue.fields.customfield_26999, '%Y-%m-%d', '%d.%m.%Y')
    expected_end_dt = gismu3lp_project.change_date_format(issue.fields.customfield_23515, '%Y-%m-%d', '%d.%m.%Y')
    department = issue.fields.customfield_26998
    description = issue.fields.description
    fio = gismu3lp_project.get_fio(description)
    comment = gismu3lp_project.get_last_comment(issue)
    status = issue.fields.status.name
    array_to_google_excel = [sd, registration_dt, expected_end_dt, fio, department, description, '', '', '', '', comment]
    issue_pd = pd.DataFrame(
        {
            'SD': [sd],
            'Задача в джире': ['https://jira.phoenixit.ru/browse/' + current_key.key],
            'Дата рег обращения в асуэ': [registration_dt],
            'Ожидаемая дата решения': [expected_end_dt],
            'Подразделение': [department],
            'ФИО': [fio],
            'Текст обращения': [description],
            'Последний комент в джире': [comment],
            'Приоритет': [''],
            'Категория': [''],
            'Статус': [status],
            'В рамках чьего поручения (если указано в тексте обращения)': [''],
            'Кто изменил приоритет': [''],
            'Примечание': ['']
            })
    result = pd.concat([result, issue_pd])
    # if not google_service.check_condition_in_line(file_id, 'B', sd):
    #     google_service.print_in_excel_file(file_id, start_line, 'B', array_to_google_excel)
    #     start_line += 1

path = str(Path.home())

"""Создаём файл с результатами"""
with pd.ExcelWriter(path + '\\Desktop\\Обращения за текущий день.xlsx') as writer:
    result.to_excel(writer, index=False, sheet_name='Лист1')

print('Путь до файла: ' + path + '\\Desktop\\Обращения за текущий день.xlsx')
