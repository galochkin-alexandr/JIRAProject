from jira import JIRA
from GISMUException import *
import pandas as pd
import os
from datetime import datetime, timedelta


class Project:
    """Класс для проекта Jira"""

    jira = None
    jira_project = None

    def __init__(self, server_name, project_name, token_auth):
        self.jira = JIRA(options={'server': server_name}, token_auth=token_auth)
        self.jira_project = self.jira.project(project_name, expand=None)

    def find_issue(self, number_sd):
        """Достаёт имя задачи по её SD"""

        try:
            search_query = """ "Номер обращения АСУЭ" ~ """ + number_sd.upper()
            result = self.jira.search_issues(search_query)
            if result.__len__() > 0:
                return result[0]
            else:
                return None
        except Exception as find_issue_except:
            raise GISMUException("Ошибка при получении задачи по sd ", find_issue_except)

    def get_issue(self, issue_key):
        """Достаёт задачу по её имени (GISMU3LP-22197)"""

        try:
            issue = self.jira.issue(issue_key)
            return issue
        except Exception as get_issue_except:
            raise GISMUException("Ошибка при получении задачи по ключу" + issue_key, get_issue_except)

    def get_user(self, login):
        """Достаёт пользователя по его имени (agalochkin)"""

        try:
            jira_user = self.jira.user(login)
            result_user = {'displayName': jira_user.displayName, 'key': jira_user.key, 'name': jira_user.name}
        except Exception as get_user_except:
            raise GISMUException("Ошибка при получении пользователя " + login, get_user_except)
        return result_user

    @staticmethod
    def labels_and_assignee(labels):
        """Собирает исполнителя и области по строке меток через пробел (РП ЗП)"""
        try:
            if pd.isnull(labels):
                return {'labels': None, 'assignee': None}
            array_labels = labels.split(' ')
            result_labels = []
            assignee_login = None
            excel_labels = pd.read_excel("../resources/Области.xlsx", sheet_name=0).to_numpy()
            for current_labels in array_labels:
                for current_excel in excel_labels:
                    if (current_labels.upper() == str(current_excel[0]).upper() or
                        current_labels.upper() == str(current_excel[1]).upper()):
                        result_labels.append(current_excel[0])
                        if assignee_login is None:
                            assignee_login = current_excel[2]
                        break
            return {'labels': result_labels, 'assignee': assignee_login}
        except Exception as labels_except:
            str_labels = ', '.join(labels)
            raise GISMUException("Ошибка при сопоставлении меток и пользователей " + str_labels, labels_except)

    @staticmethod
    def match_region(region):
        """Сопоставляет региону из АСУЭ регион в Jira"""

        try:
            if pd.isnull(region):
                return None
            excel_regions = pd.read_excel("../resources/Регионы.xlsx", sheet_name=0).to_numpy()
            for current_region in excel_regions:
                if str(current_region[0]).upper() in region.upper():
                    return {'value': current_region[1], 'id': str(current_region[2])}
        except Exception as region_except:
            raise GISMUException("Ошибка при сопоставлении региона" + region, region_except)

    @staticmethod
    def reproducibility(reproduce_type):
        """Собирает поле 'Воспроизводится'
           '1' - у 1, '2' - в 100%, иначе - Иное"""

        try:
            if str(reproduce_type) == '1' or str(reproduce_type) == '1.0':
                reproduce = {'value': 'у 1 пользователя/АРМ', 'id': '21755'}
            elif str(reproduce_type) == '2' or str(reproduce_type) == '2.0':
                reproduce = {'value': 'в 100% случаев (у всех пользователей)', 'id': '21754'}
            else:
                reproduce = {'value': 'Иное', 'id': '21756'}
            return reproduce
        except Exception as reproduce_except:
            raise GISMUException("Ошибка при сопоставлении поля 'Воспроизводится' " + reproduce_type, reproduce_except)

    @staticmethod
    def get_prolong_date(prolong_date):
        """Преобразовывает дату для поля 'Ожидаемая дата решения'
            На вход строка в формате YYYY-MM-DD, YYYY.MM.DD или YYYY/MM/DD"""

        try:
            if pd.isnull(prolong_date) or prolong_date == '':
                return None
            return prolong_date.replace('/', '-').replace('.', '-')
        except Exception as prolong_except:
            raise GISMUException("Ошибка при преобразовании даты решения" + prolong_date, prolong_except)

    @staticmethod
    def get_registration_date(registration_date):
        """Преобразовывает дату для поля 'Регистрация в АСУЭ'
            На вход строка в формате YYYY-MM-DD, YYYY.MM.DD или YYYY/MM/DD"""

        try:
            if pd.isnull(registration_date) or registration_date == '':
                return None
            return registration_date.replace('/', '-').replace('.', '-')
        except Exception as registration_except:
            raise GISMUException("Ошибка при преобразовании даты регистрации в АСУЭ" + registration_date,
                registration_except)

    @staticmethod
    def get_department(department):
        """Подразделение пользователя"""
        try:
            if pd.isnull(department) or department is None:
                return None
            return str(department)
        except Exception as department_except:
            raise GISMUException("Ошибка при получении подразделения пользователя" + department, department_except)

    @staticmethod
    def category(category_type):
        """Собирает поле 'Категория' '1' - Консультация, '2' - Запрос на изменение,
        '3' - Выгрузка, '4' - Вопрос по качеству данных, иначе - Не определено, """

        try:
            if str(category_type) == '1' or str(category_type) == '1.0':
                category = {'id': '26977'}
            elif str(category_type) == '2' or str(category_type) == '2.0':
                category = {'id': '26978'}
            elif str(category_type) == '3' or str(category_type) == '3.0':
                category = {'id': '26979'}
            elif str(category_type) == '4' or str(category_type) == '4.0':
                category = {'id': '26980'}
            elif str(category_type) == '5' or str(category_type) == '5.0':
                category = {'id': '27001'}
            else:
                category = None
            return category
        except Exception as category_except:
            raise GISMUException("Ошибка при сопоставлении поля 'Категория' " + category_type, category_except)

    def create_issue(self, name, description, sd, labels, reproduce_type, region, category_type, prolong_date,
        department, registration_date):
        """Создание задачи в Jira"""

        print("Создание задачи " + sd)
        try:
            labels_and_assignee = self.labels_and_assignee(labels)
            reproduce = self.reproducibility(reproduce_type)
            region = self.match_region(region)
            category = self.category(category_type)
            prolong_date = self.get_prolong_date(prolong_date)
            department = self.get_department(department)
            registration_date = self.get_registration_date(registration_date)
            issue_dict = {
                'project': {'key': self.jira_project.key},
                'summary': name.replace('\n', ''),
                'description': description,
                'issuetype': {'name': 'Обращение'},
                "customfield_23496": reproduce,
                "components": [{'name': 'ЦПОиБА', "id": '27849'}],
                "customfield_23497": sd.upper(),
                "labels": labels_and_assignee['labels'],
                "customfield_23514": region,
                "customfield_26111": category,
                "customfield_13590": 'GISMU3LP-544',
                "customfield_23515": prolong_date,
                "customfield_26998": department,
                "customfield_26999": registration_date
            }
            new_issue = self.jira.create_issue(fields=issue_dict)
            self.jira.assign_issue(new_issue.key, labels_and_assignee['assignee'])
            return new_issue
        except Exception as create_except:
            raise GISMUException("Ошибка при создании задачи " + sd, create_except)

    def add_comment(self, issue_key, text, prolong_date):
        """К задаче по её имени (GISMU3LP-22197) добавляет комментарий text"""

        print("Комментарий к " + issue_key)
        try:
            comment = self.jira.add_comment(issue_key, text)
            prolong_date = self.get_prolong_date(prolong_date)
            if prolong_date is not None:
                issue = self.get_issue(issue_key)
                issue.update(fields={"customfield_23515": prolong_date})
            new_status = self.update_status(issue_key, text)
            if new_status is not None:
                print("Новый статус: " + new_status)
            return comment
        except Exception as comment_except:
            raise GISMUException("Ошибка при создании комментария " + issue_key, comment_except)

    def add_attachments(self, directory):
        """Прикрепляет к задаче все вложения из папки с названием SD... Возвращает pd"""

        issue_key = self.find_issue(directory.name)
        counter = 0
        if issue_key is not None:
            with os.scandir(directory.path) as all_attachments:
                for current_attachment in all_attachments:
                    with open(current_attachment.path, 'rb') as current_file:
                        self.jira.add_attachment(issue=issue_key, attachment=current_file)
                    counter += 1
            print({'Имя обращения': issue_key, 'Кол-во вложений': counter})
            return pd.DataFrame({'Имя обращения': [issue_key], 'Кол-во вложений': [counter]})
        else:
            raise GISMUException("Отсутствует обращение с sd " + directory.name)

    @staticmethod
    def issue_to_pd(issue, action):
        """Собирает поля issue в dataframe (для эксельки)"""

        if issue.fields.customfield_23514 is None:
            region = 'Не заполнено'
        else:
            region = issue.fields.customfield_23514.value

        if issue.fields.customfield_23496 is None:
            reproduce = 'Не заполнено'
        else:
            reproduce = issue.fields.customfield_23496.value

        if issue.fields.customfield_26111 is None:
            category = 'Не заполнено'
        else:
            category = issue.fields.customfield_26111.value

        if issue.fields.customfield_23515 is None:
            prolong_date = 'Не заполнено'
        else:
            prolong_date = issue.fields.customfield_23515

        issue_pd = pd.DataFrame(
            {'Имя': [issue.key], 'SD': [issue.fields.customfield_23497],
                'Метки': [' '.join(issue.fields.labels)],
                'Регион': [region],
                'Воспроизводится': [reproduce], 'Категория': [category],
                'Статус': [issue.fields.status.name],
                'Название': [issue.fields.summary], 'Ожидаемая дата решения': [prolong_date],
                'Подразделение пользователя': [issue.fields.customfield_26998],
                'Дата регистрации в АСУЭ': [issue.fields.customfield_26999],
                'Описание': [issue.fields.description], 'Действие': [action]})
        return issue_pd

    def get_unique_value_for_field(self, field_name, number_of_query):
        set_of_value = set()
        name = self.jira_project.name
        all_values = self.jira.search_issues('project=' + self.jira_project.key, maxResults=number_of_query,
            fields=field_name)
        for value in all_values:
            set_of_value.add(value)
        return set_of_value

    def update_status(self, issue_key, text):
        issue = self.get_issue(issue_key)
        if text.upper().startswith('ЗАКРЫТ'):
            self.jira.transition_issue(issue, 'Анализ')
            self.jira.transition_issue(issue, 'В работу')
            self.jira.transition_issue(issue, 'В ожидании подтверждения')
            self.jira.transition_issue(issue, 'Выполнено')
            return ('Выполнено')
        if text.upper().startswith('ОТВЕТ') and issue.fields.status.id == '19101':
            self.jira.transition_issue(issue, 'Анализ')
            return ('Анализ')
        if text.upper().startswith('ЗАПРОШ') and issue.fields.status.id != '19101':
            if issue.fields.status.id == '1' or issue.fields.status.id == '18096' or issue.fields.status.id == '19099':
                self.jira.transition_issue(issue, 'Анализ')
            self.jira.transition_issue(issue, 'Требует уточнения')
            return ('Требует уточнения')
        if text.upper().startswith('ОТКЛОН') and (
            issue.fields.status.id == '19099' or issue.fields.status.id == '18096'):
            self.jira.transition_issue(issue, 'Анализ')
            return ('Анализ')
        return (None)

    def get_current_day_issues(self):
        current_day = datetime.today().strftime('%Y-%m-%d')
        search_query = f""" 
            created > {current_day} 
            and project = 'GISMU3LP' 
            and "Ожидаемая дата решения" is not null
        """
        result = self.jira.search_issues(search_query, maxResults=500)
        return result

    def get_prev_day_issues(self):
        current_day = datetime.today().strftime('%Y-%m-%d')
        prev_day = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')
        search_query = f""" 
            created > {prev_day} and created < {current_day}
            and project = 'GISMU3LP' 
            and "Ожидаемая дата решения" is not null
        """
        result = self.jira.search_issues(search_query, maxResults=500)
        return result

    @staticmethod
    def get_fio(description):
        start_index = description.find('Контактные данные:')
        temp_desc = description[start_index + 18:]
        fio = temp_desc[:temp_desc.find('\n')]
        if fio[0] == ' ':
            fio = fio[1:]
        return fio

    @staticmethod
    def get_last_comment(issue):
        all_comments = issue.fields.comment.comments
        if all_comments is None or len(all_comments) == 0:
            return ''
        last_comment = sorted(all_comments, key=lambda comment: comment.created, reverse=True)[0]
        return last_comment

    @staticmethod
    def change_date_format(date, old_format, new_format):
        if date is None or date == '':
            return ''
        new_date = datetime.strptime(date, old_format).strftime(new_format)
        return str(new_date)
