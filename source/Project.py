from jira import JIRA, JIRAError
import pandas as pd
import os


class Project:
    """Класс для проекта Jira"""

    jira = None
    jira_project = None

    def __init__(self, server_name, project_name, token_auth):
        self.jira = JIRA(options={'server': server_name}, token_auth=token_auth)
        self.jira_project = self.jira.project(project_name, expand=None)

    def find_issue(self, number_sd):
        """Достаёт имя задачи по её SD"""

        search_query = """ "Номер обращения АСУЭ" ~ """ + number_sd.upper()
        result = self.jira.search_issues(search_query)
        if result.__len__() > 0:
            return result[0]
        return None

    def get_issue(self, issue_key):
        """Достаёт задачу по её имени (GISMU3LP-22197)"""

        try:
            issue = self.jira.issue(issue_key)
        except JIRAError:
            issue = None
        return issue

    def get_user(self, login):
        """Достаёт пользователя по его имени (agalochkin)"""

        try:
            jira_user = self.jira.user(login)
            result_user = {'displayName': jira_user.displayName, 'key': jira_user.key, 'name': jira_user.name}
        except JIRAError:
            result_user = None
        return result_user

    def labels_and_assignee(self, labels):
        """Собирает исполнителя и области по строке меток через пробел (РП ЗП)"""

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

    @staticmethod
    def match_region(region):
        """Сопоставляет региону из АСУЭ регион в Jira"""

        if pd.isnull(region):
            return None
        excel_regions = pd.read_excel("../resources/Регионы.xlsx", sheet_name=0).to_numpy()
        for current_region in excel_regions:
            if str(current_region[0]).upper() in region.upper():
                return {'value': current_region[1], 'id': str(current_region[2])}
        return None

    @staticmethod
    def reproducibility(reproduce_type):
        """Собирает поле 'Воспроизводится'
           '' - Иное, '1' - у 1, иначе - в 100%"""

        if str(reproduce_type) == '1':
            reproduce = {'value': 'у 1 пользователя/АРМ', 'id': '21755'}
        elif str(reproduce_type) == '2':
            reproduce = {'value': 'в 100% случаев (у всех пользователей)', 'id': '21754'}
        else:
            reproduce = {'value': 'Иное', 'id': '21756'}
        return reproduce

    def create_issue(self, name, description, sd, labels, reproduce_type, region):
        labels_and_assignee = self.labels_and_assignee(labels)
        reproduce = self.reproducibility(reproduce_type)
        region = self.match_region(region)
        issue_dict = {
            'project': {'key': self.jira_project.key},
            'summary': name,
            'description': description,
            'issuetype': {'name': 'Обращение'},
            "customfield_23496": reproduce,
            "components": [{'name': 'ЦПОиБА', "id": '27849'}],
            "customfield_23497": sd.upper(),
            "labels": labels_and_assignee['labels'],
            "customfield_23514": region
        }
        new_issue = self.jira.create_issue(fields=issue_dict)
        self.jira.assign_issue(new_issue.key, labels_and_assignee['assignee'])
        return new_issue

    def add_comment(self, issue_key, text):
        """К задаче по её имени (GISMU3LP-22197) добавляет комментарий text"""

        comment = self.jira.add_comment(issue_key, text)
        issue = self.get_issue(issue_key)
        self.jira.transition_issue(issue, 'Анализ')
        return comment

    def add_attachments(self, directory):
        """Прикрепляет к задаче все вложения из папки с названием SD... Возвращает количество"""

        issue_key = self.find_issue(directory.name)
        counter = 0
        if issue_key is not None:
            with os.scandir(directory.path) as all_attachments:
                for current_attachment in all_attachments:
                    with open(current_attachment.path, 'rb') as current_file:
                        self.jira.add_attachment(issue=issue_key, attachment=current_file)
                    counter += 1
        return {'issue_key': issue_key, 'number_of_attachments': counter}

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

        issue_pd = pd.DataFrame(
            {'Имя': [issue.key], 'SD': [issue.fields.customfield_23497],
                'Метки': [' '.join(issue.fields.labels)],
                'Регион': [region],
                'Воспроизводится': [reproduce],
                'Название': [issue.fields.summary], 'Описание': [issue.fields.description], 'Действие': [action]})
        return issue_pd

    def get_unique_value_for_field(self, field_name, number_of_query):
        set_of_value = set()
        name = self.jira_project.name
        all_values = self.jira.search_issues('project=' + self.jira_project.key, maxResults=number_of_query,
            fields=field_name)
        for value in all_values:
            set_of_value.add(value)
        return set_of_value
