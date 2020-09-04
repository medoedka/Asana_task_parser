import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials

headers = {
    'Accept': 'application/json',
    'Authorization': 'Bearer {your_bearer_token}'
}

params = (
    ('workspace', 'your_workspace_id'),
    ('owner', 'your_owner_id'),
)

# connecting to worksheet
credentials = ServiceAccountCredentials.from_json_keyfile_name("C:\\Users\\User1\\Downloads\\exchange_rates.json")
googlesheet = gspread.authorize(credentials)
worksheet = googlesheet.open('Курсы валют').sheet1

response = requests.get('https://app.asana.com/api/1.0/portfolios', headers=headers, params=params)
primary_key = 1
for project_info in response.json()['data']:
    portfolio_id = project_info['gid']
    response_projects = requests.get(f'https://app.asana.com/api/1.0/portfolios/{portfolio_id}/items',
                                     headers=headers, params=params)
    projects = response_projects.json()['data']
    for project in projects:
        project_id = project['gid']
        project_tasks = requests.get(f'https://app.asana.com/api/1.0/projects/{project_id}/tasks', headers=headers)
        tasks = project_tasks.json()['data']
        for task in tasks:
            try:
                task_id = task['gid']
                task_data = requests.get(f'https://app.asana.com/api/1.0/tasks/{task_id}', headers=headers)
                task_info = task_data.json()['data']
                task_name = task['name'] if task['name'] is not None else ' '
                project_name = project['name'] if project['name'] is not None else ' '
                start = task_info['created_at'] if task_info['created_at'] is not None else ' '
                end = task_info['due_on'] if task_info['due_on'] is not None else ' '
                completion = task_info['completed_at'] if task_info['completed_at'] is not None else ' '
            except KeyError:
                start = ' '
                end = ' '
                completion = ' '
            try:
                assignee_name = task_info['assignee']['name'] if task_info['assignee']['name'] is not None else ' '
            except TypeError:
                assignee_name = ' '
            except KeyError:
                assignee_name = ' '
            try:
                payment = task_info['custom_fields'][1]['number_value'] if task_info['custom_fields'][1]['number_value'] is not None else ' '
            except TypeError:
                payment = ' '
            except KeyError:
                payment = ' '
            except IndexError:
                pass
            try:
                time_for_task = task_info['custom_fields'][0]['number_value'] if task_info['custom_fields'][0]['number_value'] is not None else ' '
            except TypeError:
                time_for_task = ' '
            except KeyError:
                time_for_task = ' '
            except IndexError:
                pass
            section_name = task_info['memberships'][0]['section']['name']
            worksheet.insert_row([primary_key, task_name, project_name, start[:10], end, completion[:10], assignee_name, payment, time_for_task, section_name, task_id], primary_key)
            primary_key += 1
