import sys
from   datetime import datetime
import base64
import requests
import json
from   pprint import pprint
import urllib.parse
from   zoneinfo import ZoneInfo
import os
import pytz
from datetime import timedelta 
import pyperclip

base_url = ''

# YYYY-MM-ddTHH:MM:ss.sss+0200
def convert_time(dt, to_tz):
    fm_dt = datetime.fromisoformat(dt)
    tz2 = pytz.timezone(to_tz)
    to_dt = fm_dt.astimezone(tz2)
    return to_dt

def display_done():
    pprint('done')

def url_encode(url):
    return urllib.parse.quote(url)

def json_bool_convertor(bool_val):
    if(str(bool_val).upper() == 'TRUE'):
        return True
    else:
        return False

# Please input full string in para_val, i.e.: state: 'active', jql: 'STATUS = "In Progress" AND issueType != "Sub-task"
def add_get_paras(paras: dict, para_name, para_val):
    paras[para_name] = url_encode(para_val)

def prepare_jira_agile_authentication(is_test):
    user_ac = ''
    token = ''

    base64_token_byte = base64.b64encode(bytes(user_ac+':'+token, 'utf-8')) # bytes
    base64_token_str = base64_token_byte.decode('utf-8') # convert bytes to string

    auth_str = 'Basic '+ user_ac + ':' + token
    headers = {'Authorization': 'Basic ' + base64_token_str}
    return headers

def get_Jira_api_full_response(base_url, headers, paras: dict, obj_key):

    results = ()

    isLast = False
    startAt = 0

    while(not isLast):
        url = base_url + '?startAt=' + str(startAt)
        
        for key, val in paras.items():
            url += f'&{key}={val}'
        
        response = requests.get(url, headers=headers)
        if(response.status_code == 200):
            current_json = response.json()
            if('isLast' in current_json):
                isLast = json_bool_convertor(current_json['isLast'])
                startAt = startAt + current_json['maxResults']
            else:
                isLast = True
        else:
            isLast = True
            break
        
        for val in current_json[obj_key]:
            results += (val, )
    
    return results

def get_active_boards(headers):
    return get_Jira_api_full_response(f'{base_url}/rest/agile/1.0/board', headers, {}, 'values')

def get_active_sprint_per_board(headers, board_id):
    paras = {}
    add_get_paras(paras, 'state', 'active')

    url = f'{base_url}/rest/agile/1.0/board/{board_id}/sprint'
    return get_Jira_api_full_response(url, headers, paras, 'values')

def get_open_issues_per_sprint(headers, sprint_id):
    paras = {}
    # add_get_paras(paras, 'jql', 'status!="Closed" AND status!="Verified"')
    add_get_paras(paras, 'jql', 'status!="Closed"')

    url = f'{base_url}/rest/agile/1.0/sprint/{str(sprint_id)}/issue'
    return get_Jira_api_full_response(url, headers, paras, 'issues')
    
def validate_date_format(date):
    format = '%Y-%m-%d'
    try:
        res = bool(datetime.strptime(date, format))
    except ValueError:
        res = False
    
    return res

def validate_date_format_2(date):
    format = '%Y%m%d'
    try:
        res = bool(datetime.strptime(date, format))
    except ValueError:
        res = False
    
    return res

def convert_str_to_int(input):
    try:
        ret = int(input)
    except ValueError:
        ret = ''
    return ret

def create_new_sprint(headers, sprint_json):

    headers_loc = headers
    headers_loc['Content-type'] = 'application/json'

    url = f'{base_url}/rest/agile/1.0/sprint'
    
    response = requests.post(url, headers=headers_loc, data=sprint_json)
    if(response.status_code == 201):
        response_json = response.json()
        return int(response_json['id'])
    else:
        response_json = response.json()
        if('errors' in response_json):
            for error in response_json['errors']:
                print(error)
        elif('errorMessages' in response_json):
            print(response_json['errorMessages'])
        return 0
    
def update_sprint(headers, sprint_id, sprint_json):

    headers_loc = headers
    headers_loc['Content-type'] = 'application/json'

    url = f'{base_url}/rest/agile/1.0/sprint/{sprint_id}'
    
    response = requests.post(url, headers=headers_loc, data=sprint_json)
    if(response.status_code == 200):
        response_json = response.json()
        return int(response_json['id'])
    else:
        response_json = response.json()
        if('errors' in response_json):
            for error in response_json['errors']:
                print(error)
        elif('errorMessages' in response_json):
            print(response_json['errorMessages'])
        return 0

def move_issues_to_sprint(headers, sprint_id, issues: list):

    headers_loc = headers
    headers_loc['Content-type'] = 'application/json'

    moved_issues = []

    url = f'{base_url}/rest/agile/1.0/sprint/{sprint_id}/issue'
    issue_cnt = 0
    issue_list_json = {}
    issue_list_json['issues'] = []

    for issue in issues:
        issue_list_json['issues'].append(issue['key'])
        issue_cnt += 1

        if (issue_cnt == 50):
            response = requests.post(url, headers=headers_loc, data=json.dumps(issue_list_json))
            if(response.status_code == 204):
                for val in issue_list_json['issues']:
                    moved_issues.append(val)
                issue_list_json['issues'] = []
                issue_cnt = 0

    if(issue_cnt > 0):
        response = requests.post(url, headers=headers, data=json.dumps(issue_list_json))
        if(response.status_code == 204):
            for val in issue_list_json['issues']:
                moved_issues.append(val)
            issue_list_json['issues'] = []
            issue_cnt = 0

    return moved_issues

###############################################################


is_test = True

if(is_test):
    base_url = ''
else:
    base_url = ''

print('Sprint Completion Program')
# move only un-closed or un-verified issue to next sprint

auth_headers = prepare_jira_agile_authentication(is_test)
boards = get_active_boards(auth_headers)

board_dict = {}

for board in boards:
    if(board['type'] == 'scrum'):
        board_dict[board['id']] = board['name']

if(len(board_dict) == 0):
    print('No scrum board available')
    print('Press any key to continue...')
    input()
    quit()

while(True):

    # Board
    print('Press "Ctrl + C" to end program at any step, or press "E" to exit current step')

    for key, val in board_dict.items():
        print(f'{key}: {val}')

    board_key = input('Enter a board id: ')
    if(str(board_key).upper() == 'E'):
        break

    board_key = convert_str_to_int(board_key)

    if(board_key not in board_dict.keys()):
        print('Invalid board id')
        print()
        continue

    # Sprint
    sprint_dict = {}
    sprints = get_active_sprint_per_board(auth_headers, board_key)
    for sprint in sprints:
        sprint_dict[sprint['id']] = sprint['name']

    if(len(sprint_dict) == 0):
        print('No active sprint available')
        print()
        continue

    while(True):
        for key, val in sprint_dict.items():
            print(f'\t{key}: {val}')

        if(len(sprint_dict) == 0):
            print('No active sprint available')
        
        sprint_key = input('Enter an active sprint id: ')

        if(str(sprint_key).upper() == 'E'):
            break

        sprint_key = convert_str_to_int(sprint_key)

        if(sprint_key not in sprint_dict.keys()):
            print('Invalid sprint id')
            print()
            continue    

        while(True):
            pyperclip.copy(sprint_dict[sprint_key])
            new_sprint_name = input(f'Rename "{sprint_dict[sprint_key]}" as (Press ctrl+v to paste original sprint name): ')
            if(new_sprint_name.upper() == 'E'):
                new_sprint_name = ''
                break
            elif(new_sprint_name == '' or new_sprint_name == sprint_dict[sprint_key]):
                print('Invalid new sprint name')
                print()
                continue
            else:
                break
        
        if(new_sprint_name == ''):
            continue

        while(True):
            utc0_timezone = pytz.timezone('Africa/Abidjan')
            hk_timezone = pytz.timezone('Asia/Hong_Kong')

            new_start_date = input('Input sprint start date (YYYYMMdd) (Leave empty to be today date): ')
            if(str(new_start_date).upper() == 'E'):
                new_sprint_name = ''
                break

            if(new_start_date == ''):
                new_start_date = datetime.now()
                new_start_date = new_start_date.replace(hour=0, minute=0, second=0, microsecond=0)
                new_start_date = utc0_timezone.localize(new_start_date)

            if(isinstance(new_start_date, str)):
                if(validate_date_format_2(new_start_date)):
                    new_start_date = datetime.strptime(new_start_date, '%Y%m%d')
                    new_start_date = utc0_timezone.localize(new_start_date)
                else:
                    continue

            new_end_date = input('Leave empty -> start date + 14, Press "2" -> start date +28, else (YYYYMMdd): ')
            if(str(new_end_date).upper() == 'E'):
                new_sprint_name = ''
                break

            if(str(new_end_date) == ''):
                new_end_date = new_start_date + timedelta(days=14)
                break
            elif(str(new_end_date) == '2'):
                new_end_date = new_start_date + + timedelta(days=28)
                break
            else:
                if(isinstance(new_end_date, str)):
                    if(validate_date_format_2(new_end_date)):
                        new_end_date = datetime.strptime(new_end_date, '%Y%m%d')
                        new_end_date = utc0_timezone.localize(new_end_date)
                        break
                    else:
                        continue
        
        if(new_sprint_name == ''):
            continue

        new_sprint_json_raw_dict = {}
        new_sprint_json_raw_dict['name'] = new_sprint_name
        new_sprint_json_raw_dict['startDate'] = new_start_date.isoformat(timespec='milliseconds')
        new_sprint_json_raw_dict['endDate'] = new_end_date.isoformat(timespec='milliseconds')
        new_sprint_json_raw_dict['originBoardId'] = int(board_key)
        
        new_sprint_json = json.dumps(new_sprint_json_raw_dict)

        new_sprint = 0
        open_issues_in_old_sprint = get_open_issues_per_sprint(auth_headers, sprint_key)
        if(len(open_issues_in_old_sprint) > 0):
            new_sprint = create_new_sprint(auth_headers, new_sprint_json)
        if(new_sprint > 0 or len(open_issues_in_old_sprint) == 0):
            complete_sprint_json_raw_dict = {}
            complete_sprint_json_raw_dict['state'] = 'closed'
            complete_sprint_json_raw_dict['completeDate'] = datetime.now(hk_timezone).isoformat(timespec='milliseconds')
            
            complete_sprint_json = json.dumps(complete_sprint_json_raw_dict)

            if(update_sprint(auth_headers, sprint_key, complete_sprint_json) > 0):
                if(new_sprint > 0):
                    moved_issue_cnt = move_issues_to_sprint(auth_headers, new_sprint, open_issues_in_old_sprint)
                    if(len(moved_issue_cnt) == len(open_issues_in_old_sprint)):
                        sprint_dict.pop(sprint_key)

                        start_sprint_json_raw_dict = {}
                        start_sprint_json_raw_dict['state'] = 'active'

                        start_sprint_json = json.dumps(start_sprint_json_raw_dict)
                        if(update_sprint(auth_headers, new_sprint, start_sprint_json) == 0):
                            print(f'Error starting sprint {new_sprint_name}')
                            quit()
                elif(len(open_issues_in_old_sprint) == 0):
                    print('No new sprint required')
                    
                else:
                    print('Error moving issue')
                    quit()

    print()


display_done()
print('Press any key to continue...')
input()