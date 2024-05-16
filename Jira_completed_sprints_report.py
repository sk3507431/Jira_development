import sys
from   datetime import datetime
import base64
import requests
import json
from   pprint import pprint
import csv
import pandas as pd
import matplotlib.pyplot as plot
import openpyxl
import urllib.parse
import pytz
from   zoneinfo import ZoneInfo
import xlsxwriter
import os

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

def prepare_jira_agile_authentication():
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

def get_json_struct(json_struct, obj, curr_lv):
    obj_key_list = list(obj.keys())

    # for key in obj.keys():
    #     print(key + ': ' + str(type(obj[key])))

    for i in range(len(obj_key_list)):
        key = obj_key_list[i]
        if(isinstance(obj[key], dict)):
            get_json_struct(json_struct, obj[key], key)
        else:
            if(curr_lv != ''):
                json_struct.append(curr_lv + '.' + key)
            else:
                json_struct.append(key)

def get_json_value(key_chain, obj):
    keys = key_chain.split('.')
    curr_obj = obj
    for i in range(len(keys)):
        if(i < len(keys) - 1):
            curr_obj = curr_obj[keys[i]]
        else:
            return curr_obj[keys[i]]

def convert_contents_to_lines(headers, contents):
    lines = []
    for content in contents:
        curr_line = []
        for header in headers:
            curr_line.append(get_json_value(header, content))
        lines.append(curr_line)

    return lines

def get_active_boards(headers):
    return get_Jira_api_full_response(f'{base_url}/rest/agile/1.0/board', headers, {}, 'values')

def get_active_sprint_per_board(headers, board_id):
    paras = {}
    add_get_paras(paras, 'state', 'active')

    url = f'{base_url}/rest/agile/1.0/board/{board_id}/sprint'
    return get_Jira_api_full_response(url, headers, paras, 'values')

def get_closed_sprint_per_board(headers, board_id, complete_date):
    paras = {}
    add_get_paras(paras, 'state', 'closed')

    url = f'{base_url}/rest/agile/1.0/board/{board_id}/sprint'
    sprints = list(get_Jira_api_full_response(url, headers, paras, 'values'))
    for sprint in sprints.copy():
        if('completeDate' in sprint and sprint["completeDate"] != ''):
            sprint_complete_date = convert_time(sprint["completeDate"], 'Asia/Hong_Kong').date()
            if(sprint_complete_date < complete_date):
                sprints.remove(sprint)
        else:
            sprints.remove(sprint)

    return sprints


def get_main_tasks_per_sprint(headers, sprint_id):
    paras = {}
    add_get_paras(paras, 'fields', 'issuetype,summary,status')
    add_get_paras(paras, 'jql', 'issueType!="Sub-task"')
    

    url = f'{base_url}/rest/agile/1.0/sprint/{str(sprint_id)}/issue'
    return get_Jira_api_full_response(url, headers, paras, 'issues')

def get_active_issue_stat_line(sprint_name, issues):
    issue_stat_line = [sprint_name]

    issue_stats = dict.fromkeys(['Open','In Progress', 'Done', 'Total'], 0)
    for issue in issues:
        issue_status = issue['fields']['status']
        if(issue_status['statusCategory']['name'] == 'To Do' and issue_status['name'] != 'Verified'):
            # Open 
            issue_stats['Open'] += 1
        elif(issue_status['statusCategory']['name'] == "In Progress"):
            # In Progress
            issue_stats['In Progress'] += 1
        else:
            # Done
            issue_stats['Done'] += 1

    issue_stats['Total'] = issue_stats['Open'] + issue_stats['In Progress'] + issue_stats['Done']
    for val in issue_stats.values():
        issue_stat_line.append(val)
    
    return issue_stat_line

def get_issue_change_log(headers, issue_key):
    url = f'{base_url}/rest/agile/1.0/issue/{issue_key}?expand=changelog'
    issue_chg_log = ()

    response = requests.get(url, headers=headers)
    if(response.status_code == 200):
        current_json = response.json()
        issue_chg_log = (current_json['changelog']['histories'])
    
    return issue_chg_log
    
def get_issue_status_changelog(headers, issue_key):
    chg_log = get_issue_change_log(headers, issue_key)
    status_chg_log = []
    for history in chg_log:
        for item in history['items']:
            if(item['field'] == 'status'):
                status_chg_log.append(history)

    return status_chg_log

# update your own filtering criteria here
def get_issue_stat_history_line(headers, sprint_name, issues, sprint_complete_date):
    issue_stat_line = [sprint_name]

    issue_stats = dict.fromkeys(['Open','In Progress', 'Done', 'Total'], 0)
    for issue in issues:
        
        got_upon_status = False
        issue_status_change_history = get_issue_status_changelog(headers, issue['key'])
        for history in issue_status_change_history:
            for item in history['items']:
                if(item['field'] == 'status'):
                    chg_datetime = convert_time(history['created'], 'Asia/Hong_Kong')
                    upon_status = item['toString']
                    got_upon_status = True

            # Assupmtion: histories sort in decending order
            if(got_upon_status and chg_datetime < sprint_complete_date):
                break
        
        if(not got_upon_status):
            issue_stats['Open'] += 1
            continue

        upon_status = upon_status.upper()

        if(upon_status == 'CLOSED' or upon_status == 'RESOLVED' or 
           upon_status == 'UAT' or upon_status == 'VERIFIED'):
            # Done
            issue_stats['Done'] += 1
        elif(upon_status == 'IN PROGRESS' or upon_status == 'IN DEVELOPMENT' or 
             upon_status == 'REQUIREMENT COLLECTION' or upon_status == 'IMPACT ANALYSIS' or
             upon_status == 'INTERNAL TESTING IN PROGRESS'):
            # In Progress
            issue_stats['In Progress'] += 1
        else:
            # Open
            issue_stats['Open'] += 1

    issue_stats['Total'] = issue_stats['Open'] + issue_stats['In Progress'] + issue_stats['Done']
    for val in issue_stats.values():
        issue_stat_line.append(val)
    
    return issue_stat_line

def validate_date_format(date):
    format = '%Y-%m-%d'
    try:
        res = bool(datetime.strptime(date, format))
    except ValueError:
        res = False
    
    return res

def gen_active_sprint_stat(headers, board_name, board_lines, sprint):
    issues = get_main_tasks_per_sprint(headers, sprint['id'])
    issue_stat_line = get_active_issue_stat_line(board_name, issues)
    
    board_lines.append(issue_stat_line)

def gen_closed_sprint_stat(headers, board_name, board_lines, sprint):
    issues = get_main_tasks_per_sprint(headers, sprint['id'])
    issue_stat_line = get_issue_stat_history_line(headers, board_name, issues, convert_time(sprint["completeDate"], 'Asia/Hong_Kong'))
    
    board_lines.append(issue_stat_line)

def show_stacked_horizontal_barchart(title, issue_stat):
    data = dict.fromkeys(['Open', 'In Progress', 'Done'])
    for key in data.keys():
        data[key] = []
    index = []

    for stat in issue_stat:
        index.append(stat[0])   # Sprint name

        for i in list(range(1, 4)):
            if(i == 1): # Open
                data['Open'].append(stat[i])
            elif(i == 2): # In Progress
                data['In Progress'].append(stat[i])
            elif(i == 3): # Done
                data['Done'].append(stat[i])
    
    df = pd.DataFrame(data=data, index=index)
    df.plot.barh(stacked=True, title=title)
    # plot.show(block=True)

    return df

def plot_chart_in_excel(title, workbook, sheet_name, stat, start_row):
    row_cnt = len(stat)
    end_row = start_row + row_cnt 

    chart1 = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})
    
    # Open
    chart1.add_series({
        'name':       f'={sheet_name}!$C${str(start_row)}',
        'categories': f'={sheet_name}!$B${str(start_row + 1)}:$B${end_row}',
        'values':     f'={sheet_name}!$C${str(start_row + 1)}:$C${end_row}',
    })

    # In Progress"
    chart1.add_series({
        'name':       f'={sheet_name}!$D${str(start_row)}',
        'categories': f'={sheet_name}!$B${str(start_row + 1)}:$B${end_row}',
        'values':     f'={sheet_name}!$D${str(start_row + 1)}:$D${end_row}',
    })

    # Done"
    chart1.add_series({
        'name':       f'={sheet_name}!$E${str(start_row)}',
        'categories': f'={sheet_name}!$B${str(start_row + 1)}:$B${end_row}',
        'values':     f'={sheet_name}!$E${str(start_row + 1)}:$E${end_row}',
    })

    chart1.set_title({'name': title})

    chart1.set_y_axis({'reverse': True,
                      'crossing': 'max'})

    # chart1.set_style(11)
    
    return chart1

def adjust_excel_column(writer, sheetname, df, start_col):
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column) + start_col
        writer.sheets[sheetname].set_column(col_idx, col_idx, column_length)

###############################################################

filepath = ''
filepath = input('File path (leave empty to store under current directory): ')
if(filepath == ''):
    filepath = os.getcwd() + f'\\Sprint Reports {datetime.now().year}{datetime.now().month:02}.xlsx'

print(filepath)

complete_after_str = input('Get sprints completed since (YYYY-MM-dd): ')
if(not validate_date_format(str(complete_after_str))):
    exit()

complete_after_datetime = datetime.strptime(complete_after_str, '%Y-%m-%d').date()

auth_headers = prepare_jira_agile_authentication()
boards = get_active_boards(auth_headers)


with pd.ExcelWriter(filepath) as writer:
    workbook = writer.book

    for board in boards:

        chart_scale = {'x_scale': 1.2, 'y_scale': 0.9}

        if(board['type'] == 'scrum'):
            print(str(board['id']) + ': ' + board['name'])
            
            sheetname = board['name'].split(' ')[0]
            worksheet = workbook.add_worksheet(sheetname)

            ttl_data_row = 2
            
            issue_stat_in_closed_sprint = []
            issue_stat_in_active_sprint = []

            closed_sprints = get_closed_sprint_per_board(auth_headers, board['id'], complete_after_datetime)
            for sprint in closed_sprints:
                gen_closed_sprint_stat(auth_headers, sprint['name'], issue_stat_in_closed_sprint, sprint)
                issue_stat_in_closed_sprint = sorted(issue_stat_in_closed_sprint)
                closed_sprint_df = pd.DataFrame(issue_stat_in_closed_sprint, columns=['Completed Sprints', 'Open', 'In Progress', 'Done', 'Total'])
                closed_sprint_df.to_excel(writer, sheet_name=sheetname, startrow=ttl_data_row, startcol=1, header=True, index=False)

            
            if(len(issue_stat_in_closed_sprint) > 0):
                # closed_sprint_h_barchart_df = show_stacked_horizontal_barchart('Completed Sprints - Status', issue_stat_in_closed_sprint)
                closed_sprint_chart = plot_chart_in_excel('Completed Sprints - Status', workbook, sheetname, issue_stat_in_closed_sprint, ttl_data_row + 1)
                worksheet.insert_chart(f'B{ttl_data_row + len(issue_stat_in_closed_sprint) + 3}', closed_sprint_chart, chart_scale)

            ttl_data_row += len(issue_stat_in_closed_sprint)

            ttl_data_row += 3 + len(issue_stat_in_closed_sprint) * 4 + 2 # empty rows to hold chart

            active_sprints = get_active_sprint_per_board(auth_headers, str(board['id']))
            for sprint in active_sprints:
                gen_active_sprint_stat(auth_headers, sprint['name'], issue_stat_in_active_sprint, sprint)
                issue_stat_in_active_sprint = sorted(issue_stat_in_active_sprint)
                active_sprint_df = pd.DataFrame(issue_stat_in_active_sprint, columns=['Active Sprints', 'Open', 'In Progress', 'Done', 'Total'])
                active_sprint_df.to_excel(writer, sheet_name=sheetname, header=True, index=False, startrow=ttl_data_row, startcol=1)

            if(len(issue_stat_in_active_sprint) > 0):
                # active_sprint_h_barchart_df = show_stacked_horizontal_barchart('Active Sprints - Status', issue_stat_in_active_sprint)
                active_sprint_chart = plot_chart_in_excel('Active Sprints - Status', workbook, sheetname, issue_stat_in_active_sprint, ttl_data_row + 1)
                worksheet.insert_chart(f'B{ttl_data_row + len(issue_stat_in_active_sprint) + 3}', active_sprint_chart, chart_scale)

            ttl_data_row = ttl_data_row + len(issue_stat_in_active_sprint)

            adjust_excel_column(writer=writer, sheetname=sheetname, df=active_sprint_df, start_col=1)

display_done()