import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

def col_letter_to_num(col_letter):
    num = 0
    for char in col_letter:
        num = num * 26 + (ord(char.upper()) - ord('A') + 1)
    return num

def col_num_to_letter(col_num):
    letter = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(remainder + ord('A')) + letter
    return letter

scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key('16QakQHtFK7TTm5c8-bDBHqSul7BYRNzERqPy6xpYK7M')
worksheet = spreadsheet.worksheet('test_api')

spawn_col = "C"
spawn_row = 10
current_date = datetime.now().strftime("%Y-%m-%d")
program_name = "Programme 1"
workout_data = {
    'pompes': [15, 16, 16, 12],
    'tractions': [18, 17, 15],
    'dips': [10, 10, 10, 11, 12, 13],
    'gainage': [15, 30]
}

headers = [program_name] + [exercise.capitalize() for exercise in workout_data.keys()]
max_length = max(len(v) for v in workout_data.values())
data = [headers] + [[f"série {i + 1}"] + [workout_data[exercise][i] if i < len(workout_data[exercise]) else "" for exercise in workout_data] for i in range(max_length)]

start_col_index = col_letter_to_num(spawn_col)
end_col_index = start_col_index + len(headers) - 1
start_row_index = spawn_row - 1
end_row_index = start_row_index + max_length  

all_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index,
    "endRowIndex": end_row_index + 1,
    "startColumnIndex": start_col_index - 1,
    "endColumnIndex": end_col_index + 1
}
all_range_color = {"red": 56.0, "green": 90.0, "blue": 10.0, "alpha": 1}

date_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index,
    "endRowIndex": end_row_index + 1,
    "startColumnIndex": start_col_index - 1,
    "endColumnIndex": start_col_index
}
date_range_color = {"red": 56.0, "green": 90.0, "blue": 90.0, "alpha": 1}

cols_headers_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index,
    "endRowIndex": start_row_index + 1,
    "startColumnIndex": start_col_index + 1,
    "endColumnIndex": end_col_index + 1
}
cols_headers_range_color = {"red": 56.0, "green": 90.0, "blue": 10.0, "alpha": 1}

rows_headers_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index + 1,
    "endRowIndex": end_row_index + 1,
    "startColumnIndex": start_col_index,
    "endColumnIndex": start_col_index + 1
}
rows_headers_range_color = {"red": 5.0, "green": 70.0, "blue": 80.0, "alpha": 1}

program_name_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index,
    "endRowIndex": start_row_index + 1,
    "startColumnIndex": start_col_index,
    "endColumnIndex": start_col_index + 1
}
program_name_range_color = {"red": 56.0, "green": 10.0, "blue": 48.0, "alpha": 1}

table_range = {
    'sheetId': worksheet._properties['sheetId'],
    'startRowIndex': spawn_row - 1,
    'endRowIndex': end_row_index + 1,
    'startColumnIndex': start_col_index,
    'endColumnIndex': end_col_index + 1
}
program_name_range_color = {"red": 40.0, "green": 90.0, "blue": 62.0, "alpha": 1}

table_value_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index + 1,
    "endRowIndex": end_row_index + 1,
    "startColumnIndex": start_col_index + 1,
    "endColumnIndex": end_col_index + 1
}
table_value_range_color = {"red": 90.0, "green": 10.0, "blue": 10.0, "alpha": 1}

ranges_and_colors = [
    (all_range, all_range_color),
    (date_range, date_range_color),
    (cols_headers_range, cols_headers_range_color),
    (rows_headers_range, rows_headers_range_color),
    (program_name_range, program_name_range_color),
    (table_value_range, table_value_range_color)
]


static_requests = [
    {
        "mergeCells": {
            "mergeType": "MERGE_COLUMNS",
            "range": date_range
        }
    },
    {
        'updateCells': {
            'rows': [{'values': [{'userEnteredValue': {'stringValue': str(cell)}} for cell in row]} for row in data],
            'fields': 'userEnteredValue',
            'range': table_range
        },
    },
    {
        'updateCells': {
            'rows': [{'values': [{'userEnteredValue': {'stringValue': current_date}}]}],
            'fields': 'userEnteredValue',
            'range': date_range
        }
    },
    {
        "repeatCell": {
            "range": date_range,
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE"
                }
            },
            "fields": "userEnteredFormat(horizontalAlignment, verticalAlignment)"
        }
    },
    {
        "repeatCell": {
            "range": table_range,
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER"
                }
            },
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    }
]

color_requests = [
    {
        "repeatCell": {
            "range": range_info,
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": color_info
                }
            },
            "fields": "userEnteredFormat.backgroundColor"
        }
    } for range_info, color_info in ranges_and_colors
]

requests = static_requests + color_requests

formatted_json = json.dumps(requests, indent=4)
print(formatted_json)

spreadsheet.batch_update({'requests': requests})
