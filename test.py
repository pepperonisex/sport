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

spawn_col = "A"
spawn_row = 1
current_date = datetime.now().strftime("%Y-%m-%d")  # Format date as needed
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
    "endRowIndex": end_row_index,
    "startColumnIndex": start_col_index - 1,
    "endColumnIndex": end_col_index
}

date_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index,
    "endRowIndex": start_row_index + 1,
    "startColumnIndex": start_col_index - 1,
    "endColumnIndex": start_col_index
}

cols_headers_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index,
    "endRowIndex": start_row_index + 1,
    "startColumnIndex": start_col_index,
    "endColumnIndex": end_col_index
}

rows_headers_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index + 1,
    "endRowIndex": end_row_index,
    "startColumnIndex": start_col_index - 1,
    "endColumnIndex": start_col_index
}

program_name_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index,
    "endRowIndex": start_row_index + 1,
    "startColumnIndex": start_col_index,
    "endColumnIndex": start_col_index + 1
}

table_range = {
    "sheetId": worksheet._properties['sheetId'],
    "startRowIndex": start_row_index + 1,
    "endRowIndex": end_row_index,
    "startColumnIndex": start_col_index,
    "endColumnIndex": end_col_index
}


requests = [
    {
        "mergeCells": {
            "mergeType": "MERGE_COLUMNS",
            "range": {
                "sheetId": worksheet._properties['sheetId'],
                "startRowIndex": spawn_row - 1,
                "endRowIndex": spawn_row + len(data) -1,
                "startColumnIndex": start_col_index - 1,
                "endColumnIndex": start_col_index
            }
        }
    },
    {
        'updateCells': {
            'rows': [{'values': [{'userEnteredValue': {'stringValue': str(cell)}} for cell in row]} for row in data],
            'fields': 'userEnteredValue',
            'range': {
                'sheetId': worksheet._properties['sheetId'],
                'startRowIndex': spawn_row - 1,
                'endRowIndex': spawn_row + len(data) - 1,
                'startColumnIndex': start_col_index,
                'endColumnIndex': end_col_index + 1
            }
        },
    },
    {
        'updateCells': {
            'rows': [{'values': [{'userEnteredValue': {'stringValue': current_date}}]}],
            'fields': 'userEnteredValue',
            'range': {
                'sheetId': worksheet._properties['sheetId'],
                'startRowIndex': spawn_row - 1,
                'startColumnIndex': start_col_index - 1,
                'endRowIndex': spawn_row + 1,
                'endColumnIndex': start_col_index
            }
        }
    },
    {
        "repeatCell": {
            "range": {
                "sheetId": worksheet._properties['sheetId'],
                "startRowIndex": spawn_row - 1,
                "endRowIndex": spawn_row,
                "startColumnIndex": start_col_index - 1,
                "endColumnIndex": start_col_index
            },
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
            "range": {
                "sheetId": worksheet._properties['sheetId'],
                "startRowIndex": spawn_row - 1,
                "endRowIndex": spawn_row + len(data) - 1,
                "startColumnIndex": start_col_index,
                "endColumnIndex": end_col_index + 1
            },
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER"
                }
            },
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    }
]

formatted_json = json.dumps(requests, indent=4)
print(formatted_json)
spreadsheet.batch_update({'requests': requests})
