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

def create_range(sheet_id, start_row, end_row, start_col, end_col):
    return {
        "sheetId": sheet_id,
        "startRowIndex": start_row,
        "endRowIndex": end_row,
        "startColumnIndex": start_col,
        "endColumnIndex": end_col
    }

def create_color_request(range_info, color_info):
    return {
        "repeatCell": {
            "range": range_info,
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": color_info
                }
            },
            "fields": "userEnteredFormat.backgroundColor"
        }
    }

def create_border_request(range_info):
    return {
        "updateBorders": {
            "range": range_info,
            "top": {
                "style": "SOLID",
                "width": 1,
                "color": {"red": 0, "green": 0, "blue": 0}
            },
            "bottom": {
                "style": "SOLID",
                "width": 1,
                "color": {"red": 0, "green": 0, "blue": 0}
            },
            "left": {
                "style": "SOLID",
                "width": 1,
                "color": {"red": 0, "green": 0, "blue": 0}
            },
            "right": {
                "style": "SOLID",
                "width": 1,
                "color": {"red": 0, "green": 0, "blue": 0}
            },
            "innerHorizontal": {
                "style": "SOLID",
                "width": 1,
                "color": {"red": 0, "green": 0, "blue": 0}
            },
            "innerVertical": {
                "style": "SOLID",
                "width": 1,
                "color": {"red": 0, "green": 0, "blue": 0}
            }
        }
    }

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

sheet_id = worksheet._properties['sheetId']

ranges_colors = {
    "all_range": (create_range(sheet_id, start_row_index, end_row_index + 1, start_col_index - 1, end_col_index + 1), {"red": 1.0, "green": 0.0, "blue": 0.0, "alpha": 1}),  # Rouge vif
    "date_range": (create_range(sheet_id, start_row_index, end_row_index + 1, start_col_index - 1, start_col_index), {"red": 0.0, "green": 1.0, "blue": 0.0, "alpha": 1}),  # Vert vif
    "cols_headers_range": (create_range(sheet_id, start_row_index, start_row_index + 1, start_col_index + 1, end_col_index + 1), {"red": 0.0, "green": 0.0, "blue": 1.0, "alpha": 1}),  # Bleu vif
    "rows_headers_range": (create_range(sheet_id, start_row_index + 1, end_row_index + 1, start_col_index, start_col_index + 1), {"red": 1.0, "green": 1.0, "blue": 0.0, "alpha": 1}),  # Jaune vif
    "table_range": (create_range(sheet_id, start_row_index, end_row_index + 1, start_col_index, end_col_index + 1), {"red": 0.0, "green": 1.0, "blue": 1.0, "alpha": 1}),  # Cyan vif
    "table_value_range": (create_range(sheet_id, start_row_index + 1, end_row_index + 1, start_col_index + 1, end_col_index + 1), {"red": 1.0, "green": 0.5, "blue": 0.0, "alpha": 1}),  # Orange vif
    "program_name_range": (create_range(sheet_id, start_row_index, start_row_index + 1, start_col_index, start_col_index + 1), {"red": 1.0, "green": 0.0, "blue": 1.0, "alpha": 1}),  # Magenta vif
}

static_requests = [
    {
        "mergeCells": {
            "mergeType": "MERGE_COLUMNS",
            "range": ranges_colors["date_range"][0]
        }
    },
    {
        'updateCells': {
            'rows': [{'values': [{'userEnteredValue': {'stringValue': str(cell)}} for cell in row]} for row in data],
            'fields': 'userEnteredValue',
            'range': ranges_colors["table_range"][0]
        },
    },
    {
        'updateCells': {
            'rows': [{'values': [{'userEnteredValue': {'stringValue': current_date}}]}],
            'fields': 'userEnteredValue',
            'range': ranges_colors["date_range"][0]
        }
    },
    {
        "repeatCell": {
            "range": ranges_colors["date_range"][0],
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
            "range": ranges_colors["all_range"][0],
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER"
                }
            },
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    }
]


color_requests = [create_color_request(rng, color) for rng, color in ranges_colors.values()]
border_requests = [create_border_request(ranges_colors["all_range"][0])]

requests = static_requests + color_requests + border_requests

formatted_json = json.dumps(requests, indent=4)
print(formatted_json)

spreadsheet.batch_update({'requests': requests})

