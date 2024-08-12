import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import config
import locale

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

def update_spreadsheet(spreadsheet : gspread.Spreadsheet, current_date, program_name, workout_data, commentaire):
    
    original_locale = locale.getlocale(locale.LC_TIME)
    locale.setlocale(locale.LC_TIME, 'fr_FR')
    current_date2 = datetime.now()
    worksheet_name = current_date2.strftime("%B_%Y")

    worksheet_exists = False
    for sheet in spreadsheet.worksheets():
        if sheet.title == worksheet_name:
            worksheet_exists = True
            worksheet = sheet 
            break

    if not worksheet_exists:
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="1000", cols="26")
        print(f"Nouvelle worksheet créée: {worksheet_name}")
    else:
        print(f"Une worksheet avec le nom {worksheet.title} existe déjà.")

    locale.setlocale(locale.LC_TIME, original_locale)
    
    headers = [program_name] + [exercise.capitalize() for exercise in workout_data.keys()]
    max_length = max(len(v) for v in workout_data.values())
    data = [headers] + [[f"série {i + 1}"] + [workout_data[exercise][i] if i < len(workout_data[exercise]) else "" for exercise in workout_data] for i in range(max_length)]

    first_spawn_col = "B"
    first_spawn_row = 3

    start_col_index = col_letter_to_num(first_spawn_col)
    end_col_index = start_col_index + len(headers) - 1
    start_row_index = first_spawn_row - 1
    end_row_index = start_row_index + max_length  


    offset_between_tables = 2
    
    col_values = worksheet.col_values(start_col_index)
    first_spawn_row = col_values[-1].row_value + offset_between_tables
    print(first_spawn_row)
    sheet_id = worksheet._properties['sheetId']

    ranges_colors = {
        "all_range": (create_range(sheet_id, start_row_index, end_row_index + 1, start_col_index - 1, end_col_index + 2), {"red": 0.98, "green": 0.98, "blue": 0.98, "alpha": 1}),              # Presque blanc, léger gris
        "date_range": (create_range(sheet_id, start_row_index, end_row_index + 1, start_col_index - 1, start_col_index), {"red": 0.85, "green": 0.93, "blue": 0.97, "alpha": 1}),               # Bleu clair doux
        "cols_headers_range": (create_range(sheet_id, start_row_index, start_row_index + 1, start_col_index + 1, end_col_index + 1), {"red": 0.9, "green": 0.9, "blue": 0.9, "alpha": 1}),      # Gris clair
        "rows_headers_range": (create_range(sheet_id, start_row_index + 1, end_row_index + 1, start_col_index, start_col_index + 1), {"red": 0.95, "green": 0.95, "blue": 0.95, "alpha": 1}),   # Gris très léger
        "table_range": (create_range(sheet_id, start_row_index, end_row_index + 1, start_col_index, end_col_index + 1), {"red": 0.98, "green": 0.96, "blue": 0.89, "alpha": 1}),                # Blanc pur
        "table_value_range": (create_range(sheet_id, start_row_index + 1, end_row_index + 1, start_col_index + 1, end_col_index + 1), {"red": 1.0, "green": 1.0, "blue": 1.0, "alpha": 1}),     # Beige très clair
        "program_name_range": (create_range(sheet_id, start_row_index, start_row_index + 1, start_col_index, start_col_index + 1), {"red": 0.93, "green": 0.9, "blue": 0.96, "alpha": 1}),      # Lavande légère
        "commentaire_range": (create_range(sheet_id, start_row_index, end_row_index + 1, end_col_index + 1, end_col_index + 2), {"red": 0.88, "green": 1.0, "blue": 0.88, "alpha": 1}),         # Vert menthe clair
    }

    static_requests = [
        {
            "mergeCells": {
                "mergeType": "MERGE_COLUMNS",
                "range": ranges_colors["date_range"][0]
            }
        },
        {
            "mergeCells": {
                "mergeType": "MERGE_COLUMNS",
                "range": ranges_colors["commentaire_range"][0]
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
            'updateCells': {
                'rows': [{'values': [{'userEnteredValue': {'stringValue': commentaire}}]}],
                'fields': 'userEnteredValue',
                'range': ranges_colors["commentaire_range"][0]
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
        },
        {
            "repeatCell": {
                "range": ranges_colors["commentaire_range"][0],
                "cell": {
                    "userEnteredFormat": {
                        "wrapStrategy": "WRAP",
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment, verticalAlignment, wrapStrategy)"
            }
        }
    ]

    color_requests = [create_color_request(rng, color) for rng, color in ranges_colors.values()]
    border_requests = [create_border_request(ranges_colors["all_range"][0])]

    requests = static_requests + color_requests + border_requests
    spreadsheet.batch_update({'requests': requests})


if __name__ == "__main__":
    current_date = datetime.now().strftime("%Y-%m-%d")
    program_name = "Programme 1"
    workout_data = {
        'pompes': [15, 16],
        'tractions': [18, 17, 15],
        'dips': [10, 10, ],
        'megastyle': [10, 10],
        'gainage': [15, 30]
    }
    commentaire = "et donc  la j arrive dans la soupe dans la soupe ok ok test micro ok test jarivee dans la soupe"

    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key('16QakQHtFK7TTm5c8-bDBHqSul7BYRNzERqPy6xpYK7M')

    update_spreadsheet(spreadsheet, current_date, program_name, workout_data, commentaire)
