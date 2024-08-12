import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from config import *
import locale
from database import *

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

def create_range(sheet_id, default_row, end_row, default_col, end_col):
    return {
        "sheetId": sheet_id,
        "startRowIndex": default_row,
        "endRowIndex": end_row,
        "startColumnIndex": default_col,
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

def update_spreadsheet(conn, spreadsheet : gspread.Spreadsheet, current_date : datetime, program_name, workout_data, commentaire):
    
    original_locale = locale.getlocale(locale.LC_TIME)
    locale.setlocale(locale.LC_TIME, 'fr_FR')
    
    current_date_worskeet = datetime.now()
    worksheet_name = current_date_worskeet.strftime("%B_%Y")

    worksheet = next((sheet for sheet in spreadsheet.worksheets() if sheet.title == worksheet_name), None)

    if worksheet is None:
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="1000", cols="26")
        add_worksheet_info(conn, worksheet.id, default_col, default_row)
        worksheet_info = get_worksheet_info_by_id(conn, worksheet.id)
        print(f"Nouvelle worksheet créée: {worksheet_name}")
    else:
        worksheet_info = get_worksheet_info_by_id(conn, worksheet.id)
        if not worksheet_info:
            add_worksheet_info(conn, worksheet.id, default_col, default_row)
        print(f"Une worksheet avec le nom {worksheet.title} existe déjà.")
    
    worksheet_info = get_worksheet_info_by_id(conn, worksheet.id)
    
    locale.setlocale(locale.LC_TIME, original_locale)
    
    
    current_col, current_row, last_updated = worksheet_info[2], worksheet_info[3], worksheet_info[4]
    
    headers = [program_name] + [exercise.capitalize() for exercise in workout_data.keys()]
    max_height = max(len(v) for v in workout_data.values())
    max_width = len(headers)
    data = [headers] + [[f"série {i + 1}"] + [workout_data[exercise][i] if i < len(workout_data[exercise]) else "" for exercise in workout_data] for i in range(max_height)]

    start_col_index = col_letter_to_num(current_col)
    end_col_index = start_col_index + len(headers) - 1
    start_row_index = current_row - 1
    end_row_index = start_row_index + max_height  
    
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
                'rows': [{'values': [{'userEnteredValue': {'stringValue': current_date.strftime('%Y-%m-%d')}}]}],
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

    last_updated = datetime.strptime(last_updated, '%Y-%m-%d').date()
    if last_updated == current_date.date():
        new_next_col = col_num_to_letter(col_letter_to_num(current_col) + space_between_col + max_width + 2)
        new_next_row = current_row
    else:
        new_next_col = default_col
        new_next_row = current_row + space_between_row + max_height

    update_worksheet_info(conn, sheet_id, new_next_col, new_next_row)
    

if __name__ == "__main__":
    conn = setup_database()
    current_date = datetime.now()
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

    update_spreadsheet(conn, spreadsheet, current_date, program_name, workout_data, commentaire)
    conn.close()