import os
import gspread
import itertools
from oauth2client.service_account import ServiceAccountCredentials
from utils import *
import datetime
import time
import threading
from api_spreadsheet import update_spreadsheet
from database import *
from config import spreadsheet_id

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

message = ""
message_lock = threading.Lock()

def input_date(prompt="Date (DD/MM/YYYY) ?: ", default_date=None):
    if default_date is None:
        default_date = datetime.now()
    
    date_str = input(prompt)
    
    if date_str.strip() == "" or date_str == "0":
        return default_date
    else:
        try:
            return datetime.strptime(date_str, "%d/%m/%Y")
        except ValueError:
            print("Format de date invalide. Utilisation de la date par défaut.")
            return default_date
        
def spinner(stop_event):
    global message
    for frame in itertools.cycle('|/-\\'):
        if stop_event.is_set():
            break
        clear_console()
        with message_lock:
            print(frame + " " + message)
        time.sleep(0.1)

def execute_with_loading(description, func, *args, **kwargs):
    global message
    stop_event = threading.Event()
    spinner_thread = threading.Thread(target=spinner, args=(stop_event,))

    try:
        with message_lock:
            message = description
        spinner_thread.start()
        result = func(*args, **kwargs)
    finally:
        stop_event.set()
        spinner_thread.join()
        clear_console()
    
    return result

def charger_programmes(worksheet_programmes):
    programmes = {}
    for col in range(1, worksheet_programmes.col_count + 1):
        col_values = worksheet_programmes.col_values(col)
        if not any(col_values):
            break
        if col_values[0]:
            programme_name = col_values[0]
            exercices = [ex for ex in col_values[1:] if ex]
            programmes[programme_name] = exercices
    return programmes

def save_poid(worksheet_poids, date : datetime, poid):
    poids = worksheet_poids.col_values(1)
    next_row = len(poids) + 1
    worksheet_poids.update_cell(next_row, 1, date.strftime('%Y-%m-%d'))
    worksheet_poids.update_cell(next_row, 2, poid)


def menu_principal(conn, programmes, current_date, spreadsheet, worksheet_poids):
        while True:
            afficher_menu_principal(programmes)
            choix = input("Choix : ")
            if choix.isdigit():
                choix_int = int(choix)
                if choix_int in range(1, len(programmes) + 1):
                    nom_programme = list(programmes.keys())[choix_int - 1]
                    exercices = programmes[nom_programme]
                    print(f"Programme sélectionné : {nom_programme}")
                    series_repetitions = {}
                    for exercice in exercices:
                        series_repetitions[exercice] = []
                        print(f"Répétitions pour {exercice}.")
                        while True:
                            reps = input(f"Nombre de répétitions pour {exercice} (0 pour suivant) : ")
                            if reps.isdigit():
                                if reps == '0':
                                    break
                                else:
                                    series_repetitions[exercice].append(int(reps))
                                    print(f"{exercice}: {reps} répétitions ajoutées.")
                    commentaire = input("Commentaire pour la séance : ")
                    date = input_date("Date (DD/MM/YYYY) ?", default_date=current_date)
                    
                    execute_with_loading("Sauvegarde de la session", update_spreadsheet, conn, spreadsheet, date, nom_programme, series_repetitions, commentaire)
                    print(f"Récapitulatif pour le programme {nom_programme}:")
                    for exercice, series in series_repetitions.items():
                        series_str = ', '.join(map(str, series))
                        print(f"{exercice}: Séries -> {series_str}")
                elif choix_int == len(programmes) + 1:
                    poid = input("Poids actuel (en kg) : ")
                    date = input_date("Date (DD/MM/YYYY) ?", default_date=current_date)
                    execute_with_loading("Sauvegarde du poids", save_poid, worksheet_poids, date, poid)
                elif choix_int == 0:
                    print("Fermeture de l'application de musculation.")
                    break
                else:
                    print("Choix non valide. Réessayer.")
            else:
                print("Entrée invalide.")


if __name__ == "__main__":
    creds = execute_with_loading("Connexion à l'API Google", ServiceAccountCredentials.from_json_keyfile_name, "credentials.json", scope)
    client = execute_with_loading("Autorisation Google Sheets", gspread.authorize, creds)

    spreadsheet = execute_with_loading("Ouverture du fichier de Google Sheets", client.open_by_key, spreadsheet_id)
    worksheet_programmes = execute_with_loading("Accès à la feuille 'programmes'", spreadsheet.worksheet, 'programmes')
    worksheet_poids = execute_with_loading("Accès à la feuille 'poids'", spreadsheet.worksheet, 'poids')
    worksheet_seances = execute_with_loading("Accès à la feuille 'test_api'", spreadsheet.worksheet, 'test_api')

    conn = execute_with_loading("Configuration de la base de données", setup_database)
    current_date = datetime.now()
    programmes = execute_with_loading("Chargement des programmes", charger_programmes, worksheet_programmes)
    
    menu_principal(conn, programmes, current_date, spreadsheet, worksheet_poids)
    conn.close()
