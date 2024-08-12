import gspread
from oauth2client.service_account import ServiceAccountCredentials
import menus  
import datetime
import json 
from api_spreadsheet import update_spreadsheet
from database import *

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

spreadsheet = client.open_by_key('16QakQHtFK7TTm5c8-bDBHqSul7BYRNzERqPy6xpYK7M')
worksheet_programmes = spreadsheet.worksheet('programmes')
worksheet_poids = spreadsheet.worksheet('poids')
worksheet_seances = spreadsheet.worksheet('test_api')

def charger_programmes():
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

def save_poid(worksheet_poids, current_date, poid):
    poids = worksheet_poids.col_values(1)
    next_row = len(poids) + 1
    worksheet_poids.update_cell(next_row, 1, current_date)
    worksheet_poids.update_cell(next_row, 2, poid)

def save_seance(conn, current_date, spreadsheet, nom_programme, series_repetitions, commentaire):
    update_spreadsheet(conn, spreadsheet, current_date, nom_programme, series_repetitions, commentaire)
    

def menu_principal(conn, programmes, current_date):
    while True:
        menus.afficher_menu_principal(programmes)
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
                save_seance(conn, current_date, spreadsheet, nom_programme, series_repetitions, commentaire)
                print(f"Récapitulatif pour le programme {nom_programme}:")
                for exercice, series in series_repetitions.items():
                    series_str = ', '.join(map(str, series))
                    print(f"{exercice}: Séries -> {series_str}")
            elif choix_int == len(programmes) + 1:
                poid = input("Poids actuel (en kg) : ")
                save_poid(worksheet_poids, current_date, poid)
            elif choix_int == 0:
                print("Fermeture de l'application de musculation.")
                break
            else:
                print("Choix non valide. Réessayer.")
        else:
            print("Entrée invalide.")

if __name__ == "__main__":
    print("hello")
    conn = setup_database()
    current_date = datetime.now()
    programmes = charger_programmes()
    menu_principal(conn, programmes, current_date)
    conn.close()