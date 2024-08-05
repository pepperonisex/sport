def afficher_menu_principal(programmes):
    print("Choisir le programme:")
    for index, programme in enumerate(programmes, start=1):
        print(f"{index} - {programme}")
    print(f"{len(programmes) + 1} - Poid")
    print("0 - Quitter")


