import datetime
import tkinter as tk
from tkinter import filedialog, messagebox


def generer_mardis(debut=None, nb_semaines=12):
    """
    Génère une liste de mardis à partir d'une date donnée (par défaut aujourd'hui).

    NOTES:
        AJOUTER OPTIONS POUR SELECTIONNER AUTRE JOUR DE DISTRIBUTION QUE MARDI !
    """
    if debut is None:
        debut = datetime.date.today()
        print(debut)

    # trouver le prochain mardi
    jours_a_attendre = (1 - debut.weekday()) % 7  # lundi=0, mardi=1, etc.
    premier_mardi = debut + datetime.timedelta(days=jours_a_attendre)

    # générer la liste des mardis pour nb_semaines
    mardis = [premier_mardi + datetime.timedelta(weeks=i) for i in range(nb_semaines)]
    return mardis


def choisir_dates():
    """
    Affiche une fenêtre Tkinter permettant de choisir une ou plusieurs dates.
    """
    root_dates = tk.Tk()
    root_dates.title("Choix des dates de livraison")

    # date de début en test: à sortir et construire des tests !
    # datetime.datetime(2025, 7, 14)
    mardis = generer_mardis(nb_semaines=12)  # par défaut 3 mois
    selections = []

    tk.Label(root_dates, text="Sélectionnez les dates de livraison :", font=("Arial", 12)).pack(pady=10)

    vars_dict = {}
    # vérifie quelles dates sont sélectionnées par l'utilisateur
    for date in mardis:
        var = tk.BooleanVar()
        cb = tk.Checkbutton(root_dates, text=date.strftime("%Y-%m-%d"), variable=var)
        cb.pack(anchor="w")
        vars_dict[date] = var

    # applique les dates choisies par l'utilisateur
    def valider():
        dates_choisies = [date.strftime("%Y-%m-%d") for date, var in vars_dict.items() if var.get()]
        if not dates_choisies:
            # si rien choisi, prendre toutes les dates
            dates_choisies = [dat.strftime("%Y-%m-%d") for dat in mardis]
        selections.extend(dates_choisies)
        root_dates.destroy()
        print(selections)
        return selections

    btn = tk.Button(root_dates, text="Valider", command=valider, bg="#4CAF50", fg="white")
    btn.pack(pady=10)

    root_dates.mainloop()
    print(selections)
    return selections


# Test rapide
if __name__ == "__main__":
    dates = choisir_dates()
    print("Dates retenues :", [d for d in dates])

