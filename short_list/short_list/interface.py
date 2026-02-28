import tkinter as tk
from tkinter import filedialog, messagebox
import os
from .traitement import build_excel
from .interface_dates import generer_mardis
import datetime


def choisir_dossier():
    """
    fonction pour choisir le dossier de stockage des fichiers extraits de clic amap.
    """
    dossier = filedialog.askdirectory()
    if dossier:
        dossier_var.set(dossier)


def choisir_sortie():
    """
    Fonction pour nommer le nom du fichier Excel de sorties
    """
    fichier = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                           filetypes=[("Excel files", "*.xlsx")])
    if fichier:
        sortie_var.set(fichier)


def generer_excel():
    """
    Fonction pour construire le bouton permettant d'appeler la fontion
    qui génère le Excel final.
    """
    # va chercher le dossier où sont stockés les fichiers
    dossier = dossier_var.get()
    # récupère le nom du fichier de sortie
    sortie = sortie_var.get() or os.path.join(os.path.expanduser("~"), "Downloads", "short_list.xlsx")

    # vérifie que le dossier existe
    if not dossier:
        messagebox.showerror("Erreur", "Merci de choisir un dossier source")
        return
    # lance l'execution de la génération du excel
    try:
        fichier_genere = build_excel(dossier, sortie, date_livraison_list=selections)
        messagebox.showinfo("Succès", f"Fichier généré :\n{fichier_genere}")
    # todo: améliorer le message d'erreur pour faciliter le debugage
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue :\n{e}")


root = tk.Tk()
root.title("Générateur Excel AMAP")

window_width = 300
window_height = 600

# récupère les dimensions de la fenêtre d'affichage
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# trouve le centre de l'écran
center_x = int(screen_width/2 - window_width / 2)
center_y = int(screen_height/2 - window_height / 2)

# place la fenêtre de l'outil au centre de l'écran
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')


dossier_var = tk.StringVar()
sortie_var = tk.StringVar()

tk.Label(root, text="Dossier source :").pack()
tk.Entry(root, textvariable=dossier_var, width=50).pack()
tk.Button(root, text="Parcourir...", command=choisir_dossier).pack()

tk.Label(root, text="Fichier Excel final (optionnel) :").pack()
tk.Entry(root, textvariable=sortie_var, width=50).pack()
tk.Button(root, text="Parcourir...", command=choisir_sortie).pack()

# génère les dates qui seront affichés dans l'outils
# todo: ajouter option pour changer de jour de distributions ?
mardis = generer_mardis(debut=datetime.datetime(2025, 7, 14), nb_semaines=12)  # par défaut 3 mois
selections = []

tk.Label(root, text="Sélectionnez les dates de livraison :", font=("Arial", 12)).pack(pady=10)

vars_dict = {}
for date in mardis:
    var = tk.BooleanVar()
    cb = tk.Checkbutton(root, text=date.strftime("%Y-%m-%d"), variable=var)
    cb.pack(anchor="w")
    vars_dict[date] = var

def valider():
    dates_choisies = [date.strftime("%Y-%m-%d") for date, var in vars_dict.items() if var.get()]
    if not dates_choisies:
        # si rien choisi, prendre toutes les dates
        dates_choisies = [dat.strftime("%Y-%m-%d") for dat in mardis]
        print(dates_choisies)
    selections.extend(dates_choisies)
    return selections

btn = tk.Button(root, text="Valider les dates", command=valider, bg="#4CAF50", fg="white")
btn.pack(pady=10)

tk.Button(root, text="Générer Excel", command=generer_excel, bg="lightgreen").pack(pady=10)

root.mainloop()

