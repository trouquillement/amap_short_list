Cette librairie sert à construire une short liste pour faciliter la distribution des AMAP.
Elle part des récapitulatifs de chaque contrat des AMAPIENS (issue de clic amap) et les rassemblent
en un document synthétique pour la distribution.

# Générateur Excel AMAP

Outil Python permettant de générer automatiquement des fichiers Excel de distribution AMAP à partir des
exports produits de clic amap. La connexion avec clic amap n'est pas direct. L'utilisateur doit téléchargement
lui même les exports de chaque produit sur clic amap et les ranger dans un même dossier.

L'application permet :

- Sélection d’un dossier contenant les fichiers sources
- Sélection d’une ou plusieurs dates de livraison (mardis dans le prototype car pour une AMAP en particulier)
- Fusion automatique des produits par Nom / Prénom
- Gestion automatique des lignes "total"
- Mise en forme Excel (bordures, alternance de couleurs, fusion cellules)
- Interface graphique simple via Tkinter 
- Génération possible d'un .exe pour distribution aux utilisateurs en amap (plus besoin de coder)

## Structure du projet
short_list/
│
├── short_list short_list/
    │
    ├── interface.py          # Interface principale Tkinter
    ├── interface_dates.py    # Sélection des dates de livraison
    ├── traitement.py         # Logique de fusion et génération Excel
    ├── mapping.py            # quelques mapping pour faciliter la lecture de la short list
    ├── __init__.py        
├── amap_short_list.py        # Point d'entrée (si utilisé par python, mais aussi pour générer l'executable)
└── README.md

## Installation (mode développeur)
1 - Créer un environnement virtuel

(ajouter tuto)
(ajouter un requirement.txt ou une manière d'installer l'env (fichier .yml ?))

2 - Installer les dépendances
pip install pandas openpyxl xlsxwriter pyinstaller (ça dépend si l'env est fourni dans la lib?)
Tkinter est inclus par défaut avec Python.

## Lancer l'application
se placer dans le dossier de la lib: \short_list
puis executer:

python amap_short_list.py

Les fichiers de distribution de chaque contrat sont téléchargés sur clic amap selon cette procédure:
- gestionnaire référent
- gestion des contrats signés
- choisir la ferme et le contrat
- télécharger
- Toutes les feuilles de livraisons (excel)

/!\ attention, les fichiers téléchargés sur clic amap sont au format xls !
à ce jour, il faut d'abord changer le format en xlsx pour que ça fonctionne.

Ensuite les fichiers doivent être mis dans un dossier par type de produit.
Les dossiers pour chaque produit doivent être rassemblé dans un dossier que l'on sélectionne
en lançant l'application.
L'architecture est dans le dossier "fichiers_exemple".

Le fichier de sortie d'exemple est aussi dans ce dossier: "short_list_demo.xlsx"

## Générer l'exécutable (.exe)

Depuis l’environnement activé :

pyinstaller --onefile --noconsole amap_short_list.py

Le fichier généré sera disponible dans :
short_list/dist/amap_short_list.exe

### Gestion des dates

Les dates proposées sont automatiquement les prochains mardis (≈ 3 mois).
(ajouter une option pour un autre jour de distribution ?)
L'utilisateur peut sélectionner une ou plusieurs dates.
Si aucune date n’est sélectionnée, toutes les dates proposées sont utilisées.
 
### Logique de fusion des produits

Pour chaque date :
Les contrats légumes sont utilisés en priorité comme base.
Les autres produits sont fusionnés via un outer merge sur :
Nom
Prénom
Une ligne "total" est ajoutée en fin de tableau pour faciliter la vérification des produits en début
de distribution.
Si le nombre de colonnes dépasse max_cols, les lignes sont automatiquement découpées proprement.

### Mise en forme Excel

Bordures horizontales sur toutes les lignes.
Bordures complètes sur la ligne "total".
Alternance de couleurs par bénéficiaire pour la visibilité.
Fusion des cellules Nom / Prénom si plusieurs lignes.
Largeur de colonnes ajustée.


## Auteur
trouquillement
Projet développé pour automatiser la génération des feuilles de distribution AMAP.
Certains éléments ont été développé avec l'aide d'un LLM.

## licence 
à voir ?