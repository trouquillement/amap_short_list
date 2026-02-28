import os
import pandas as pd
import glob
from .mapping import mapping


def parse_clic_amap_file(file_path, produit):
    """
    Traite un fichier ClicAMAP et retourne un dictionnaire {date: DataFrame_propre}
    """
    result = {}
    file_name = os.path.basename(file_path)
    # On suppose le format du fichier : produit.xlsx

    # Lire toutes les feuilles du fichier Excel
    sheets = pd.read_excel(file_path, sheet_name=None)

    for sheet_name, df_raw in sheets.items():
        # Sauter les feuilles vides ou trop petites
        if df_raw.empty or df_raw.shape[0] < 5:
            continue

        try:
            # --- Traitement spécifique ---
            # Trouver la ligne "Nom"
            start_row = df_raw[df_raw.iloc[:, 0] == "Nom"].index[0]

            # Construire dynamiquement les noms de colonnes
            column_names = []
            for col_idx in range(df_raw.shape[1]):
                if col_idx <= 1:
                    column_names.append(df_raw.iloc[start_row, col_idx])
                else:
                    # permet de récupérer le nom du produit et la quantité (carton ou bouteille pour les bières)
                    col_name = f"{df_raw.iloc[start_row+1, col_idx]}"
                    # col_name = f"{df_raw.iloc[start_row+1, col_idx]}_{df_raw.iloc[start_row+3, col_idx]}" # permet de récupéré le nom du produit et la quantitié (carton ou bouteille pour les bières)
                    if pd.notna(col_name):
                        if mapping[produit] is not None:
                            column_names.append(mapping[produit][col_name])
                        else:
                            column_names.append(col_name)

            # Trouver la ligne "Cumul"
            cumul_row = df_raw[df_raw.iloc[:, 0] == "Cumul"].index[0]
            data_start_row = cumul_row + 2 

            # Extraire les données sous forme de DataFrame
            df_data = df_raw.iloc[data_start_row:, :len(column_names)]
            df_data.columns = column_names

            # Nettoyage : remplacer les NaN numériques par 0 et cast en int
            for col in df_data.columns[2:]:
                df_data[col] = df_data[col].fillna(0).astype(int)

            # on mets l'index à Nom
            # df_data.index = df_data["Nom"]
            
            # On stocke par date
            result[sheet_name] = df_data

        except Exception as e:
            print(f"Erreur lors du traitement de la feuille '{sheet_name}' dans {file_name}: {e}")
            continue

    return result


def create_amap_dict(fichier_clic_amap_path):
    """
    Créer un dictionnaire {date: DataFrame_propre}.

    Chaque clé du dictionnaire de sortie est une date de livraison.
    Chaque valeurs est un tableau avec des colonnes: nom et prénom
    + une colonne par produit commandé avec la quantité.

    On ajoute aussi le total par ligne et par colonne.

    parameters
    -----------
    fichier_clic_amap_path: str | Path ?
        dossier où sont stocké les fichiers extraits de clic amap

    return
    -----------
    amap_dict: dict
        rendu final. {date: DataFrame_propre}
    """
    amap_dict = {}

    # boucle sur chaque dossier contenu dans le dossier indiqué
    for produit in os.listdir(fichier_clic_amap_path):
        produit_path = os.path.join(fichier_clic_amap_path, produit, "*.xlsx")
        produit_files = glob.glob(produit_path)

        if not produit_files:
            continue  # Aucun fichier pour ce produit

        for file_path in produit_files:
            # Appeler la fonction pour parser ce fichier
            parsed_data = parse_clic_amap_file(file_path, produit)

            if not parsed_data:
                continue  # Ignorer si parsing vide ou en erreur

            # Initialiser le produit dans amap_dict si nécessaire
            if produit not in amap_dict:
                amap_dict[produit] = {}

            # Ajouter toutes les dates et leurs DataFrames dans amap_dict
            amap_dict[produit].update(parsed_data)

    return amap_dict


def create_short_list_dict(amap_dict, date_livraison_list=None):
    """
    Construit le dictionnaire pour la short list

    Parameters
    -----------
    amap_dict: str | Path ?
        dictionnaire de livraison construit par la fonction create_amap_dict.
    date_livraison_list: list | None
        Liste des dates de livraison pour la short list.

    Return
    -----------
    short_list_dict: dict
    """
    short_list_dict = {}
    for date_livraison in date_livraison_list:
        print(date_livraison)

        short_list_df = None

        # 1. On commence avec les légumes si dispo
        if "légumes" in amap_dict and date_livraison in amap_dict["légumes"]:
            short_list_df = amap_dict["légumes"][date_livraison]

        # 2. Sinon on prend le premier produit dispo
        if short_list_df is None:
            for produit, produits_dict in amap_dict.items():
                if date_livraison in produits_dict:
                    short_list_df = produits_dict[date_livraison]
                    break  # on arrête après le premier trouvé

        # 3. Fusion avec les autres produits (sauf la base déjà choisie)
        for produit, produits_dict in amap_dict.items():
            if date_livraison in produits_dict:
                if produit == "légumes":
                    continue  # déjà pris comme base
                short_list_df = pd.merge(
                    short_list_df,
                    produits_dict[date_livraison],
                    on=["Nom", "Prénom"],
                    how="outer"
                )

        # Ajouter la colonne "Total" = somme de toutes les colonnes de produits par ligne
        short_list_df["Total"] = short_list_df.drop(columns=["Nom", "Prénom"]).sum(axis=1)

        # Créer la ligne "total"

        total_row = short_list_df.sum()
        total_row['Nom'] = 'total'
        total_row['Prénom'] = "toto"
        # Ajouter la ligne au DataFrame
        short_list_df = pd.concat([short_list_df, pd.DataFrame([total_row])], ignore_index=True)
        short_list_dict[date_livraison] = short_list_df

    return short_list_dict


# Paramètre d'entrée : nombre max de produits par ligne
MAX_PRODUCTS = 5


def explode_row(nom, prenom, produits, max_products=MAX_PRODUCTS):
    """
    Transforme une ligne (Nom, Prénom, [produits...]) en plusieurs lignes si besoin.
    Retourne une liste de listes : [[Nom, Prénom, prod1..prodN], ...]
    """
    chunks = [produits[i:i+max_products] for i in range(0, len(produits), max_products)]
    rows = []
    for i, chunk in enumerate(chunks):
        row = [nom, prenom] + chunk
        # compléter jusqu'à max_products
        row.extend([""] * (max_products - len(chunk)))
        rows.append(row)
    return rows


def construct_excel_from_dict(short_list_dict, output_file):
    """
    Fonction pour construire et formater le fichier excel à partir d'une dictionnaire.

    Parameters:
    ------------
    short_list_dict: dict
    output_file: str
        Chemin d'accès complet avec le nom du fichier excel de sortie.

    return
    ------------
    """
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book

        # === Formats ===
        title_format = workbook.add_format({
            'bold': True, 'align': 'left', 'valign': 'vcenter',
            'font_size': 14
        })

        header_format = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#D9E1F2'
        })

        row_format_white = workbook.add_format({'border': 1, 'bg_color': "#FFFFFF"})
        row_format_gray = workbook.add_format({'border': 1, 'bg_color': "#D9D9D9"})

        total_block_format = workbook.add_format({
            'border': 1, 'bg_color': '#B0B0B0',
            'bold': True, 'align': 'center', 'valign': 'vcenter'
        })

        for date_livraison, df in short_list_dict.items():
            worksheet = workbook.add_worksheet(date_livraison[:31])

            # === Ligne de titre ===
            worksheet.write(0, 0, f"Date de livraison : {date_livraison}", title_format)

            # === En-têtes ===
            worksheet.write(1, 0, "Nom", header_format)
            worksheet.write(1, 1, "Prénom", header_format)
            worksheet.merge_range(1, 2, 1, 2 + MAX_PRODUCTS - 1, "Produits_commandés", header_format)

            row_idx = 2  # ligne Excel courante
            fmt_idx = 0

            for r, row in enumerate(df.itertuples(index=False)):
                nom, prenom = row[0], row[1]

                # Construire la liste de produits commandés
                produits = []
                for col, val in zip(df.columns[2:], row[2:]):
                    if val > 0:
                        produits.append(f"{col}: {int(val)}")

                # Découper la ligne en blocs
                exploded_rows = explode_row(nom, prenom, produits, MAX_PRODUCTS)
                total_lines = len(exploded_rows)

                # Choix du format : normal ou total
                if nom == "total" and prenom == "toto":
                    fmt = total_block_format
                else:
                    fmt = row_format_gray if (fmt_idx % 2 == 0) else row_format_white

                # Fusion Nom et Prénom si plusieurs lignes
                if total_lines > 1:
                    worksheet.merge_range(row_idx, 0, row_idx + total_lines - 1, 0, nom, fmt)
                    worksheet.merge_range(row_idx, 1, row_idx + total_lines - 1, 1, prenom, fmt)
                else:
                    worksheet.write(row_idx, 0, nom, fmt)
                    worksheet.write(row_idx, 1, prenom, fmt)

                # Écrire les produits
                for i, erow in enumerate(exploded_rows):
                    if i == 0:
                        fmt_idx += 1
                    for j, val in enumerate(erow[2:]):
                        if val != "   ":
                            worksheet.write(row_idx + i, 2 + j, val, fmt)
                        

                row_idx += total_lines

            # Largeur des colonnes
            worksheet.set_column(0, 0, 15)  # Nom
            worksheet.set_column(1, 1, 15)  # Prénom
            for col in range(2, 2 + MAX_PRODUCTS):
                worksheet.set_column(col, col, 25)


def build_excel(input_folder, output_file, date_livraison_list):
    """
    fonction d'orchestration des autres.

    Parameters:
    -------------
    input_folder:
    output_file:
    date_livraison_list: list
        Listes des dates de livraisons

    return:
    ---------
    """

    amap_dict = create_amap_dict(input_folder)

    #date_livraison_list = choisir_dates()
    short_list_dict = create_short_list_dict(amap_dict=amap_dict, date_livraison_list=date_livraison_list)

    construct_excel_from_dict(short_list_dict=short_list_dict, output_file=output_file)
    
    print(f"Fichier généré : {output_file}")

    return output_file


if __name__ == '__main__':
    fichier_clic_amap_path = r"C:\Users\toesca\OneDrive - CSTBGroup\Documents\PERSO\amap\fichiers_clic_amap"
    output_file = os.path.join(
    r"C:\Users\toesca\OneDrive - CSTBGroup\Documents\PERSO\amap",
    "commandes_clic_amap_blocs_pytest.xlsx"
)
    date_livraison_list = ['2025-08-05']
    build_excel(input_folder=fichier_clic_amap_path, output_file=output_file, date_livraison_list=date_livraison_list)

    #['2025-09-09', '2025-09-16', '2025-09-23']