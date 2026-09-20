import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def generer_tableau_non_conformites():
    # 1. Données d'exemple adaptées aux cuisines centrales d'hôtel
    donnees = [
        {
            "ID": "NC-2026-001",
            "Date Détection": "2026-03-15",
            "Secteur / Zone": "Chambre Froide Positive 1",
            "Catégorie HACCP": "Chaîne du froid",
            "Description de la Non-Conformité": "Température relevée à +8°C au lieu de +3°C max.",
            "Niveau de Risque": "Critique",
            "Action Corrective Immédiate": "Mise en isolement des denrées et transfert vers CF2.",
            "Responsable Action": "Chef Équipe Froid",
            "Date Échéance": "2026-03-15",
            "Statut": "Clôturé",
        },
        {
            "ID": "NC-2026-002",
            "Date Détection": "2026-03-16",
            "Secteur / Zone": "Zone Réception / Quai",
            "Catégorie HACCP": "Traçabilité & Livraison",
            "Description de la Non-Conformité": "Colis de poisson frais reçus sans étiquette de traçabilité d'origine.",
            "Niveau de Risque": "Majeur",
            "Action Corrective Immédiate": "Refus de la livraison auprès du fournisseur et enregistrement au registre.",
            "Responsable Action": "Responsable Achats / Réception",
            "Date Échéance": "2026-03-16",
            "Statut": "Clôturé",
        },
        {
            "ID": "NC-2026-003",
            "Date Détection": "2026-03-18",
            "Secteur / Zone": "Légumerie / Préparation",
            "Catégorie HACCP": "Hygiène & Nettoyage",
            "Description de la Non-Conformité": "Plan de nettoyage non respecté sur les plongeurs automatiques.",
            "Niveau de Risque": "Mineur",
            "Action Corrective Immédiate": "Refaire le nettoyage immédiat et rappel des procédures à l'équipe du soir.",
            "Responsable Action": "Responsable Stewarding",
            "Date Échéance": "2026-03-20",
            "Statut": "En cours",
        },
        {
            "ID": "NC-2026-004",
            "Date Détection": "2026-03-20",
            "Secteur / Zone": "Cuisine Chaude / Pâtisserie",
            "Catégorie HACCP": "Matériel & Maintenance",
            "Description de la Non-Conformité": "Joint de porte du four à vapeur endommagé.",
            "Niveau de Risque": "Majeur",
            "Action Corrective Immédiate": "Demande d'intervention urgente auprès du service technique.",
            "Responsable Action": "Chef de Cuisine / Tech",
            "Date Échéance": "2026-03-25",
            "Statut": "Ouvert",
        },
    ]

    df = pd.DataFrame(donnees)

    # 2. Création du fichier Excel avec OpenPyXL
    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi NC Cuisine"

    # En-tête principal
    ws.merge_cells("A1:J1")
    title_cell = ws["A1"]
    title_cell.value = "SUIVI DES NON-CONFORMITÉS - CUISINE CENTRALE HÔTEL"
    title_cell.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    # Colonnes de données (ligne 3)
    headers = list(df.columns)
    for col_num, header_title in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.value = header_title
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            start_color="2F5597", end_color="2F5597", fill_type="solid"
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = Border(
            bottom=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
        )

    ws.row_dimensions[3].height = 25

    # Injection des données
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for row_idx, row_data in enumerate(df.values, 4):
        ws.row_dimensions[row_idx].height = 22
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.font = Font(name="Arial", size=9)
            cell.border = thin_border

            # Alignements spécifiques
            if col_idx in [1, 2, 6, 9, 10]:  # ID, Dates, Risque, Statut
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
            else:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

    # 3. Mise en forme conditionnelle pour les Statuts
    red_fill = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )
    red_font = Font(color="C00000", bold=True, name="Arial", size=9)

    yellow_fill = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )
    yellow_font = Font(color="B25900", bold=True, name="Arial", size=9)

    green_fill = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )
    green_font = Font(color="375623", bold=True, name="Arial", size=9)

    last_row = len(df) + 3
    statut_range = f"J4:J{last_row}"

    ws.conditional_formatting.add(
        statut_range,
        CellIsRule(
            operator="equal",
            formula=['"Ouvert"'],
            fill=red_fill,
            font=red_font,
        ),
    )
    ws.conditional_formatting.add(
        statut_range,
        CellIsRule(
            operator="equal",
            formula=['"En cours"'],
            fill=yellow_fill,
            font=yellow_font,
        ),
    )
    ws.conditional_formatting.add(
        statut_range,
        CellIsRule(
            operator="equal",
            formula=['"Clôturé"'],
            fill=green_fill,
            font=green_font,
        ),
    )

    # 4. Ajustement automatique de la largeur des colonnes
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            # Ignorer la première ligne fusionnée pour le calcul des largeurs
            if cell.row == 1:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        # Ajustement sur mesure selon le type de contenu
        if col_letter in ["E", "G"]:  # Descriptions et Actions
            ws.column_dimensions[col_letter].width = 35
        elif col_letter in ["C", "D"]:  # Secteurs et Catégories
            ws.column_dimensions[col_letter].width = 24
        else:
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Sauvegarde du fichier Excel
    nom_fichier = "Suivi_Non_Conformites_Cuisine.xlsx"
    wb.save(nom_fichier)
    print(f"Tableau généré avec succès : {nom_fichier}")


if __name__ == "__main__":
    generer_tableau_non_conformites()
