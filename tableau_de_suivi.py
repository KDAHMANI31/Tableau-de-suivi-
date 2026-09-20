import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(
    page_title="Suivi Non-Conformités Cuisine",
    page_icon="🍳",
    layout="wide",
)

st.title("🍳 Suivi des Non-Conformités - Cuisine Centrale")
st.write(
    "Application de gestion et d'exportation des audits HACCP et contrôles d'hygiène."
)

# Données d'exemple
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
        "Description de la Non-Conformité": "Colis de poisson frais reçus sans étiquette de traçabilité.",
        "Niveau de Risque": "Majeur",
        "Action Corrective Immédiate": "Refus de la livraison auprès du fournisseur.",
        "Responsable Action": "Responsable Achats",
        "Date Échéance": "2026-03-16",
        "Statut": "Clôturé",
    },
    {
        "ID": "NC-2026-003",
        "Date Détection": "2026-03-18",
        "Secteur / Zone": "Légumerie / Machines à glaçons",
        "Catégorie HACCP": "Hygiène & Nettoyage",
        "Description de la Non-Conformité": "Machine à glaçons : absence de contrôle de propreté des surfaces en contact avec la glace.",
        "Niveau de Risque": "Majeur",
        "Action Corrective Immédiate": "Nettoyage complet et vérification des joints et bacs.",
        "Responsable Action": "Responsable Stewarding",
        "Date Échéance": "2026-03-20",
        "Statut": "En cours",
    },
]

df = pd.DataFrame(donnees)

# Affichage du tableau dans Streamlit
st.subheader("📋 Tableau récapitulatif")
st.dataframe(df, use_container_width=True)


# Fonction de génération Excel
def generer_excel(df_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi NC Cuisine"

    ws.merge_cells("A1:J1")
    title_cell = ws["A1"]
    title_cell.value = "SUIVI DES NON-CONFORMITÉS - CUISINE CENTRALE HÔTEL"
    title_cell.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    headers = list(df_data.columns)
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

    ws.row_dimensions[3].height = 25

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for row_idx, row_data in enumerate(df_data.values, 4):
        ws.row_dimensions[row_idx].height = 22
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.font = Font(name="Arial", size=9)
            cell.border = thin_border
            if col_idx in [1, 2, 6, 9, 10]:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
            else:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row == 1:
                continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# Bouton de téléchargement
st.subheader("📥 Exporter les données")
excel_data = generer_excel(df)
st.download_button(
    label="Télécharger le rapport Excel (.xlsx)",
    data=excel_data,
    file_name="Suivi_Non_Conformites_Cuisine.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
