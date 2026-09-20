import datetime
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

st.title("🍳 Plateforme de Suivi des Non-Conformités - Cuisine Centrale")
st.write(
    "Saisissez vos contrôles HACCP, gérez vos données en direct et exportez le rapport Excel moderne."
)

# Initialisation des données
if "df_nc" not in st.session_state:
    st.session_state.df_nc = pd.DataFrame(
        [
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
    )

# ---------------------------------------------------------
# BARRE LATÉRALE : FORMULAIRE DE SAISIE
# ---------------------------------------------------------
st.sidebar.header("➕ Saisir une Non-Conformité")

with st.sidebar.form(key="form_saisie", clear_on_submit=True):
    next_id = f"NC-2026-00{len(st.session_state.df_nc) + 1}"
    date_detec = st.date_input(
        "Date de Détection", value=datetime.date.today()
    )

    secteur = st.selectbox(
        "Secteur / Zone",
        [
            "Chambre Froide Positive",
            "Chambre Froide Négative",
            "Zone Réception / Quai",
            "Cuisine Chaude",
            "Pâtisserie",
            "Légumerie",
            "Zone Plonge / Stewarding",
            "Local Glaçons",
        ],
    )

    categorie = st.selectbox(
        "Catégorie HACCP",
        [
            "Chaîne du froid",
            "Traçabilité & Livraison",
            "Hygiène & Nettoyage",
            "Matériel & Maintenance",
            "Tenue & Hygiène Personnel",
            "Gestion des Déchets",
        ],
    )

    description = st.text_area("Description de la Non-Conformité")
    risque = st.select_slider(
        "Niveau de Risque", options=["Mineur", "Majeur", "Critique"]
    )
    action = st.text_area("Action Corrective Immédiate")
    responsable = st.text_input("Responsable de l'Action")
    date_echeance = st.date_input("Date Échéance", value=datetime.date.today())
    statut = st.selectbox("Statut", ["Ouvert", "En cours", "Clôturé"])

    btn_submit = st.form_submit_button(label="Enregistrer la NC")

if btn_submit:
    nouvelle_nc = {
        "ID": next_id,
        "Date Détection": str(date_detec),
        "Secteur / Zone": secteur,
        "Catégorie HACCP": categorie,
        "Description de la Non-Conformité": description,
        "Niveau de Risque": risque,
        "Action Corrective Immédiate": action,
        "Responsable Action": responsable,
        "Date Échéance": str(date_echeance),
        "Statut": statut,
    }

    st.session_state.df_nc = pd.concat(
        [st.session_state.df_nc, pd.DataFrame([nouvelle_nc])], ignore_index=True
    )
    st.sidebar.success(f"Non-conformité {next_id} enregistrée !")

# ---------------------------------------------------------
# AFFICHAGE INTERACTIF
# ---------------------------------------------------------
st.subheader("📋 Tableau dynamique des Non-Conformités")

edited_df = st.data_editor(
    st.session_state.df_nc,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Statut": st.column_config.SelectboxColumn(
            "Statut", options=["Ouvert", "En cours", "Clôturé"], required=True
        ),
        "Niveau de Risque": st.column_config.SelectboxColumn(
            "Niveau de Risque",
            options=["Mineur", "Majeur", "Critique"],
            required=True,
        ),
    },
)

st.session_state.df_nc = edited_df


# ---------------------------------------------------------
# GÉNÉRATION DU FICHIER EXCEL DESIGN (.XLSX)
# ---------------------------------------------------------
def generer_excel_moderne(df_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Suivi NC Cuisine"
    ws.views.sheetView[0].showGridLines = True

    # Polices
    font_title = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Segoe UI", size=9, color="1E293B")

    # Couleurs
    fill_title = PatternFill(
        start_color="1E293B", end_color="1E293B", fill_type="solid"
    )
    fill_header = PatternFill(
        start_color="2B579A", end_color="2B579A", fill_type="solid"
    )
    fill_zebra = PatternFill(
        start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
    )

    border_light = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    # Titre principal
    ws.merge_cells("A1:J1")
    title_cell = ws["A1"]
    title_cell.value = " REGISTRE DES NON-CONFORMITÉS HACCP - CUISINE CENTRALE"
    title_cell.font = font_title
    title_cell.fill = fill_title
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 10

    # Colonnes
    headers = list(df_data.columns)
    for col_num, header_title in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.value = header_title.upper()
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    ws.row_dimensions[3].height = 28

    # Lignes de données
    for row_idx, row_data in enumerate(df_data.values, 4):
        ws.row_dimensions[row_idx].height = 26
        is_even = row_idx % 2 == 0

        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = value
            cell.font = font_body
            cell.border = border_light

            if is_even:
                cell.fill = fill_zebra

            if col_idx in [1, 2, 6, 9, 10]:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
            else:
                cell.alignment = Alignment(
                    horizontal="left", vertical="center", wrap_text=True
                )

    # Couleurs de Statut
    last_row = max(len(df_data) + 3, 4)
    ws.conditional_formatting.add(
        f"J4:J{last_row}",
        CellIsRule(
            operator="equal",
            formula=['"Ouvert"'],
            fill=PatternFill(
                start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"
            ),
            font=Font(color="991B1B", bold=True, name="Segoe UI", size=9),
        ),
    )
    ws.conditional_formatting.add(
        f"J4:J{last_row}",
        CellIsRule(
            operator="equal",
            formula=['"En cours"'],
            fill=PatternFill(
                start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"
            ),
            font=Font(color="92400E", bold=True, name="Segoe UI", size=9),
        ),
    )
    ws.conditional_formatting.add(
        f"J4:J{last_row}",
        CellIsRule(
            operator="equal",
            formula=['"Clôturé"'],
            fill=PatternFill(
                start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"
            ),
            font=Font(color="166534", bold=True, name="Segoe UI", size=9),
        ),
    )

    # Largeurs des colonnes
    column_widths = {
        "A": 16,
        "B": 15,
        "C": 26,
        "D": 24,
        "E": 40,
        "F": 16,
        "G": 40,
        "H": 22,
        "I": 15,
        "J": 14,
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


st.subheader("📥 Exporter en Excel")
excel_data = generer_excel_moderne(st.session_state.df_nc)

# Bouton de téléchargement Excel officiel (.xlsx)
st.download_button(
    label="📊 Télécharger le rapport Excel Moderne (.xlsx)",
    data=excel_data,
    file_name="Suivi_Non_Conformites_Cuisine.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
