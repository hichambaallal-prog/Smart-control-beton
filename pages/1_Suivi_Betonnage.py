import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Titre de l'application
st.title("📋 Suivi et Contrôle Qualité Béton")

# --- FORMULAIRE DE SAISIE ---
with st.form("form_saisie_beton"):
    st.subheader("Saisie d'un nouveau contrôle sur site")
    
    col1, col2 = st.columns(2)
    
    with col1:
        date_livraison = st.date_input("Date de livraison")
        bon_livraison = st.text_input("N° Bon de Livraison (BL)")
        # Utilisation de la dénomination exacte "Classe béton"
        classe_beton = st.selectbox(
            "Classe béton", 
            ["C20/25", "C25/30", "C30/37", "C35/45", "C40/50"]
        )
        element_ouvrage = st.text_input("Élément d'ouvrage / Emplacement")

    with col2:
        temperature = st.number_input("Température du béton (°C)", min_value=0.0, max_value=50.0, value=20.0)
        affaissement = st.number_input("Affaissement / Slump (cm)", min_value=0.0, max_value=30.0, value=15.0)
        
        # --- MISE À JOUR : Case Prélèvement avec référence NF EN 12390-2 ---
        prelevement_effectue = st.selectbox(
            "Prélèvement : NF EN 12390-2", 
            ["OUI - Conforme (NF EN 12390-2)", "NON", "Sans objet"]
        )
        nb_eprouvettes = st.number_input("Nombre d'éprouvettes confectionnées", min_value=0, max_value=12, value=6)

    submit = st.form_submit_button("Enregistrer le contrôle")

if submit:
    st.success("Données enregistrées avec succès !")

# --- EXEMPLE DE DATAFRAME ET TABLEAU DE SUIVI ---
data = {
    "Date": ["2026-08-10"],
    "N° BL": ["BL-2026-0891"],
    "Classe béton": ["C30/37"],
    "Élément d'ouvrage": ["Voile P2 - Niveau R+1"],
    "Température (°C)": [21.5],
    "Affaissement (cm)": [16.0],
    "Prélèvement : NF EN 12390-2": [f"{prelevement_effectue} ({nb_eprouvettes} éprouvettes)"] if submit else ["OUI - Conforme (NF EN 12390-2) (6 éprouvettes)"]
}

df = pd.DataFrame(data)

st.subheader("📊 Tableau de suivi")
st.dataframe(df)

# --- FONCTION DE GÉNÉRATION DU RAPPORT EXCEL STYLISÉ ---
def generer_excel_rapport(df_data):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Suivi Béton"

    # Styles Excel
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Titre du rapport
    ws.merge_cells("A1:G1")
    ws["A1"] = "RAPPORT DE SUIVI ET PRÉLÈVEMENT DU BÉTON (NF EN 12390-2)"
    ws["A1"].font = Font(size=13, bold=True, color="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # En-têtes de colonnes mis à jour
    headers = [
        "Date", 
        "N° BL", 
        "Classe béton", 
        "Élément d'ouvrage", 
        "Température (°C)", 
        "Affaissement (cm)", 
        "Prélèvement : NF EN 12390-2"  # Référence exacte ajoutée
    ]
    
    ws.append([])  # Ligne 2 vide
    ws.append(headers)  # Ligne 3 : En-têtes

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Inscription des lignes de données
    for _, row in df_data.iterrows():
        ws.append([
            row["Date"],
            row["N° BL"],
            row["Classe béton"],
            row["Élément d'ouvrage"],
            row["Température (°C)"],
            row["Affaissement (cm)"],
            row["Prélèvement : NF EN 12390-2"]
        ])

    # Application des bordures et ajustement automatique de la largeur des colonnes
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border
            if cell.row > 3:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 18)

    wb.save(output)
    return output.getvalue()

# Bouton de téléchargement du fichier Excel
excel_bytes = generer_excel_rapport(df)
st.download_button(
    label="📥 Télécharger le Rapport Excel (.xlsx)",
    data=excel_bytes,
    file_name="Rapport_Suivi_Prelevement_Beton.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
