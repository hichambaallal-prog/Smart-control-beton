import streamlit as st
import pandas as pd
from datetime import date, datetime
from supabase import create_client, Client
import io

# Configuration globale de la page
st.set_page_config(page_title="LPEE CTR-CSB - LGV CASA SUD", layout="wide")

# ==============================================================================
# 🔐 1. GESTION DU MOT DE PASSE ET AUTHENTIFICATION
# ==============================================================================
MOT_DE_PASSE_ACCES = "lpee2026"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# ==============================================================================
# 🖨️ FONCTION D'EXPORT EXCEL PROFESSIONNEL
# ==============================================================================
def generer_excel_recap(df_data, titre_rapport):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_data.to_excel(writer, index=False, sheet_name='Recap', startrow=5)
    
    workbook = writer.book
    worksheet = writer.sheets['Recap']
    worksheet.set_paper(9)
    worksheet.set_portrait()
    worksheet.fit_to_pages(1, 0)
    worksheet.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    
    # Styles
    fmt_titre = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'font_color': '#1B365D', 'border': 1, 'bg_color': '#F2F4F8'})
    fmt_entete = workbook.add_format({'bold': True, 'font_size': 9, 'font_color': '#FFFFFF', 'bg_color': '#1B365D', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_cellule = workbook.add_format({'font_size': 9, 'valign': 'vcenter', 'align': 'center', 'border': 1})

    max_col_idx = max(len(df_data.columns) - 1, 1)
    worksheet.merge_range(2, 0, 2, max_col_idx, titre_rapport, fmt_titre)
    
    for col_num, value in enumerate(df_data.columns.values):
        worksheet.write(5, col_num, value, fmt_entete)
        
    for row_idx in range(len(df_data)):
        for col_idx in range(len(df_data.columns)):
            worksheet.write(6 + row_idx, col_idx, str(df_data.iloc[row_idx, col_idx]), fmt_cellule)

    writer.close()
    return output.getvalue()

# --- AUTH ---
if not st.session_state["authenticated"]:
    col_c = st.columns([1, 2, 1])[1]
    with col_c:
        st.title("🔒 Connexion")
        pwd_input = st.text_input("Mot de passe :", type="password")
        if st.button("Se connecter"):
            if pwd_input == MOT_DE_PASSE_ACCES:
                st.session_state["authenticated"] = True
                st.rerun()
    st.stop()

# ==============================================================================
# ⚙️ CONNEXION SUPABASE
# ==============================================================================
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0"
supabase: Client = create_client(URL, CLE)

# ==============================================================================
# PAGE : ESSAI À LA PLAQUE (CORRIGÉE)
# ==============================================================================
page = st.sidebar.radio("Menu", ["Essai à la Plaque", "Synthèse Essai à la Plaque"])

if page == "Essai à la Plaque":
    st.title("🪨 Saisie Essai à la Plaque")
    date_p = st.date_input("Date", value=date.today())
    technicien_p = st.text_input("Technicien", value="Agent LPEE")
    localisation = st.text_input("Localisation")
    projet_lgv = st.text_input("Projet", value="LGV - CASA SUD")
    type_plateforme = st.selectbox("Type", ["Arase", "Remblai", "PST", "Couche de forme"])
    z1 = st.number_input("Z1", value=1.50)
    z2 = st.number_input("Z2", value=1.00)
    obs_p = st.text_area("Observations")

    if st.button("Enregistrer"):
        # CORRECTION : Suppression de la clé 'client' qui causait l'erreur
        row_p = {
            "date_essai": date_p.strftime("%d/%m/%Y"),
            "technicien": technicien_p,
            "localisation": localisation,
            "projet": projet_lgv,
            "type_plateforme": type_plateforme,
            "z1": float(z1),
            "z2": float(z2),
            "ev1": float(112.5 / (z1 * 2)),
            "ev2": float(90.0 / (z2 * 2)),
            "k": float((90.0 / (z2 * 2)) / (112.5 / (z1 * 2))),
            "observations": obs_p
        }
        try:
            supabase.table("essais_plaque").insert(row_p).execute()
            st.success("✅ Enregistré !")
        except Exception as e:
            st.error(f"Erreur : {e}")

# ... (Le reste du code pour la page "Synthèse" reste identique)
