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
MOT_DE_PASSE_ADMIN = "admin2026"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if "admin_authenticated" not in st.session_state:
    st.session_state["admin_authenticated"] = False

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
    worksheet.set_print_scale(70)
    worksheet.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    
    fmt_titre = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'font_color': '#1B365D', 'border': 1, 'bg_color': '#F2F4F8'})
    fmt_sous_titre = workbook.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'font_color': '#555555'})
    fmt_entete = workbook.add_format({'bold': True, 'font_size': 9, 'font_color': '#FFFFFF', 'bg_color': '#1B365D', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    fmt_cellule = workbook.add_format({'font_size': 9, 'valign': 'vcenter', 'align': 'center', 'border': 1, 'text_wrap': True})
    fmt_signature = workbook.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FAFAFA'})

    max_col_idx = max(len(df_data.columns) - 1, 1)
    worksheet.merge_range(1, 0, 1, max_col_idx, "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - LGV CASA SUD", fmt_sous_titre)
    worksheet.merge_range(2, 0, 2, max_col_idx, titre_rapport, fmt_titre)
    
    for col_num, value in enumerate(df_data.columns.values):
        worksheet.write(5, col_num, value, fmt_entete)
    worksheet.set_row(5, 28)
        
    for row_idx in range(len(df_data)):
        worksheet.set_row(6 + row_idx, 22)
        for col_idx in range(len(df_data.columns)):
            valeur_cellule = df_data.iloc[row_idx, col_idx]
            worksheet.write(6 + row_idx, col_idx, "" if pd.isna(valeur_cellule) else str(valeur_cellule), fmt_cellule)

    for i, col in enumerate(df_data.columns):
        max_len = max(df_data[col].astype(str).map(len).max(), len(str(col))) + 3
        worksheet.set_column(i, i, max(max_len, 11))

    derniere_ligne = len(df_data) + 9
    milieu_col = max_col_idx // 2
    worksheet.merge_range(derniere_ligne, 0, derniere_ligne, min(milieu_col, max_col_idx), "Responsable d'essai", fmt_signature)
    worksheet.merge_range(derniere_ligne, milieu_col + 1, derniere_ligne, max_col_idx, "Chef du laboratoire", fmt_signature)
    
    writer.close()
    return output.getvalue()

# --- ÉCRAN DE CONNEXION ---
if not st.session_state["authenticated"]:
    col_g, col_c, col_d = st.columns([1, 2, 1])
    with col_c:
        st.title("🔒 Connexion au Portail")
        pwd_input = st.text_input("Veuillez saisir le mot de passe :", type="password")
        if st.button("Se connecter", type="primary"):
            if pwd_input == MOT_DE_PASSE_ACCES:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Mot de passe incorrect.")
    st.stop()

# ==============================================================================
# ⚙️ 2. CONNEXION SUPABASE
# ==============================================================================
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0"
supabase: Client = create_client(URL, CLE)

# --- MENU ---
with st.sidebar:
    st.title("🏢 LPEE - CTR-CSB")
    page = st.radio("📌 Menu Principal", ["🏠 Accueil", "🪨 Essai à la Plaque", "🏗️ Suivi de Bétonnage", "📊 Synthèse Béton", "📈 Synthèse Essai à la Plaque"])
    if st.button("🚪 Déconnexion"):
        st.session_state["authenticated"] = False
        st.rerun()

# ==============================================================================
# 📄 3. CONTENU DES PAGES
# ==============================================================================

# ----------------- PAGE : ESSAI À LA PLAQUE -----------------
if page == "🪨 Essai à la Plaque":
    st.title("🪨 Contrôle de Portance - Essai à la Plaque")
    
    date_p = st.date_input("📅 Date de l'essai :", value=date.today())
    str_date_p = date_p.strftime("%d/%m/%Y")

    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        technicien_p = st.text_input("👤 Nom du Technicien LPEE", value="Agent LPEE")
    with col_p2:
        localisation = st.text_input("📍 Localisation / PK / Ouvrage", value="Zone de plateforme PK 0+000")
        projet_lgv = st.text_input("Projet", value="LGV - CASA SUD", disabled=True)
    with col_p3:
        type_plateforme = st.selectbox("Sélectionner type de plateforme", ["Arase", "Remblai", "PST", "Couche de forme"])

    z1 = st.number_input("Z1 - 1er Chargement (mm)", min_value=0.001, value=1.50, step=0.05, format="%.3f")
    z2 = st.number_input("Z2 - 2ème Chargement (mm)", min_value=0.001, value=1.00, step=0.05, format="%.3f")

    ev1 = 112.5 / (z1 * 2) if z1 > 0 else 0.0
    ev2 = 90.0 / (z2 * 2) if z2 > 0 else 0.0
    k_val = (ev2 / ev1) if ev1 > 0 else 0.0

    obs_p = st.text_area("Observations / Conformité", value="Plateforme conforme")

    if st.button("💾 Enregistrer l'Essai à la Plaque", type="primary"):
        # Dictionnaire corrigé (clé 'client' supprimée pour éviter l'erreur)
        row_p = {
            "date_essai": str_date_p,
            "technicien": technicien_p,
            "localisation": localisation,
            "projet": projet_lgv,
            "type_plateforme": type_plateforme,
            "z1": float(z1),
            "z2": float(z2),
            "ev1": float(ev1),
            "ev2": float(ev2),
            "k": float(k_val),
            "observations": obs_p
        }
        try:
            supabase.table("essais_plaque").insert(row_p).execute()
            st.success("✅ Essai à la plaque enregistré avec succès !")
            st.rerun()
        except Exception as e:
            st.error(f"Erreur lors de l'enregistrement dans Supabase : {e}")

# ... (Le reste de vos pages reste identique)
