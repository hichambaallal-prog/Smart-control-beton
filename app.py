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
    df_data.to_excel(writer, index=False, sheet_name='Recap', startrow=7)
    
    workbook = writer.book
    worksheet = writer.sheets['Recap']
    
    # Configuration A4 Paysage
    worksheet.set_paper(9)
    worksheet.set_landscape()
    
    # Formats
    fmt_titre = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
    
    # Tentative d'insertion du logo
    try:
        worksheet.insert_image('A1', 'logo_lpee.png', {'x_scale': 0.4, 'y_scale': 0.4})
    except:
        pass
    
    worksheet.merge_range('C2:E2', titre_rapport, fmt_titre)
    
    # Signature
    derniere_ligne = len(df_data) + 12
    worksheet.write(f'B{derniere_ligne}', "Responsable d'essai")
    worksheet.write(f'E{derniere_ligne}', "Chef du laboratoire")
    
    writer.close()
    return output.getvalue()

# --- ÉCRAN DE CONNEXION ---
if not st.session_state["authenticated"]:
    col_g, col_c, col_d = st.columns([1, 2, 1])
    
    with col_c:
        url_image_al_boraq = "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/ONCF_Al_boraq.jpeg/1280px-ONCF_Al_boraq.jpeg"
        st.image(url_image_al_boraq, caption="Projet LGV CASA SUD - LPEE CTR-CSB", use_container_width=True)
        
        st.title("🔒 Connexion au Portail")
        st.markdown("##### **LPEE - CTR-CSB** | Projet : **LGV CASA SUD** | Client : **TGCC**")
        st.markdown("---")
        
        pwd_input = st.text_input("Veuillez saisir le mot de passe :", type="password")
        
        if st.button("Se connecter", type="primary", use_container_width=True):
            if pwd_input == MOT_DE_PASSE_ACCES:
                st.session_state["authenticated"] = True
                st.success("Accès autorisé !")
                st.rerun()
            else:
                st.error("❌ Mot de passe incorrect.")
    st.stop()

# ==============================================================================
# ⚙️ 2. CONNEXION SUPABASE & MENU LATÉRAL
# ==============================================================================
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0"

try:
    supabase: Client = create_client(URL, CLE)
except Exception as e:
    st.error(f"Erreur Supabase : {e}")
    st.stop()

with st.sidebar:
    st.title("🏢 LPEE - CTR-CSB")
    st.caption("Projet : **LGV CASA SUD** | Client : **TGCC**")
    st.write("---")
    
    page = st.radio(
        "📌 Menu Principal",
        ["🏠 Accueil", "🪨 Essai à la Plaque", "🏗️ Suivi de Bétonnage", "📊 Synthèse Béton"]
    )
    
    st.write("---")
    if st.button("🚪 Déconnexion", type="secondary", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

# ------------------------------------------------------------------------------
# PAGE 4 : SYNTHÈSE BÉTON
# ------------------------------------------------------------------------------
elif page == "📊 Synthèse Béton":
    st.title("📊 Récapitulatif et Synthèse du Bétonnage")
    
    try:
        resp_beton = supabase.table("suivi_beton").select("*").execute()
        data_all_beton = resp_beton.data or []
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        data_all_beton = []

    if not data_all_beton:
        st.warning("⚠️ Aucune donnée de bétonnage n'est encore enregistrée.")
    else:
        df = pd.DataFrame(data_all_beton)
        df['date_livraison_dt'] = pd.to_datetime(df['date_livraison'], format='%d/%m/%Y', errors='coerce')
        
        tab_jour, tab_mois = st.tabs(["📅 Bilan Journalier", "📆 Bilan Mensuel"])
        
        # Dictionnaire mis à jour pour la colonne Température ambiante
        colonnes_a_afficher = {
            "date_livraison": "Date de suivi",
            "ouvrage": "Partie d'ouvrage",
            "bl_num": "N° de BL",
            "affaissement": "Affaissement (cm)",
            "temperature": "Temp. Béton (°C)",
            "temperature": "Température ambiante (°C)"
        }
        
        with tab_jour:
            st.subheader("Filtrage par jour")
            d_jour = st.date_input("Sélectionnez une date :", value=date.today(), key="input_date_jour")
            df_jour = df[df['date_livraison_dt'].dt.date == d_jour]
            
            if df_jour.empty:
                st.info(f"Aucun coulage enregistré pour le {d_jour.strftime('%d/%m/%Y')}.")
            else:
                st.markdown("#### 📄 Détail des Coulages (Chantier)")
                recap_j_detail = df_jour[list(colonnes_a_afficher.keys())].rename(columns=colonnes_a_afficher)
                st.dataframe(recap_j_detail, use_container_width=True, hide_index=True)
                
                titre_j = f"Recapitulatif Journalier - {d_jour.strftime('%d/%m/%Y')}"
                excel_data_j = generer_excel_recap(recap_j_detail, titre_j)
                st.download_button("📥 Télécharger le Bilan Journalier en Excel", data=excel_data_j, file_name=f"Recap_Journalier_{d_jour.strftime('%d-%m-%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with tab_mois:
            st.subheader("Filtrage par mois et année")
            col_m1, col_m2 = st.columns(2)
            with col_m1: mois_choisi = st.selectbox("Mois", range(1, 13), index=date.today().month - 1, key="select_mois")
            with col_m2: annee_choisie = st.selectbox("Année", range(2024, 2030), index=2, key="select_annee")
                
            df_mois = df[(df['date_livraison_dt'].dt.month == mois_choisi) & (df['date_livraison_dt'].dt.year == annee_choisie)]
            
            if df_mois.empty:
                st.info(f"Aucun coulage enregistré pour la période {mois_choisi:02d}/{annee_choisie}.")
            else:
                st.markdown(f"#### 📄 Synthèse Détaillée pour {mois_choisi:02d}/{annee_choisie}")
                recap_m_detail = df_mois[list(colonnes_a_afficher.keys())].rename(columns=colonnes_a_afficher)
                st.dataframe(recap_m_detail, use_container_width=True, hide_index=True)
                
                titre_m = f"Recapitulatif Mensuel - {mois_choisi:02d}/{annee_choisie}"
                excel_data_m = generer_excel_recap(recap_m_detail, titre_m)
                st.download_button("📥 Télécharger le Bilan Mensuel en Excel", data=excel_data_m, file_name=f"Recap_Mensuel_{mois_choisi:02d}_{annee_choisie}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
