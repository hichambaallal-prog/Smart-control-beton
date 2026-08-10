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
# 🖨️ FONCTION D'EXPORT EXCEL PROFESSIONNEL (MISE EN PAGE COULEUR & A4)
# ==============================================================================
def generer_excel_recap(df_data, titre_rapport):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Écriture des données à partir de la ligne 6 pour laisser de la place à l'en-tête
    df_data.to_excel(writer, index=False, sheet_name='Recap', startrow=5)
    
    workbook = writer.book
    worksheet = writer.sheets['Recap']
    
    # Configuration Impression A4 Paysage et ajustement automatique à 1 page de large
    worksheet.set_paper(9)  # Format A4
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)  # 1 page de large, hauteur automatique
    worksheet.set_print_scale(85)  # Échelle d'impression confortable
    
    # --- Définition des Styles et Couleurs ---
    fmt_titre = workbook.add_format({
        'bold': True, 
        'font_size': 14, 
        'align': 'center', 
        'valign': 'vcenter',
        'font_color': '#1B365D',
        'border': 1
    })
    
    fmt_sous_titre = workbook.add_format({
        'bold': True, 
        'font_size': 10, 
        'align': 'center', 
        'valign': 'vcenter',
        'font_color': '#555555'
    })
    
    fmt_entete = workbook.add_format({
        'bold': True,
        'font_size': 10,
        'font_color': '#FFFFFF',
        'bg_color': '#1B365D',  # Bleu marine professionnel
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    fmt_cellule = workbook.add_format({
        'font_size': 10,
        'valign': 'vcenter',
        'align': 'center',
        'border': 1
    })
    
    fmt_signature = workbook.add_format({
        'bold': True,
        'font_size': 10,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Insertion d'un en-tête fusionné pour le titre du rapport
    worksheet.merge_range('B2:G2', "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - LGV CASA SUD", fmt_sous_titre)
    worksheet.merge_range('B3:G3', titre_rapport, fmt_titre)
    
    # Application des styles sur les en-têtes du tableau (Ligne 6, index 5)
    for col_num, value in enumerate(df_data.columns.values):
        worksheet.write(5, col_num, value, fmt_entete)
        worksheet.set_row(5, 25) # Hauteur de l'en-tête
        
    # Application des styles sur les cellules de données
    for row_idx in range(len(df_data)):
        worksheet.set_row(6 + row_idx, 20) # Hauteur des lignes de données
        for col_idx in range(len(df_data.columns)):
            valeur_cellule = df_data.iloc[row_idx, col_idx]
            # Gestion propre des valeurs manquantes / NaN
            if pd.isna(valeur_cellule):
                valeur_cellule = ""
            worksheet.write(6 + row_idx, col_idx, valeur_cellule, fmt_cellule)

    # Ajustement automatique de la largeur des colonnes avec une marge de sécurité
    for i, col in enumerate(df_data.columns):
        max_len = max(df_data[col].astype(str).map(len).max(), len(str(col))) + 4
        worksheet.set_column(i, i, max(max_len, 15))

    # Blocs de Signature en bas du document
    derniere_ligne = len(df_data) + 9
    worksheet.merge_range(f'B{derniere_ligne}:C{derniere_ligne}', "Responsable d'Essai LPEE", fmt_signature)
    worksheet.merge_range(f'E{derniere_ligne}:F{derniere_ligne}', "Chef du Laboratoire LGV CASA SUD", fmt_signature)
    
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
        [
            "🏠 Accueil", 
            "🪨 Essai à la Plaque", 
            "🏗️ Suivi de Bétonnage",
            "📊 Synthèse Béton"
        ]
    )
    
    st.write("---")
    if st.button("🚪 Déconnexion", type="secondary", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

# ==============================================================================
# 📄 3. CONTENU DES PAGES
# ==============================================================================

# ------------------------------------------------------------------------------
# PAGE 1 : ACCUEIL
# ------------------------------------------------------------------------------
if page == "🏠 Accueil":
    st.title("👋 Bienvenue sur le Portail de Contrôle Qualité")
    st.subheader("Laboratoire Public d'Essais et d'Études (LPEE)")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("""
        Ce portail vous permet de gérer les essais :
        * **🪨 Essai à la Plaque :** Calculs EV1, EV2 et rapport k.
        * **🏗️ Suivi de Bétonnage :** Gestion des bons de livraison et prélèvements.
        * **📊 Synthèse Béton :** Bilan journalier et mensuel détaillé.
        """)
    with col2:
        st.info("**Projet :** LGV CASA SUD\n\n**Client :** TGCC\n\n**Centrale :** TG PREFA")

# ------------------------------------------------------------------------------
# PAGE 2 : ESSAI À LA PLAQUE
# ------------------------------------------------------------------------------
elif page == "🪨 Essai à la Plaque":
    st.title("🪨 Contrôle de Portance - Essai à la Plaque")
    st.info("Module d'essai à la plaque.") 

# ------------------------------------------------------------------------------
# PAGE 3 : SUIVI DE BÉTONNAGE
# ------------------------------------------------------------------------------
elif page == "🏗️ Suivi de Bétonnage":
    st.title("🏗️ Suivi et Contrôle Qualité Béton")
    
    try:
        resp_beton = supabase.table("suivi_beton").select("*").execute()
        data_all_beton = resp_beton.data or []
    except Exception:
        data_all_beton = []

    date_b = st.date_input("📅 Date de livraison :", value=date.today())
    str_date_b = date_b.strftime("%d/%m/%Y")

    st.subheader(f"📝 Saisie d'un contrôle ({str_date_b})")

    col_h1, col_h2, col_h3 = st.columns(3)
    with col_h1: technicien = st.text_input("👤 Nom du Technicien LPEE", value="Agent LPEE")
    with col_h2: client_b = st.text_input("🏢 Client", value="TGCC", disabled=True)
    with col_h3: centrale_b = st.text_input("🏭 Centrale à Béton", value="TG PREFA")

    st.markdown("---")
    c_b1, c_b2, c_b3 = st.columns(3)
    
    with c_b1:
        bl_num = st.text_input("N° BL", value="BL-2026-001")
        ouvrage = st.text_input("Ouvrage", value="Voile / Semelle")
        quantite_m3 = st.number_input("Quantité (m³)", value=8.0, step=0.5)

    with c_b2:
        t_arrivee = st.time_input("🕒 Heure arrivée", value=datetime.strptime("08:30", "%H:%M").time())
        t_fin = st.time_input("🏁 Heure fin", value=datetime.strptime("09:15", "%H:%M").time())
        classe_b = st.selectbox("Classe", ["C20/25", "C25/30", "C30/37", "C35/45", "C40/50"])

    with c_b3:
        meteo = st.selectbox("Météo", ["Ensoleillé ☀️", "Nuageux ⛅", "Pluie 🌧️", "Vent fort 💨", "Chaleur extrême 🔴"])
        temp_beton = st.number_input("🌡️ Température du béton (°C)", value=20.0, step=0.5)
        temp_ambiante = st.number_input("🌤️ Température ambiante (°C)", value=25.0, step=0.5)
        affaisse = st.number_input("Affaissement (mm)", value=150.0, step=10.0)
        
        prelev = st.selectbox("Prélèvement", ["OUI - Conforme (NF EN 12390-2)", "NON", "Sans objet"])
        is_disabled = (prelev == "NON")
        nb_ep = st.number_input("Nb d'éprouvettes", min_value=0, max_value=12, value=0 if is_disabled else 6, disabled=is_disabled)

    obs_b = st.text_area("Observations", value="Béton conforme")

    if st.button("💾 Enregistrer", type="primary"):
        row_b = {
            "date_livraison": str_date_b, 
            "technicien": technicien, 
            "client": client_b, 
            "centrale_beton": centrale_b,
            "bl_num": bl_num, 
            "ouvrage": ouvrage, 
            "heure_arrivee": t_arrivee.strftime("%H h %M min"), 
            "heure_fin_coulage": t_fin.strftime("%H h %M min"),
            "quantite_m3": float(quantite_m3), 
            "classe_beton": classe_b, 
            "meteo": meteo, 
            "temperature_beton": float(temp_beton), 
            "temperature_ambiante": float(temp_ambiante), 
            "affaissement": float(affaisse), 
            "prelevement": prelev, 
            "nb_eprouvettes": int(nb_ep), 
            "observations": obs_b
        }
        supabase.table("suivi_beton").insert(row_b).execute()
        st.success("✅ Enregistré !")
        st.rerun()

    st.markdown("---")
    st.subheader("📋 Historique")
    if data_all_beton:
        st.dataframe(pd.DataFrame(data_all_beton), use_container_width=True)

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
        
        colonnes_a_afficher = {
            "date_livraison": "Date de suivi",
            "ouvrage": "Partie d'ouvrage",
            "bl_num": "N° de BL",
            "classe_beton": "Classe de béton",
            "affaissement": "Affaissement (mm)",
            "temperature_beton": "Temp. Béton (°C)",
            "temperature_ambiante": "Temp. Ambiante (°C)",
            "meteo": "Météo"
        }
        
        colonnes_disponibles = {k: v for k, v in colonnes_a_afficher.items() if k in df.columns}
        
        classes_uniques = sorted(df['classe_beton'].dropna().unique().tolist()) if 'classe_beton' in df.columns else []
        options_filtre_classe = ["Toutes"] + classes_uniques
        
        # ==========================================
        # ONGLET 1 : BILAN JOURNALIER
        # ==========================================
        with tab_jour:
            st.subheader("Filtrage par jour et par classe de béton")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                d_jour = st.date_input("Sélectionnez une date :", value=date.today(), key="input_date_jour")
            with col_f2:
                classe_filtre_j = st.selectbox("Filtrer par classe de béton :", options_filtre_classe, key="filtre_classe_j")
            
            df_jour = df[df['date_livraison_dt'].dt.date == d_jour]
            
            if classe_filtre_j != "Toutes":
                df_jour = df_jour[df_jour['classe_beton'] == classe_filtre_j]
            
            if df_jour.empty:
                st.info(f"Aucun coulage enregistré pour les critères sélectionnés.")
            else:
                total_vol_jour = df_jour["quantite_m3"].sum()
                total_liv_jour = df_jour["bl_num"].count()
                
                col1, col2 = st.columns(2)
                col1.metric(label="Total Volume Coulé (m³)", value=f"{total_vol_jour:.2f} m³")
                col2.metric(label="Nombre de Toupies / BL", value=total_liv_jour)
                
                st.markdown("#### 📄 Détail des Coulages (Chantier)")
                recap_j_detail = df_jour[list(colonnes_disponibles.keys())].rename(columns=colonnes_disponibles)
                
                st.dataframe(recap_j_detail, use_container_width=True, hide_index=True)
                
                titre_j = f"Recapitulatif Journalier - {d_jour.strftime('%d/%m/%Y')}" + (f" ({classe_filtre_j})" if classe_filtre_j != "Toutes" else "")
                excel_data_j = generer_excel_recap(recap_j_detail, titre_j)
                st.download_button(
                    label="📥 Télécharger le Bilan Journalier en Excel",
                    data=excel_data_j,
                    file_name=f"Recap_Journalier_{d_jour.strftime('%d-%m-%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_excel_j"
                )

        # ==========================================
        # ONGLET 2 : BILAN MENSUEL
        # ==========================================
        with tab_mois:
            st.subheader("Filtrage par mois, année et classe de béton")
            col_m1, col_m2, col_m3 = st.columns(3)
            
            with col_m1:
                mois_choisi = st.selectbox("Mois", range(1, 13), index=date.today().month - 1, key="select_mois")
            with col_m2:
                annee_choisie = st.selectbox("Année", range(2024, 2030), index=2, key="select_annee")
            with col_m3:
                classe_filtre_m = st.selectbox("Filtrer par classe de béton :", options_filtre_classe, key="filtre_classe_m")
                
            df_mois = df[
                (df['date_livraison_dt'].dt.month == mois_choisi) & 
                (df['date_livraison_dt'].dt.year == annee_choisie)
            ]
            
            if classe_filtre_m != "Toutes":
                df_mois = df_mois[df_mois['classe_beton'] == classe_filtre_m]
            
            if df_mois.empty:
                st.info(f"Aucun coulage enregistré pour la période et la classe sélectionnées.")
            else:
                total_vol_mois = df_mois["quantite_m3"].sum()
                total_liv_mois = df_mois["bl_num"].count()
                
                col1, col2 = st.columns(2)
                col1.metric(label="Volume Mensuel Cumulé (m³)", value=f"{total_vol_mois:.2f} m³")
                col2.metric(label="Nombre de Toupies / BL", value=total_liv_mois)
                
                st.markdown(f"#### 📄 Synthèse Détaillée pour {mois_choisi:02d}/{annee_choisie}")
                recap_m_detail = df_mois[list(colonnes_disponibles.keys())].rename(columns=colonnes_disponibles)
                
                st.dataframe(recap_m_detail, use_container_width=True, hide_index=True)
                
                titre_m = f"Recapitulatif Mensuel - {mois_choisi:02d}/{annee_choisie}" + (f" ({classe_filtre_m})" if classe_filtre_m != "Toutes" else "")
                excel_data_m = generer_excel_recap(recap_m_detail, titre_m)
                st.download_button(
                    label="📥 Télécharger le Bilan Mensuel en Excel",
                    data=excel_data_m,
                    file_name=f"Recap_Mensuel_{mois_choisi:02d}_{annee_choisie}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_excel_m"
                )
