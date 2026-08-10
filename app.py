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
MOT_DE_PASSE_ADMIN = "admin2026"  # Mot de passe administrateur pour modifier/supprimer

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if "admin_authenticated" not in st.session_state:
    st.session_state["admin_authenticated"] = False

# ==============================================================================
# 🖨️ FONCTION D'EXPORT EXCEL PROFESSIONNEL (MISE EN PAGE COULEUR & A4 PORTRAIT)
# ==============================================================================
def generer_excel_recap(df_data, titre_rapport):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Écriture des données à partir de la ligne 6 pour laisser de la place à l'en-tête
    df_data.to_excel(writer, index=False, sheet_name='Recap', startrow=5)
    
    workbook = writer.book
    worksheet = writer.sheets['Recap']
    
    # Configuration Impression A4 Portrait et ajustement automatique à 1 page de large
    worksheet.set_paper(9)  # Format A4
    worksheet.set_portrait() # Orientation Portrait
    worksheet.fit_to_pages(1, 0)  # 1 page de large, hauteur automatique
    worksheet.set_print_scale(70)  # Échelle d'impression optimisée pour portrait
    
    # Marges de la page pour l'impression (en pouces)
    worksheet.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    
    # --- Définition des Styles et Couleurs ---
    fmt_titre = workbook.add_format({
        'bold': True, 
        'font_size': 12, 
        'align': 'center', 
        'valign': 'vcenter',
        'font_color': '#1B365D',
        'border': 1,
        'bg_color': '#F2F4F8'
    })
    
    fmt_sous_titre = workbook.add_format({
        'bold': True, 
        'font_size': 9, 
        'align': 'center', 
        'valign': 'vcenter',
        'font_color': '#555555'
    })
    
    fmt_entete = workbook.add_format({
        'bold': True,
        'font_size': 9,
        'font_color': '#FFFFFF',
        'bg_color': '#1B365D',  # Bleu marine professionnel
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True
    })
    
    fmt_cellule = workbook.add_format({
        'font_size': 9,
        'valign': 'vcenter',
        'align': 'center',
        'border': 1,
        'text_wrap': True
    })
    
    fmt_signature = workbook.add_format({
        'bold': True,
        'font_size': 9,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'bg_color': '#FAFAFA'
    })

    # Détermination dynamique de la dernière colonne pour les fusions d'en-tête
    max_col_idx = max(len(df_data.columns) - 1, 1)
    
    # Insertion d'un en-tête fusionné adapté au format portrait
    worksheet.merge_range(1, 0, 1, max_col_idx, "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - LGV CASA SUD", fmt_sous_titre)
    worksheet.merge_range(2, 0, 2, max_col_idx, titre_rapport, fmt_titre)
    
    # Application des styles sur les en-têtes du tableau (Ligne 6, index 5)
    for col_num, value in enumerate(df_data.columns.values):
        worksheet.write(5, col_num, value, fmt_entete)
    worksheet.set_row(5, 28) # Hauteur de l'en-tête du tableau
        
    # Application des styles sur les cellules de données
    for row_idx in range(len(df_data)):
        worksheet.set_row(6 + row_idx, 22) # Hauteur des lignes de données
        for col_idx in range(len(df_data.columns)):
            valeur_cellule = df_data.iloc[row_idx, col_idx]
            if pd.isna(valeur_cellule):
                valeur_cellule = ""
            worksheet.write(6 + row_idx, col_idx, valeur_cellule, fmt_cellule)

    # Ajustement ciblé de la largeur des colonnes pour un affichage parfait en portrait
    for i, col in enumerate(df_data.columns):
        max_len = max(df_data[col].astype(str).map(len).max(), len(str(col))) + 3
        worksheet.set_column(i, i, max(max_len, 11))

    # Blocs de Signature en bas du document proprement répartis
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
            "📊 Synthèse Béton",
            "📈 Synthèse Essai à la Plaque"
        ]
    )
    
    st.write("---")
    if st.button("🚪 Déconnexion", type="secondary", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["admin_authenticated"] = False
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
        * **🪨 Essai à la Plaque :** Saisie des données de chargement, calculs automatiques EV1, EV2 et k.
        * **🏗️ Suivi de Bétonnage :** Module clôturé (historique et consultation).
        * **📊 Synthèse Béton :** Bilan journalier et mensuel détaillé.
        * **📈 Synthèse Essai à la Plaque :** Bilans journaliers et mensuels avec filtrage par type de plateforme et export Excel A4.
        """)
    with col2:
        st.info("**Projet :** LGV CASA SUD\n\n**Client :** TGCC\n\n**Centrale :** TG PREFA")

# ------------------------------------------------------------------------------
# PAGE 2 : ESSAI À LA PLAQUE
# ------------------------------------------------------------------------------
elif page == "🪨 Essai à la Plaque":
    st.title("🪨 Contrôle de Portance - Essai à la Plaque")
    
    try:
        resp_plaque = supabase.table("essais_plaque").select("*").execute()
        data_all_plaque = resp_plaque.data or []
    except Exception:
        data_all_plaque = []

    date_p = st.date_input("📅 Date de l'essai :", value=date.today(), key="date_plaque_input")
    str_date_p = date_p.strftime("%d/%m/%Y")

    st.subheader(f"📝 Saisie d'un Essai à la Plaque ({str_date_p})")

    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        technicien_p = st.text_input("👤 Nom du Technicien LPEE", value="Agent LPEE", key="tech_p")
        client_p = st.text_input("🏢 Client", value="TGCC", disabled=True, key="client_p")
    with col_p2:
        localisation = st.text_input("📍 Localisation / PK / Ouvrage", value="Zone de plateforme PK 0+000", key="loc_p")
        projet_lgv = st.text_input("Projet", value="LGV - CASA SUD", disabled=True, key="projet_lgv_input")
    with col_p3:
        type_plateforme = st.selectbox("Sélectionner type de plateforme", ["Arase", "Remblai", "PST", "Couche de forme"], key="type_plt_sel")

    st.markdown("---")
    st.markdown("#### ⚙️ Paramètres de Chargement & Déformations")
    
    col_z1, col_z2 = st.columns(2)
    with col_z1:
        z1 = st.number_input("Z1 - 1er Chargement (mm)", min_value=0.001, value=1.50, step=0.05, format="%.3f", key="input_z1")
    with col_z2:
        z2 = st.number_input("Z2 - 2ème Chargement (mm)", min_value=0.001, value=1.00, step=0.05, format="%.3f", key="input_z2")

    # Calculs automatiques
    ev1 = 112.5 / (z1 * 2) if z1 > 0 else 0.0
    ev2 = 90.0 / (z2 * 2) if z2 > 0 else 0.0
    K_val = (ev2 / ev1) if ev1 > 0 else 0.0

    st.markdown("---")
    st.markdown("#### 📊 Résultats Calculés")
    res_c1, res_c2, res_c3 = st.columns(3)
    res_c1.metric(label="EV1 (MPa)", value=f"{ev1:.2f} MPa")
    res_c2.metric(label="EV2 (MPa)", value=f"{ev2:.2f} MPa")
    res_c3.metric(label="Coefficient k (EV2/EV1)", value=f"{k_val:.2f}")

    obs_p = st.text_area("Observations / Conformité", value="Plateforme conforme", key="obs_plq")

    if st.button("💾 Enregistrer l'Essai à la Plaque", type="primary", key="btn_save_plaque"):
        row_p = {
            "date_essai": str_date_p,
            "technicien": technicien_p,
            "client": client_p,
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

    st.markdown("---")
    st.subheader("📋 Historique des Essais à la Plaque")
    if data_all_plaque:
        df_hist_p = pd.DataFrame(data_all_plaque)
        if "created_at" in df_hist_p.columns:
            df_hist_p = df_hist_p.drop(columns=["created_at"])
        st.dataframe(df_hist_p, use_container_width=True, hide_index=True)
    else:
        st.info("Aucun essai à la plaque enregistré pour le moment.")

# ------------------------------------------------------------------------------
# PAGE 3 : SUIVI DE BÉTONNAGE (CLÔTURÉ)
# ------------------------------------------------------------------------------
elif page == "🏗️ Suivi de Bétonnage":
    st.title("🏗️ Suivi et Contrôle Qualité Béton")
    st.info("🔒 Ce module est désormais **clôturé** en saisie. Vous pouvez consulter l'historique ci-dessous ou utiliser le module **Synthèse Béton** pour les bilans.")

    try:
        resp_beton = supabase.table("suivi_beton").select("*").execute()
        data_all_beton = resp_beton.data or []
    except Exception:
        data_all_beton = []

    if data_all_beton:
        df_hist = pd.DataFrame(data_all_beton)
        if "created_at" in df_hist.columns:
            df_hist = df_hist.drop(columns=["created_at"])
        st.dataframe(df_hist, use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ Aucune donnée de bétonnage enregistrée.")

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

# ------------------------------------------------------------------------------
# PAGE 5 : SYNTHÈSE ESSAI À LA PLAQUE
# ------------------------------------------------------------------------------
elif page == "📈 Synthèse Essai à la Plaque":
    st.title("📈 Récapitulatif et Synthèse - Essai à la Plaque")
    
    try:
        resp_plaque = supabase.table("essais_plaque").select("*").execute()
        data_all_plaque = resp_plaque.data or []
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        data_all_plaque = []

    if not data_all_plaque:
        st.warning("⚠️ Aucun essai à la plaque n'est encore enregistré.")
    else:
        df_p = pd.DataFrame(data_all_plaque)
        df_p['date_essai_dt'] = pd.to_datetime(df_p['date_essai'], format='%d/%m/%Y', errors='coerce')
        
        tab_jour_p, tab_mois_p = st.tabs(["📅 Bilan Journalier (Plaque)", "📆 Bilan Mensuel (Plaque)"])
        
        colonnes_plaque_afficher = {
            "date_essai": "Date de l'essai",
            "technicien": "Technicien",
            "client": "Client",
            "localisation": "Localisation / PK",
            "projet": "Projet",
            "type_plateforme": "Type de plateforme",
            "z1": "Z1 (mm)",
            "z2": "Z2 (mm)",
            "ev1": "EV1 (MPa)",
            "ev2": "EV2 (MPa)",
            "k": "Coefficient k",
            "observations": "Observations"
        }
        
        cols_disp_p = {k: v for k, v in colonnes_plaque_afficher.items() if k in df_p.columns}
        
        plateformes_uniques = sorted(df_p['type_plateforme'].dropna().unique().tolist()) if 'type_plateforme' in df_p.columns else []
        options_filtre_plt = ["Toutes"] + plateformes_uniques
        
        # ==========================================
        # ONGLET 1 : BILAN JOURNALIER (PLAQUE)
        # ==========================================
        with tab_jour_p:
            st.subheader("Filtrage par jour et par type de plateforme")
            col_fp1, col_fp2 = st.columns(2)
            with col_fp1:
                d_jour_p = st.date_input("Sélectionnez une date :", value=date.today(), key="input_date_jour_plaque")
            with col_fp2:
                plt_filtre_j = st.selectbox("Filtrer par type de plateforme :", options_filtre_plt, key="filtre_plt_j")
            
            df_p_jour = df_p[df_p['date_essai_dt'].dt.date == d_jour_p]
            
            if plt_filtre_j != "Toutes":
                df_p_jour = df_p_jour[df_p_jour['type_plateforme'] == plt_filtre_j]
            
            if df_p_jour.empty:
                st.info("Aucun essai à la plaque enregistré pour les critères sélectionnés.")
            else:
                total_essais_j = len(df_p_jour)
                col1_j, col2_j = st.columns(2)
                col1_j.metric(label="Nombre d'Essais Réalisés", value=total_essais_j)
                if 'ev2' in df_p_jour.columns and not df_p_jour['ev2'].empty:
                    col2_j.metric(label="EV2 Moyen (MPa)", value=f"{df_p_jour['ev2'].mean():.2f} MPa")
                
                st.markdown("#### 📄 Détail des Essais (Journalier)")
                recap_p_j_detail = df_p_jour[list(cols_disp_p.keys())].rename(columns=cols_disp_p)
                
                st.dataframe(recap_p_j_detail, use_container_width=True, hide_index=True)
                
                titre_pj = f"Recapitulatif Journalier - Essai Plaque - {d_jour_p.strftime('%d/%m/%Y')}" + (f" ({plt_filtre_j})" if plt_filtre_j != "Toutes" else "")
                excel_data_pj = generer_excel_recap(recap_p_j_detail, titre_pj)
                st.download_button(
                    label="📥 Télécharger le Bilan Journalier Plaque en Excel",
                    data=excel_data_pj,
                    file_name=f"Recap_Plaque_Journalier_{d_jour_p.strftime('%d-%m-%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_excel_pj"
                )

        # ==========================================
        # ONGLET 2 : BILAN MENSUEL (PLAQUE)
        # ==========================================
        with tab_mois_p:
            st.subheader("Filtrage par mois, année et type de plateforme")
            col_mp1, col_mp2, col_mp3 = st.columns(3)
            
            with col_mp1:
                mois_choisi_p = st.selectbox("Mois", range(1, 13), index=date.today().month - 1, key="select_mois_p")
            with col_mp2:
                annee_choisie_p = st.selectbox("Année", range(2024, 2030), index=2, key="select_annee_p")
            with col_mp3:
                plt_filtre_m = st.selectbox("Filtrer par type de plateforme :", options_filtre_plt, key="filtre_plt_m")
                
            df_p_mois = df_p[
                (df_p['date_essai_dt'].dt.month == mois_choisi_p) & 
                (df_p['date_essai_dt'].dt.year == annee_choisie_p)
            ]
            
            if plt_filtre_m != "Toutes":
                df_p_mois = df_p_mois[df_p_mois['type_plateforme'] == plt_filtre_m]
            
            if df_p_mois.empty:
                st.info("Aucun essai à la plaque enregistré pour la période et la plateforme sélectionnées.")
            else:
                total_essais_m = len(df_p_mois)
                col1_m, col2_m = st.columns(2)
                col1_m.metric(label="Nombre Total d'Essais (Mensuel)", value=total_essais_m)
                if 'ev2' in df_p_mois.columns and not df_p_mois['ev2'].empty:
                    col2_m.metric(label="EV2 Moyen (MPa)", value=f"{df_p_mois['ev2'].mean():.2f} MPa")
                
                st.markdown(f"#### 📄 Synthèse Détaillée pour {mois_choisi_p:02d}/{annee_choisie_p}")
                recap_p_m_detail = df_p_mois[list(cols_disp_p.keys())].rename(columns=cols_disp_p)
                
                st.dataframe(recap_p_m_detail, use_container_width=True, hide_index=True)
                
                titre_pm = f"Recapitulatif Mensuel - Essai Plaque - {mois_choisi_p:02d}/{annee_choisie_p}" + (f" ({plt_filtre_m})" if plt_filtre_m != "Toutes" else "")
                excel_data_pm = generer_excel_recap(recap_p_m_detail, titre_pm)
                st.download_button(
                    label="📥 Télécharger le Bilan Mensuel Plaque en Excel",
                    data=excel_data_pm,
                    file_name=f"Recap_Plaque_Mensuel_{mois_choisi_p:02d}_{annee_choisie_p}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_excel_pm"
                )
