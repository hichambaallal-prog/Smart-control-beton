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
        affaisse = st.number_input("Affaissement (cm)", value=15.0)
        
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
        
        # Ajout de la classe de béton dans les colonnes à afficher et exporter
        colonnes_a_afficher = {
            "date_livraison": "Date de suivi",
            "ouvrage": "Partie d'ouvrage",
            "bl_num": "N° de BL",
            "classe_beton": "Classe de béton",
            "affaissement": "Affaissement (cm)",
            "temperature_beton": "Temp. Béton (°C)",
            "temperature_ambiante": "Temp. Ambiante (°C)",
            "meteo": "Météo"
        }
        
        colonnes_disponibles = {k: v for k, v in colonnes_a_afficher.items() if k in df.columns}
        
        # ==========================================
        # ONGLET 1 : BILAN JOURNALIER
        # ==========================================
        with tab_jour:
            st.subheader("Filtrage par jour")
            d_jour = st.date_input("Sélectionnez une date :", value=date.today(), key="input_date_jour")
            
            df_jour = df[df['date_livraison_dt'].dt.date == d_jour]
            
            if df_jour.empty:
                st.info(f"Aucun coulage enregistré pour le {d_jour.strftime('%d/%m/%Y')}.")
            else:
                total_vol_jour = df_jour["quantite_m3"].sum()
                total_liv_jour = df_jour["bl_num"].count()
                
                col1, col2 = st.columns(2)
                col1.metric(label="Total Volume Coulé (m³)", value=f"{total_vol_jour:.2f} m³")
                col2.metric(label="Nombre de Toupies / BL", value=total_liv_jour)
                
                st.markdown("#### 📄 Détail des Coulages (Chantier)")
                recap_j_detail = df_jour[list(colonnes_disponibles.keys())].rename(columns=colonnes_disponibles)
                
                st.dataframe(recap_j_detail, use_container_width=True, hide_index=True)
                
                # Bouton Excel Journalier
                titre_j = f"Recapitulatif Journalier - {d_jour.strftime('%d/%m/%Y')}"
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
            st.subheader("Filtrage par mois et année")
            col_m1, col_m2 = st.columns(2)
            
            with col_m1:
                mois_choisi = st.selectbox("Mois", range(1, 13), index=date.today().month - 1, key="select_mois")
            with col_m2:
                annee_choisie = st.selectbox("Année", range(2024, 2030), index=2, key="select_annee")
                
            df_mois = df[
                (df['date_livraison_dt'].dt.month == mois_choisi) & 
                (df['date_livraison_dt'].dt.year == annee_choisie)
            ]
            
            if df_mois.empty:
                st.info(f"Aucun coulage enregistré pour la période {mois_choisi:02d}/{annee_choisie}.")
            else:
                total_vol_mois = df_mois["quantite_m3"].sum()
                total_liv_mois = df_mois["bl_num"].count()
                
                col1, col2 = st.columns(2)
                col1.metric(label="Volume Mensuel Cumulé (m³)", value=f"{total_vol_mois:.2f} m³")
                col2.metric(label="Nombre de Toupies / BL", value=total_liv_mois)
                
                st.markdown(f"#### 📄 Synthèse Détaillée pour {mois_choisi:02d}/{annee_choisie}")
                recap_m_detail = df_mois[list(colonnes_disponibles.keys())].rename(columns=colonnes_disponibles)
                
                st.dataframe(recap_m_detail, use_container_width=True, hide_index=True)
                
                # Bouton Excel Mensuel
                titre_m = f"Recapitulatif Mensuel - {mois_choisi:02d}/{annee_choisie}"
                excel_data_m = generer_excel_recap(recap_m_detail, titre_m)
                st.download_button(
                    label="📥 Télécharger le Bilan Mensuel en Excel",
                    data=excel_data_m,
                    file_name=f"Recap_Mensuel_{mois_choisi:02d}_{annee_choisie}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_excel_m"
                )
