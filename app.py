import streamlit as st
import pandas as pd
from datetime import date, datetime
from supabase import create_client, Client
import streamlit as st
import pandas as pd
from datetime import date, datetime
from supabase import create_client, Client
import io  # <--- AJOUTEZ-LE ICI
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
    
    # Tentative d'insertion du logo (doit s'appeler logo_lpee.png dans le dossier GitHub)
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
            "📊 Synthèse Béton"  # <--- NOUVEAU MENU AJOUTÉ ICI
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
        * **📊 Synthèse Béton :** Bilan journalier et mensuel par classe de béton.
        """)
    with col2:
        st.info("**Projet :** LGV CASA SUD\n\n**Client :** TGCC\n\n**Centrale :** TG PREFA")

# ------------------------------------------------------------------------------
# PAGE 2 : ESSAI À LA PLAQUE (Code inchangé)
# ------------------------------------------------------------------------------
elif page == "🪨 Essai à la Plaque":
    st.title("🪨 Contrôle de Portance - Essai à la Plaque")
    # ... (le code reste identique à la version précédente, je le raccourcis visuellement pour toi ici mais tu peux copier ton bloc si besoin, ou utiliser ce qui suit)
    st.info("Module d'essai à la plaque (Conservez votre code précédent pour cette page).") 
    # NOTE: Pour ne pas faire un message trop long, j'ai masqué ce bloc. Tu as déjà le bon code pour cette page.

# ------------------------------------------------------------------------------
# PAGE 3 : SUIVI DE BÉTONNAGE (Code inchangé, ajout de la gestion des données)
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
        temp = st.number_input("Temp (°C)", value=20.0)
        affaisse = st.number_input("Affaissement (cm)", value=15.0)
        
        prelev = st.selectbox("Prélèvement", ["OUI - Conforme (NF EN 12390-2)", "NON", "Sans objet"])
        is_disabled = (prelev == "NON")
        nb_ep = st.number_input("Nb d'éprouvettes", min_value=0, max_value=12, value=0 if is_disabled else 6, disabled=is_disabled)

    obs_b = st.text_area("Observations", value="Béton conforme")

    if st.button("💾 Enregistrer", type="primary"):
        row_b = {
            "date_livraison": str_date_b, "technicien": technicien, "client": client_b, "centrale_beton": centrale_b,
            "bl_num": bl_num, "ouvrage": ouvrage, "heure_arrivee": t_arrivee.strftime("%H h %M min"), "heure_fin_coulage": t_fin.strftime("%H h %M min"),
            "quantite_m3": float(quantite_m3), "classe_beton": classe_b, "meteo": meteo, "temperature": float(temp),
            "affaissement": float(affaisse), "prelevement": prelev, "nb_eprouvettes": int(nb_ep), "observations": obs_b
        }
        supabase.table("suivi_beton").insert(row_b).execute()
        st.success("✅ Enregistré !")
        st.rerun()

    st.markdown("---")
    st.subheader("📋 Historique")
    if data_all_beton:
        st.dataframe(pd.DataFrame(data_all_beton), use_container_width=True)

# ------------------------------------------------------------------------------
# PAGE 4 : NOUVELLE PAGE - SYNTHÈSE BÉTON
# ------------------------------------------------------------------------------
elif page == "📊 Synthèse Béton":
    st.title("📊 Récapitulatif et Synthèse du Bétonnage")
    
    # 1. Récupération des données depuis Supabase
    try:
        resp_beton = supabase.table("suivi_beton").select("*").execute()
        data_all_beton = resp_beton.data or []
    except Exception as e:
        st.error(f"Erreur de chargement : {e}")
        data_all_beton = []

    if not data_all_beton:
        st.warning("⚠️ Aucune donnée de bétonnage n'est encore enregistrée.")
    else:
        # 2. Nettoyage et conversion avec Pandas
        df = pd.DataFrame(data_all_beton)
        # Convertir la colonne 'date_livraison' (format texte "DD/MM/YYYY") en vrai format Date
        df['date_livraison_dt'] = pd.to_datetime(df['date_livraison'], format='%d/%m/%Y', errors='coerce')
        
        # 3. Création de deux onglets : Journalier / Mensuel
        tab_jour, tab_mois = st.tabs(["📅 Bilan Journalier", "📆 Bilan Mensuel"])
        
        # ==========================================
        # ONGLET 1 : BILAN JOURNALIER
        # ==========================================
        with tab_jour:
            st.subheader("Filtrage par jour")
            d_jour = st.date_input("Sélectionnez une date :", value=date.today())
            
            # Filtrer le DataFrame pour la date choisie
            df_jour = df[df['date_livraison_dt'].dt.date == d_jour]
            
            if df_jour.empty:
                st.info(f"Aucun coulage enregistré pour le {d_jour.strftime('%d/%m/%Y')}.")
            else:
                # Calculs des totaux de la journée
                total_vol_jour = df_jour["quantite_m3"].sum()
                total_liv_jour = df_jour["bl_num"].count()
                
                col1, col2 = st.columns(2)
                col1.metric(label="Total Volume Coulé (m³)", value=f"{total_vol_jour:.2f} m³")
                col2.metric(label="Nombre de Toupies / BL", value=total_liv_jour)
                
                # Tableau récapitulatif groupé par Classe de Béton
                st.markdown("#### Détail par Classe de Béton")
                recap_j = df_jour.groupby("classe_beton").agg(
                    Nombre_BL=("bl_num", "count"),
                    Volume_m3=("quantite_m3", "sum"),
                    Total_Eprouvettes=("nb_eprouvettes", "sum")
                ).reset_index()
                
                # Affichage propre du tableau
                st.dataframe(recap_j, use_container_width=True, hide_index=True)
# Bouton Excel
    titre_j = f"Recapitulatif - {d_jour.strftime('%d/%m/%Y')}"
    excel_data = generer_excel_recap(recap_j, titre_j)
    st.download_button(
        label="📥 Télécharger en Excel",
        data=excel_data,
        file_name="Recap_Beton.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
        # ==========================================
        # ONGLET 2 : BILAN MENSUEL
        # ==========================================
with tab_mois:
            st.subheader("Filtrage par mois et année")
            col_m1, col_m2 = st.columns(2)
            
            with col_m1:
                # Liste des mois de 1 à 12, index par défaut = mois actuel
                mois_choisi = st.selectbox("Mois", range(1, 13), index=date.today().month - 1)
            with col_m2:
                # Années possibles
                annee_choisie = st.selectbox("Année", range(2024, 2030), index=2) # 2026 par défaut
                
            # Filtrer le DataFrame pour le mois et l'année choisis
            df_mois = df[
                (df['date_livraison_dt'].dt.month == mois_choisi) & 
                (df['date_livraison_dt'].dt.year == annee_choisie)
            ]
            
            if df_mois.empty:
                st.info(f"Aucun coulage enregistré pour la période {mois_choisi:02d}/{annee_choisie}.")
            else:
                # Calculs des totaux du mois
                total_vol_mois = df_mois["quantite_m3"].sum()
                total_liv_mois = df_mois["bl_num"].count()
                
                col1, col2 = st.columns(2)
                col1.metric(label="Volume Mensuel Cumulé (m³)", value=f"{total_vol_mois:.2f} m³")
                col2.metric(label="Nombre de Toupies / BL", value=total_liv_mois)
                
                # Tableau récapitulatif groupé par Classe de Béton
                st.markdown(f"#### Synthèse des Coulages pour {mois_choisi:02d}/{annee_choisie}")
                recap_m = df_mois.groupby("classe_beton").agg(
                    Nombre_BL=("bl_num", "count"),
                    Volume_m3=("quantite_m3", "sum"),
                    Total_Eprouvettes=("nb_eprouvettes", "sum")
                ).reset_index()
                
                # Affichage propre
                st.dataframe(recap_m, use_container_width=True, hide_index=True)
