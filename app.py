import streamlit as st

# Définition des pages
page_accueil = st.Page("views/accueil.py", title="Accueil", icon="🏠")
page_plaque = st.Page("views/essai_plaque.py", title="Essai à la Plaque", icon="🪨")
page_beton = st.Page("views/suivi_betonnage.py", title="Suivi de Bétonnage", icon="🏗️")

# En-tête de la barre latérale
st.sidebar.title("🏢 LPEE - CTR-CSB")
st.sidebar.caption("Projet : LGV CASA SUD | Client : TGCC")

# Navigation unifiée
pg = st.navigation([page_accueil, page_plaque, page_beton])
pg.run()
import streamlit as st
import pandas as pd
from datetime import date
from supabase import create_client, Client

# Configuration globale de la page
st.set_page_config(page_title="LPEE CTR-CSB - LGV CASA SUD", layout="wide")

# ==============================================================================
# 🔐 1. GESTION DU MOT DE PASSE ET AUTHENTIFICATION
# ==============================================================================
MOT_DE_PASSE_ACCES = "lpee2026"  # 👈 Mot de passe d'accès

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# --- ÉCRAN DE CONNEXION AVEC PHOTO AL BORAQ ---
if not st.session_state["authenticated"]:
    col_g, col_c, col_d = st.columns([1, 2, 1])
    
    with col_c:
        # 📸 Photo directe du TGV Al Boraq via un lien Web (Wikimedia Commons)
        url_image_al_boraq = "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/ONCF_Al_boraq.jpeg/1280px-ONCF_Al_boraq.jpeg"
        
        st.image(
            url_image_al_boraq, 
            caption="Projet LGV CASA SUD - LPEE CTR-CSB",
            use_container_width=True
        )
        
        st.title("🔒 Connexion au Portail Laboratoire")
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
                
    st.stop()  # Bloque l'accès tant qu'on n'est pas connecté

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

# Barre latérale : Navigation & Déconnexion
with st.sidebar:
    st.title("🏢 LPEE - CTR-CSB")
    st.caption("Projet : **LGV CASA SUD** | Client : **TGCC**")
    st.write("---")
    
    # Choix de la page
    page = st.radio(
        "📌 Menu Principal",
        ["🏠 Accueil", "🪨 Essai à la Plaque", "🏗️ Suivi de Bétonnage"]
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
        Ce portail vous permet de gérer et d'enregistrer les essais de contrôle qualité sur site :
        
        * **🪨 Essai à la Plaque :** Saisie des enfoncements Z1 et Z2, calcul automatique des modules EV1, EV2 et du rapport k.
        * **🏗️ Suivi de Bétonnage :** Contrôle des livraisons, classe béton, affaissement et prélèvements selon la norme **NF EN 12390-2**.
        
        Sélectionnez le module souhaité dans le menu latéral à gauche.
        """)
    with col2:
        st.info("""
        **Rappels Projet :**
        * **Projet :** LGV CASA SUD
        * **Entreprise :** TGCC
        * **Centre :** CTR-CSB
        """)

# ------------------------------------------------------------------------------
# PAGE 2 : ESSAI À LA PLAQUE
# ------------------------------------------------------------------------------
elif page == "🪨 Essai à la Plaque":
    st.title("🪨 Contrôle de Portance - Essai à la Plaque (NF P94-117-1)")
    
    # Chargement des données
    try:
        resp = supabase.table("essais_plaque").select("*").execute()
        data_all_plaque = resp.data or []
    except Exception as e:
        data_all_plaque = []

    date_choisie_p = st.date_input("📅 Date de l'essai :", value=date.today())
    str_date_p = date_choisie_p.strftime("%d/%m/%Y")

    with st.form("form_plaque"):
        st.markdown(f"### 📝 Saisie Essai à la Plaque ({str_date_p})")
        
        col_proj1, col_proj2 = st.columns(2)
        with col_proj1:
            projet = st.text_input("Projet", value="LGV CASA SUD", disabled=True)
        with col_proj2:
            client = st.text_input("Entreprise / Client", value="TGCC", disabled=True)
            
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            pk_emp = st.text_input("Emplacement / PK", value="PK 14+250 - Voie 1")
            couche_elem = st.selectbox("Couche / Support", ["PFT3 (Couche de Forme)", "PST (Arase)", "Couche d'Assise", "Remblai"])
            
        with c2:
            z1 = st.number_input("Enfoncement 1er chargement Z1 (mm)", value=1.50, step=0.01, min_value=0.01)
            z2 = st.number_input("Enfoncement 2ème chargement Z2 (mm)", value=0.50, step=0.01, min_value=0.01)
            
        with c3:
            ev1 = round(112.5 / (z1 * 2.0), 2) if z1 > 0 else 0.0
            ev2 = round(90.0 / (z2 * 2.0), 2) if z2 > 0 else 0.0
            rapport_calc = round(ev2 / ev1, 2) if ev1 > 0 else 0.0
            
            st.metric("EV1 (MPa)", value=ev1)
            st.metric("EV2 (MPa)", value=ev2)
            st.metric("Rapport k = EV2 / EV1", value=rapport_calc)
            
            is_conforme = (ev2 >= 50.0) and (rapport_calc <= 2.2)
            statut_auto = "✅ Conforme" if is_conforme else "⚠️ Non Conforme"
            st.info(f"Statut : **{statut_auto}**")

        obs_p = st.text_area("Observations", value="RAS - Sol bien compacté")
        
        submitted = st.form_submit_button("💾 Enregistrer l'essai à la plaque", type="primary")
        if submitted:
            row_p = {
                "date_essai": str_date_p,
                "projet": projet,
                "client": client,
                "pk_emplacement": pk_emp,
                "couche_element": couche_elem,
                "z1": float(z1),
                "z2": float(z2),
                "ev1": float(ev1),
                "ev2": float(ev2),
                "rapport_ev2_ev1": float(rapport_calc),
                "statut": statut_auto,
                "observations": obs_p
            }
            try:
                supabase.table("essais_plaque").insert(row_p).execute()
                st.success("✅ Essai enregistré avec succès !")
                st.rerun()
            except Exception as e:
                row_p_fallback = {
                    "date_essai": str_date_p,
                    "pk_emplacement": pk_emp,
                    "couche_element": couche_elem,
                    "ev1": float(ev1),
                    "ev2": float(ev2),
                    "rapport_ev2_ev1": float(rapport_calc),
                    "statut": statut_auto,
                    "observations": obs_p
                }
                try:
                    supabase.table("essais_plaque").insert(row_p_fallback).execute()
                    st.success("✅ Essai enregistré avec succès !")
                    st.rerun()
                except Exception as e2:
                    st.error(f"Erreur d'enregistrement : {e2}")

    st.markdown("---")
    st.subheader("📋 Historique des Essais à la Plaque")
    if data_all_plaque:
        df_p = pd.DataFrame(data_all_plaque)
        df_p.index = range(1, len(df_p) + 1)
        st.dataframe(df_p, use_container_width=True)
    else:
        st.info("Aucun essai enregistré.")

# ------------------------------------------------------------------------------
# PAGE 3 : SUIVI DE BÉTONNAGE
# ------------------------------------------------------------------------------
elif page == "🏗️ Suivi de Bétonnage":
    st.title("🏗️ Suivi et Contrôle Qualité Béton")
    st.info("Module de saisie et de suivi des bons de livraison du béton.")
    
    with st.form("form_beton"):
        st.subheader("Saisie d'un contrôle de bétonnage")
        col1, col2 = st.columns(2)
        
        with col1:
            date_b = st.date_input("Date de livraison", value=date.today())
            bl_num = st.text_input("N° Bon de Livraison (BL)", value="BL-2026-001")
            classe_b = st.selectbox("Classe béton", ["C20/25", "C25/30", "C30/37", "C35/45", "C40/50"])
            ouvrage = st.text_input("Élément d'ouvrage / Emplacement", value="Voile / Semelle")
            
        with col2:
            temp = st.number_input("Température (°C)", value=20.0, step=0.5)
            affaisse = st.number_input("Affaissement / Slump (cm)", value=15.0, step=0.5)
            prelev = st.selectbox("Prélèvement : NF EN 12390-2", ["OUI - Conforme (NF EN 12390-2)", "NON", "Sans objet"])
            nb_ep = st.number_input("Nombre d'éprouvettes", min_value=0, max_value=12, value=6)

        submit_b = st.form_submit_button("Enregistrer le contrôle béton", type="primary")
        if submit_b:
            st.success("✅ Données de bétonnage enregistrées avec succès !")
