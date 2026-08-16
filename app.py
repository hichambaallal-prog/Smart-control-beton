import os
import streamlit as st
from supabase import create_client, Client

# ==========================================
# 1. CONFIGURATION DE LA PAGE STREAMLIT
# ==========================================
st.set_page_config(
    page_title="LPEE - CTR-CSB",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. GESTION DES SESSIONS & AUTHENTIFICATION
# ==========================================
if "user" not in st.session_state:
    st.session_state["user"] = None
if "role" not in st.session_state:
    st.session_state["role"] = None

# Dictionnaire des utilisateurs, mots de passe individuels et rôles
USERS_DB = {
    # Administrateur (Accès à TOUTES les vues)
    "BAALLAL": {"password": "arwa2020", "role": "admin"},
    
    # Techniciens Laboratoire (Saisie : Plaque, Suivi Bétonnage, Contrôle Béton)
    "AMINA": {"password": "amina2026", "role": "technicien"},
    "IKKEN": {"password": "ikken2026", "role": "technicien"},
    "ELHAMDANI": {"password": "elhamdani2026", "role": "technicien"},
    
    # Opérateurs Bétonnage (Accès restreint à Suivi Bétonnage)
    "ADAM": {"password": "ctr2026", "role": "restricted_betonnage"},
    "LAHCEN": {"password": "ctr2026", "role": "restricted_betonnage"},
    "ELIDRISSI": {"password": "ctr2026", "role": "restricted_betonnage"}
}

# --- ÉCRAN DE CONNEXION ---
if st.session_state["user"] is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔐 Accès Restreint - LPEE")
        st.caption("Veuillez saisir vos identifiants pour accéder à la plateforme.")
        
        with st.form("login_form", clear_on_submit=False):
            username_input = st.text_input("Nom d'utilisateur").strip().upper()
            password_input = st.text_input("Mot de passe", type="password")
            submit_btn = st.form_submit_button("Se connecter", use_container_width=True, type="primary")
            
            if submit_btn:
                # Vérification via le dictionnaire d'utilisateurs
                if username_input in USERS_DB and USERS_DB[username_input]["password"] == password_input:
                    user_role = USERS_DB[username_input]["role"]
                    st.session_state["user"] = {"username": username_input, "role": user_role}
                    st.session_state["role"] = user_role
                    st.rerun()
                # Codes de secours
                elif password_input == "admin2026":
                    username = username_input if username_input else "ADMIN"
                    st.session_state["user"] = {"username": username, "role": "admin"}
                    st.session_state["role"] = "admin"
                    st.rerun()
                elif password_input == "ctr2026":
                    username = username_input if username_input else "USER"
                    st.session_state["user"] = {"username": username, "role": "user"}
                    st.session_state["role"] = "user"
                    st.rerun()
                else:
                    st.error("❌ Nom d'utilisateur ou mot de passe incorrect.")
    st.stop()

# ==========================================
# 3. CODE PRINCIPAL (Utilisateur connecté)
# ==========================================
try:
    from views import (
        suivi_Betonnage,
        suivi_controle_beton,
        essai_Plaque,
        synthese_Beton,
        synthese_plaque
    )
except ImportError as e:
    st.error(f"❌ Erreur lors de l'importation des vues : {e}")
    st.stop()

# Connexion à la base de données Supabase
try:
    SUPABASE_URL = st.secrets.get("SUPABASE_URL", "https://votre-projet.supabase.co")
    SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1")
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    supabase = None
    st.error(f"❌ Erreur de connexion Supabase : {e}")

# Menu latéral (Sidebar)
with st.sidebar:
    st.title("LPEE - CTR-CSB")
    current_username = st.session_state["user"]["username"]
    current_role = st.session_state["role"]

    st.markdown(f"👤 **{current_username}**")
    
    # Définition des accès et affichage des rôles
    if current_role == "technicien":
        st.info("Rôle : **TECHNICIEN LABORATOIRE**")
        st.markdown("---")
        available_pages = [
            "Accueil", 
            "Essai à la Plaque", 
            "Suivi de Bétonnage", 
            "Suivi Contrôle Béton"
        ]
    elif current_role == "restricted_betonnage":
        st.info("Rôle : **OPÉRATEUR BÉTONNAGE**")
        st.markdown("---")
        available_pages = ["Suivi de Bétonnage"]
    elif current_role == "admin":
        st.info("Rôle : **ADMINISTRATEUR**")
        st.markdown("---")
        available_pages = [
            "Accueil", 
            "Essai à la Plaque", 
            "Synthèse Plaque", 
            "Suivi de Bétonnage", 
            "Suivi Contrôle Béton", 
            "Synthèse Béton"
        ]
    else:
        st.info(f"Rôle : **{current_role.upper()}**")
        st.markdown("---")
        available_pages = [
            "Accueil", 
            "Essai à la Plaque", 
            "Synthèse Plaque", 
            "Suivi de Bétonnage", 
            "Suivi Contrôle Béton", 
            "Synthèse Béton"
        ]
    
    page = st.radio("Menu Principal", available_pages)
    
    st.markdown("---")
    if st.button("🚪 Déconnexion", use_container_width=True):
        st.session_state["user"] = None
        st.session_state["role"] = None
        st.rerun()

# Routage des vues
if page == "Accueil":
    st.title("🚄 Accueil - LGV CASA SUD")
    st.markdown("### Plateforme de Suivi et Contrôle Qualité - LPEE")
    
    st.markdown("---")
    
    # Affichage sécurisé de la photo Al Boraq
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        image_path = os.path.join(os.path.dirname(__file__), "al_boraq.jpg.jpg")
        if os.path.exists(image_path):
            st.image(
                image_path, 
                caption="Al Boraq - Ligne à Grande Vitesse - Projet LGV CASA SUD", 
                use_container_width=True
            )
        else:
            st.warning("⚠️ L'image 'al_boraq.jpg.jpg' est introuvable à la racine.")
        
    st.markdown("---")
    
    # Section de présentation
    st.markdown("""
    Bienvenue sur l'application centralisée de gestion des contrôles qualité pour le projet **LGV CASA SUD**. 
    
    Utilisez le menu de navigation latéral pour accéder aux différents modules de saisie et de suivi :
    * **🏗️ Suivi Béton :** Gestion des livraisons, fiches de contrôle, températures, affaissements et prélèvements.
    * **🧪 Suivi Contrôle Béton :** Saisie des écrasements d'éprouvettes de béton (3j, 7j, 28j, 90j) associées aux prélèvements.
    * **🚜 Essai à la Plaque :** Saisie des essais de portance (Norme NF P 94-117-1) avec calculs automatiques des modules $EV_1$, $EV_2$ et du coefficient $K$.
    """)

elif page == "Essai à la Plaque":
    essai_Plaque.show(supabase)
elif page == "Synthèse Plaque":
    synthese_plaque.show(supabase)
elif page == "Suivi de Bétonnage":
    suivi_Betonnage.show(supabase)
elif page == "Suivi Contrôle Béton":
    suivi_controle_beton.show(supabase)
elif page == "Synthèse Béton":
    synthese_Beton.show(supabase)
