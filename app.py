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
# Stockage de la base utilisateurs en session pour permettre la modification dynamique
if "users_db" not in st.session_state:
    st.session_state["users_db"] = {
        # Administrateur
        "BAALLAL": {"password": "arwa2020", "role": "admin", "can_edit": True},
        
        # Techniciens Laboratoire & Responsable de dossier
        "AMINA": {"password": "amina2026", "role": "laboratoire", "can_edit": True},
        "HANINE": {"password": "hanine2026", "role": "laboratoire", "can_edit": False},
        "IKKEN": {"password": "ikken2026", "role": "laboratoire", "can_edit": False},
        "HAMDANI": {"password": "hamdani2026", "role": "laboratoire", "can_edit": False},
        
        # Opérateurs Bétonnage
        "ADAM": {"password": "ctr2026", "role": "restricted_betonnage", "can_edit": False},
        "LAHCEN": {"password": "ctr2026", "role": "restricted_betonnage", "can_edit": False},
        "ELIDRISSI": {"password": "ctr2026", "role": "restricted_betonnage", "can_edit": False}
    }

USERS_DB = st.session_state["users_db"]

if "user" not in st.session_state:
    st.session_state["user"] = None
if "role" not in st.session_state:
    st.session_state["role"] = None
if "can_edit" not in st.session_state:
    st.session_state["can_edit"] = False

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
                if username_input in USERS_DB and USERS_DB[username_input]["password"] == password_input:
                    user_role = USERS_DB[username_input]["role"]
                    can_edit = USERS_DB[username_input]["can_edit"]
                    st.session_state["user"] = {"username": username_input, "role": user_role}
                    st.session_state["role"] = user_role
                    st.session_state["can_edit"] = can_edit
                    st.rerun()
                elif password_input == "admin2026":
                    username = username_input if username_input else "ADMIN"
                    st.session_state["user"] = {"username": username, "role": "admin"}
                    st.session_state["role"] = "admin"
                    st.session_state["can_edit"] = True
                    st.rerun()
                elif password_input == "ctr2026":
                    username = username_input if username_input else "USER"
                    st.session_state["user"] = {"username": username, "role": "user"}
                    st.session_state["role"] = "user"
                    st.session_state["can_edit"] = False
                    st.rerun()
                else:
                    st.error("❌ Nom d'utilisateur ou mot de passe incorrect.")
    st.stop()

# ==========================================
# 3. CODE PRINCIPAL (Utilisateur connecté)
# ==========================================
# Masquer les boutons de téléchargement pour le rôle "user" (Consultation seule)
if st.session_state.get("role") == "user":
    st.markdown(
        """
        <style>
        .stDownloadButton { display: none !important; }
        </style>
        """,
        unsafe_allow_html=True
    )

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

# Connexion à Supabase
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
    
    # Affichage du rôle et définition des menus accessibles
    if current_role == "laboratoire" or current_role == "technicien":
        if current_username == "HANINE":
            st.info("Rôle : **RESPONSABLE DE DOSSIER**")
        elif current_username == "AMINA":
            st.info("Rôle : **TECHNICIENNE LABORATOIRE (Saisie/Modification)**")
        else:
            st.info("Rôle : **TECHNICIEN LABORATOIRE**")
        st.markdown("---")
        available_pages = [
            "Accueil", 
            "Suivi Contrôle Béton", 
            "Suivi de Bétonnage", 
            "Essai à la Plaque", 
            "Synthèse Béton", 
            "Synthèse Plaque"
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
            "Gestion Utilisateurs",
            "Essai à la Plaque", 
            "Synthèse Plaque", 
            "Suivi de Bétonnage", 
            "Suivi Contrôle Béton", 
            "Synthèse Béton"
        ]
    elif current_role == "user":
        st.info("Rôle : **CONSULTATION (LECTURE SEULE)**")
        st.markdown("---")
        # Restriction stricte pour "user" : Synthèses uniquement
        available_pages = [
            "Accueil", 
            "Synthèse Béton", 
            "Synthèse Plaque"
        ]
    else:
        st.info(f"Rôle : **{current_role.upper()}**")
        st.markdown("---")
        available_pages = [
            "Accueil", 
            "Synthèse Béton", 
            "Synthèse Plaque"
        ]
    
    page = st.radio("Menu Principal", available_pages)
    
    st.markdown("---")
    
    # --- MODULE DE MODIFICATION DE MOT DE PASSE ---
    with st.expander("🔑 Changer mon mot de passe"):
        with st.form("change_pwd_form", clear_on_submit=True):
            old_pwd = st.text_input("Ancien mot de passe", type="password")
            new_pwd = st.text_input("Nouveau mot de passe", type="password")
            confirm_pwd = st.text_input("Confirmer le mot de passe", type="password")
            submit_pwd = st.form_submit_button("Mettre à jour", use_container_width=True)
            
            if submit_pwd:
                user_record = st.session_state["users_db"].get(current_username)
                
                if user_record and old_pwd != user_record["password"]:
                    st.error("❌ L'ancien mot de passe est incorrect.")
                elif new_pwd == "":
                    st.warning("⚠️ Le nouveau mot de passe ne peut pas être vide.")
                elif new_pwd != confirm_pwd:
                    st.error("❌ Les nouveaux mots de passe ne correspondent pas.")
                else:
                    st.session_state["users_db"][current_username]["password"] = new_pwd
                    st.success("✅ Mot de passe modifié avec succès !")

    st.markdown("---")
    if st.button("🚪 Déconnexion", use_container_width=True):
        st.session_state["user"] = None
        st.session_state["role"] = None
        st.session_state["can_edit"] = False
        st.rerun()

# Routage des vues
if page == "Accueil":
    st.title("🚄 Accueil - LGV CASA SUD")
    st.markdown("### Plateforme de Suivi et Contrôle Qualité - LPEE")
    
    st.markdown("---")
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
    st.markdown("""
    Bienvenue sur l'application centralisée de gestion des contrôles qualité pour le projet **LGV CASA SUD**.
    
    Utilisez le menu de navigation latéral pour accéder aux différents modules de consultation et de suivi.
    """)

elif page == "Gestion Utilisateurs" and current_role == "admin":
    st.title("👥 Gestion des Utilisateurs & Mots de Passe")
    st.caption("Consultez et gérez la liste de tous les utilisateurs de la plateforme.")

    ROLES_LIST = ["laboratoire", "restricted_betonnage", "admin", "user"]

    col_add, col_edit = st.columns(2)

    with col_add:
        with st.expander("➕ Ajouter un utilisateur", expanded=False):
            with st.form("add_user_form", clear_on_submit=True):
                new_username = st.text_input("Nom d'utilisateur").strip().upper()
                new_password = st.text_input("Mot de passe", type="password")
                new_role = st.selectbox("Rôle", ROLES_LIST)
                new_can_edit = st.checkbox("Droit de modification (can_edit)")
                submit_add = st.form_submit_button("Ajouter l'utilisateur", use_container_width=True, type="primary")

                if submit_add:
                    if not new_username:
                        st.error("❌ Le nom d'utilisateur ne peut pas être vide.")
                    elif not new_password:
                        st.error("❌ Le mot de passe ne peut pas être vide.")
                    elif new_username in st.session_state["users_db"]:
                        st.warning(f"⚠️ L'utilisateur **{new_username}** existe déjà.")
                    else:
                        st.session_state["users_db"][new_username] = {
                            "password": new_password,
                            "role": new_role,
                            "can_edit": new_can_edit
                        }
                        st.success(f"✅ Utilisateur **{new_username}** ajouté avec succès !")
                        st.rerun()

    with col_edit:
        with st.expander("✏️ Modifier un utilisateur", expanded=False):
            user_list = list(st.session_state["users_db"].keys())
            selected_user = st.selectbox("Sélectionner un utilisateur à modifier", user_list)

            if selected_user:
                current_data = st.session_state["users_db"][selected_user]
                with st.form("edit_user_form"):
                    mod_password = st.text_input("Nouveau mot de passe (laisser vide si inchangé)", type="password")
                    
                    role_index = ROLES_LIST.index(current_data["role"]) if current_data["role"] in ROLES_LIST else 0
                    mod_role = st.selectbox("Rôle", ROLES_LIST, index=role_index)
                    mod_can_edit = st.checkbox("Droit de modification (can_edit)", value=current_data["can_edit"])
                    
                    submit_edit = st.form_submit_button("Enregistrer les modifications", use_container_width=True)

                    if submit_edit:
                        updated_password = mod_password if mod_password != "" else current_data["password"]
                        st.session_state["users_db"][selected_user] = {
                            "password": updated_password,
                            "role": mod_role,
                            "can_edit": mod_can_edit
                        }
                        st.success(f"✅ Utilisateur **{selected_user}** mis à jour !")
                        st.rerun()

    st.markdown("---")

    data_users = []
    for user, details in st.session_state["users_db"].items():
        data_users.append({
            "Utilisateur": user,
            "Mot de Passe": details["password"],
            "Rôle": details["role"],
            "Droit de modification (can_edit)": details["can_edit"]
        })
    
    st.dataframe(data_users, use_container_width=True)

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
