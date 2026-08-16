import os
import bcrypt
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
# 2. CONNEXION A LA BASE DE DONNEES SUPABASE
# ==========================================
try:
    SUPABASE_URL = st.secrets.get("SUPABASE_URL", "https://votre-projet.supabase.co")
    SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1")
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    supabase = None
    st.error(f"❌ Erreur de connexion Supabase : {e}")

# ==========================================
# 3. GESTION DES SESSIONS & AUTHENTIFICATION
# ==========================================
if "user" not in st.session_state:
    st.session_state["user"] = None
if "role" not in st.session_state:
    st.session_state["role"] = None
if "is_admin" not in st.session_state:
    st.session_state["is_admin"] = False


def verify_password(password: str, hashed_password: str) -> bool:
    """Vérifie le mot de passe saisi par rapport au hash BCrypt stocké."""
    try:
        return bcrypt.checkpw(password.encode('utf-8'), hashed_password.encode('utf-8'))
    except Exception:
        return False


def get_user_from_db(username: str):
    """Recherche un utilisateur dans la table 'users' de Supabase."""
    if not supabase:
        return None
    try:
        res = supabase.table("users").select("*").eq("username", username).execute()
        if res.data:
            return res.data[0]
    except Exception as e:
        st.error(f"Erreur lors de la vérification de l'utilisateur : {e}")
    return None


def ecran_connexion():
    """Affiche le formulaire de connexion par Username + Mot de Passe."""
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔐 Accès Restreint - LPEE")
        st.caption("Veuillez saisir vos identifiants pour accéder à la plateforme.")

        with st.form("login_form", clear_on_submit=False):
            username_input = st.text_input("Nom d'utilisateur").strip()
            password_input = st.text_input("Mot de passe", type="password")
            submit_btn = st.form_submit_button("Se connecter", use_container_width=True, type="primary")

            if submit_btn:
                if not username_input or not password_input:
                    st.warning("⚠️ Veuillez remplir tous les champs.")
                else:
                    user = get_user_from_db(username_input)
                    if user and user.get("is_active", True):
                        if verify_password(password_input, user.get("password_hash", "")):
                            role_str = str(user.get("role", "user")).lower()
                            is_admin = (role_str == "admin")

                            st.session_state["user"] = {
                                "id": user.get("id"),
                                "username": user.get("username"),
                                "role": role_str
                            }
                            st.session_state["role"] = role_str
                            st.session_state["is_admin"] = is_admin

                            st.success(f"Bienvenue {user.get('username')} !")
                            st.rerun()
                        else:
                            st.error("❌ Nom d'utilisateur ou mot de passe incorrect.")
                    else:
                        st.error("❌ Identifiants invalides ou compte désactivé.")


# Affichage du formulaire si l'utilisateur n'est pas connecté
if st.session_state["user"] is None:
    ecran_connexion()
    st.stop()

# ==========================================
# 4. CODE PRINCIPAL (Utilisateur connecté)
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

# Menu latéral (Sidebar)
with st.sidebar:
    st.title("LPEE - CTR-CSB")
    current_username = st.session_state["user"]["username"]
    current_role = st.session_state["role"].upper()

    st.markdown(f"👤 **{current_username}**")
    st.info(f"Rôle : **{current_role}**")
    st.markdown("---")

    page = st.radio(
        "Menu Principal",
        [
            "Accueil",
            "Essai à la Plaque",
            "Synthèse Plaque",
            "Suivi de Bétonnage",
            "Suivi Contrôle Béton",
            "Synthèse Béton"
        ]
    )

    st.markdown("---")
    if st.button("🚪 Déconnexion", use_container_width=True):
        st.session_state["user"] = None
        st.session_state["role"] = None
        st.session_state["is_admin"] = False
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
