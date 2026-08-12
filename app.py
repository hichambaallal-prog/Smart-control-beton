import os
import streamlit as st
from supabase import create_client, Client

# 1. Configuration de la page Streamlit
st.set_page_config(
    page_title="LPEE - CTR-CSB",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Gestion des rôles
if "role" not in st.session_state:
    st.session_state.role = None  # Peut être None, "user", ou "admin"

# --- ÉCRAN DE CONNEXION ---
if st.session_state.role is None:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.title("🔐 Accès Restreint - LPEE")
        st.caption("Veuillez saisir le mot de passe.")
        
        password = st.text_input("Mot de passe", type="password")
        
        if st.button("Se connecter", use_container_width=True):
            if password == "ctr2026": 
                st.session_state.role = "user"
                st.rerun()
            elif password == "admin2026":  # <-- MOT DE PASSE ADMIN
                st.session_state.role = "admin"
                st.rerun()
            else:
                st.error("❌ Mot de passe incorrect.")
    st.stop()

# ==========================================
# 3. CODE PRINCIPAL (Affiché si connecté)
# ==========================================
try:
    from views import suivi_Betonnage, essai_Plaque, synthese_Beton, synthese_plaque
except ImportError as e:
    st.error(f"❌ Erreur lors de l'importation des vues : {e}")
    st.stop()

try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    supabase = None
    st.error(f"❌ Erreur de connexion Supabase : {e}")

# Affichage du rôle dans la sidebar
with st.sidebar:
    st.title("LPEE - CTR-CSB")
    st.info(f"Connecté en tant que : **{st.session_state.role.upper()}**")
    st.markdown("---")
    
    page = st.radio(
        "Menu Principal",
        ["Accueil", "Essai à la Plaque", "Synthèse Plaque", "Suivi de Bétonnage", "Synthèse Béton"]
    )
    
    st.markdown("---")
    if st.button("🚪 Déconnexion", use_container_width=True):
        st.session_state.role = None
        st.rerun()

# Routage des vues
if page == "Accueil":
    st.title("🚄 Accueil - LGV CASA SUD")
    st.markdown("### Plateforme de Suivi et Contrôle Qualité - LPEE")
    
    st.markdown("---")
    
    # Affichage sécurisé de la photo Al Boraq (avec le nom exact du fichier sur GitHub)
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
            st.warning("⚠️ L'image 'al_boraq.jpg.jpg' est introuvable.")
        
    st.markdown("---")
    
    # Section de présentation
    st.markdown("""
    Bienvenue sur l'application centralisée de gestion des contrôles qualité pour le projet **LGV CASA SUD**. 
    
    Utilisez le menu de navigation latéral pour accéder aux différents modules de saisie et de suivi :
    * **🏗️ Suivi Béton :** Gestion des livraisons, fiches de contrôle, températures, affaissements et prélèvements.
    * **🧪 Essai à la Plaque :** Saisie des essais de portance (Norme NF P 94-117-1) avec calculs automatiques des modules EV1, EV2 et du coefficient K.
    """)

elif page == "Essai à la Plaque":
    essai_Plaque.show(supabase)
elif page == "Synthèse Plaque":
    synthese_plaque.show(supabase)
elif page == "Suivi de Bétonnage":
    suivi_Betonnage.show(supabase)
elif page == "Synthèse Béton":
    synthese_Beton.show(supabase)
