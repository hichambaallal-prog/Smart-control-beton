from pathlib import Path
import streamlit as st

# 1. Configuration de la page
st.set_page_config(
    page_title="LPEE - CTR-CSB", 
    page_icon="🏗️", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Chemin absolu vers le dossier de votre projet
BASE_DIR = Path(__file__).parent

# 2. Mot de passe d'accès
MOT_DE_PASSE_PROJET = "1234"

# Initialisation de la session
if "connecte" not in st.session_state:
    st.session_state["connecte"] = False


# -----------------------------------------------------------------------------
# 3. ÉCRAN DE CONNEXION
# -----------------------------------------------------------------------------
if not st.session_state["connecte"]:
    st.title("🔒 Accès Sécurisé - Mon Projet")
    st.markdown("Veuillez saisir le mot de passe pour accéder aux modules du projet.")

    mot_de_passe = st.text_input("Mot de passe :", type="password")

    if st.button("Se connecter", type="primary"):
        if mot_de_passe == MOT_DE_PASSE_PROJET:
            st.session_state["connecte"] = True
            st.success("Accès autorisé !")
            st.rerun()
        else:
            st.error("Mot de passe incorrect.")


# -----------------------------------------------------------------------------
# 4. ESPACE CONNECTÉ (Navigation unifiée)
# -----------------------------------------------------------------------------
else:
    # ⚠️ Vérifiez les noms ci-dessous avec les résultats du diagnostic :
    page_suivi = st.Page(BASE_DIR / "pages" / "1_Suivi_Betonnage.py", title="Suivi de Bétonnage", icon="🏗️")
    page_plaque = st.Page(BASE_DIR / "pages" / "2_Essai_Plaque.py", title="Essai à la Plaque", icon="🪨")

    # En-tête de la barre latérale
    st.sidebar.title("🏢 LPEE - CTR-CSB")
    st.sidebar.caption("Projet : LGV CASA SUD | Client : TGCC")
    st.sidebar.divider()

    # Navigation unifiée
    pg = st.navigation({
        "Mon Projet": [page_suivi, page_plaque]
    })

    # Bouton de déconnexion
    st.sidebar.divider()
    if st.sidebar.button("🚪 Déconnexion"):
        st.session_state["connecte"] = False
        st.rerun()

    # Exécution
    pg.run()
