import streamlit as st

# Configurer la page (Optionnel)
st.set_page_config(page_title="LPEE - CTR-CSB", page_icon="🏗️", layout="wide")

# -----------------------------------------------------------------------------
# 1. GESTION DE LA SESSION & MOT DE PASSE
# -----------------------------------------------------------------------------

# Définissez votre mot de passe d'accès unique ici
MOT_DE_PASSE_PROJET = "1234"  # 👈 Changez le mot de passe ici

# Initialisation de la variable de connexion dans la session
if "connecte" not in st.session_state:
    st.session_state["connecte"] = False


# -----------------------------------------------------------------------------
# 2. ÉCRAN DE CONNEXION (Accès avec mot de passe)
# -----------------------------------------------------------------------------

if not st.session_state["connecte"]:
    st.title("🔒 Accès Sécurisé - Mon Projet")
    st.markdown("Veuillez saisir le mot de passe pour accéder aux modules du projet.")

    # Formulaire de saisie du mot de passe
    mot_de_passe = st.text_input("Mot de passe :", type="password")

    if st.button("Se connecter", type="primary"):
        if mot_de_passe == MOT_DE_PASSE_PROJET:
            st.session_state["connecte"] = True
            st.success("Accès autorisé !")
            st.rerun()  # Recharge l'application vers l'espace connecté
        else:
            st.error("Mot de passe incorrect. Veuillez réentreprendre.")


# -----------------------------------------------------------------------------
# 3. ESPACE CONNECTÉ (Navigation sans re-demander le mot de passe)
# -----------------------------------------------------------------------------

else:
    # Déclaration des pages de vos modules
    page_suivi = st.Page("views/suivi_betonnage.py", title="Suivi de Bétonnage", icon="🏗️")
    page_controle = st.Page("views/controle_beton.py", title="Contrôle de Béton", icon="🧪")
    page_plaque = st.Page("views/essai_plaque.py", title="Essai à la Plaque", icon="🪨")

    # En-tête de la barre latérale
    st.sidebar.title("🏢 LPEE - CTR-CSB")
    st.sidebar.caption("Projet : LGV CASA SUD | Client : TGCC")
    st.sidebar.divider()

    # Navigation unifiée : l'utilisateur passe d'un module à un autre sans mot de passe
    pg = st.navigation({
        "Mon Projet": [page_suivi, page_controle, page_plaque]
    })

    # Bouton de déconnexion dans la barre latérale
    st.sidebar.divider()
    if st.sidebar.button("🚪 Déconnexion"):
        st.session_state["connecte"] = False
        st.rerun()

    # Exécution de la page sélectionnée
    pg.run()
