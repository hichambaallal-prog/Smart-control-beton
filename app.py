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
