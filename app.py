import streamlit as st
import pandas as pd
from supabase import create_client, Client

st.set_page_config(page_title="Suivi Béton - LGV Casa Sud (LPEE)", layout="wide")

# =========================================================
# 1. MOTS DE PASSE (GÉNÉRAL ET ADMINISTRATEUR)
# =========================================================
PASSWORD_GENERAL = "lpee2026"          # Pour les techniciens / consultation & ajout
PASSWORD_ADMIN = "lpee@2026"  # ⚠️ VOTE CODE SECRET POUR MODIFIER/SUPPRIMER

# Connexion générale
if "authentifie" not in st.session_state:
    st.session_state["authentifie"] = False

if "is_admin" not in st.session_state:
    st.session_state["is_admin"] = False

if not st.session_state["authentifie"]:
    st.title("🔒 Accès Sécurisé - LPEE")
    st.warning("Veuillez saisir le mot de passe du laboratoire pour accéder à l'application.")
    
    pwd = st.text_input("Mot de passe", type="password")
    if st.button("Valider l'accès"):
        if pwd == PASSWORD_GENERAL or pwd == PASSWORD_ADMIN:
            st.session_state["authentifie"] = True
            if pwd == PASSWORD_ADMIN:
                st.session_state["is_admin"] = True
            st.rerun()
        else:
            st.error("❌ Mot de passe incorrect.")
    st.stop()

# =========================================================
# 2. CONNEXION SUPABASE
# =========================================================
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1"

try:
    supabase: Client = create_client(URL, CLE)
    response = supabase.table("controles_beton").select("*").execute()
    data = response.data
except Exception as e:
    st.error(f"Erreur de connexion : {e}")
    data = []

# =========================================================
# 3. ONGLETS AVEC VERROUILLAGE ADMIN
# =========================================================
tab_ajouter, tab_modifier, tab_supprimer = st.tabs([
    "➕ Nouveau rapport", 
    "✏️ Modifier un rapport (Admin)", 
    "❌ Supprimer un rapport (Admin)"
])

# --- ONGLET 1 : AJOUTER (Accessible à tous) ---
with tab_ajouter:
    st.subheader("Saisie d'un nouveau rapport")
    # ... (Gardez le formulaire d'ajout habituel) ...

# --- ONGLET 2 : MODIFIER (Réservé à l'Admin) ---
with tab_modifier:
    if not st.session_state["is_admin"]:
        st.warning("🔒 Cet espace est réservé à l'administrateur du laboratoire.")
        admin_key = st.text_input("Saisissez le code Administrateur pour débloquer :", type="password", key="pwd_mod")
        if st.button("Débloquer la modification"):
            if admin_key == PASSWORD_ADMIN:
                st.session_state["is_admin"] = True
                st.success("Accès Administrateur activé !")
                st.rerun()
            else:
                st.error("Code Administrateur incorrect.")
    else:
        st.success("🔓 Mode Administrateur Actif")
        # ... (Placer ici le formulaire de modification) ...

# --- ONGLET 3 : SUPPRIMER (Réservé à l'Admin) ---
with tab_supprimer:
    if not st.session_state["is_admin"]:
        st.warning("🔒 Cet espace est réservé à l'administrateur du laboratoire.")
        admin_key = st.text_input("Saisissez le code Administrateur pour débloquer :", type="password", key="pwd_del")
        if st.button("Débloquer la suppression"):
            if admin_key == PASSWORD_ADMIN:
                st.session_state["is_admin"] = True
                st.success("Accès Administrateur activé !")
                st.rerun()
            else:
                st.error("Code Administrateur incorrect.")
    else:
        st.success("🔓 Mode Administrateur Actif")
        # ... (Placer ici l'option de suppression) ...
