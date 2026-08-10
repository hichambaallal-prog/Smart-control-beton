import streamlit as st

# 🔒 Sécurité : Bloque l'accès si non connecté
if not st.session_state.get("authenticated", False):
    st.error("⛔ Accès refusé. Veuillez d'abord vous connecter sur la page d'accueil.")
    st.stop()
import streamlit as st
import pandas as pd
from datetime import date
from supabase import create_client, Client

st.set_page_config(page_title="Essai à la Plaque - LPEE CTR-CSB", layout="wide")

# Connexion Supabase
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
# ⚠️ Remplacez la chaîne ci-dessous par votre clé ANON Supabase
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0"

try:
    supabase: Client = create_client(URL, CLE)
except Exception as e:
    st.error(f"Erreur Supabase : {e}")
    st.stop()

# --- EN-TÊTE DU RAPPORT ---
st.title("🪨 Contrôle de Portance - Essai à la Plaque (NF P94-117-1)")
st.subheader("🏢 LPEE - CTR-CSB")

# Chargement des données
try:
    resp = supabase.table("essais_plaque").select("*").execute()
    data_all_plaque = resp.data or []
except Exception as e:
    st.error(f"Erreur de chargement : {e}")
    data_all_plaque = []

date_choisie_p = st.date_input("📅 Date de l'essai :", value=date.today())
str_date_p = date_choisie_p.strftime("%d/%m/%Y")

# --- FORMULAIRE DE SAISIE ---
with st.form("form_plaque"):
    st.markdown(f"### 📝 Saisie Essai à la Plaque ({str_date_p})")
    
    # Champs d'en-tête du projet VERROUILLÉS (non modifiables)
    col_proj1, col_proj2 = st.columns(2)
    with col_proj1:
        projet = st.text_input("Projet", value="LGV CASA SUD", disabled=True)
    with col_proj2:
        client = st.text_input("Entreprise / Client", value="TGCC", disabled=True)
        
    st.markdown("---")
    
    # Informations techniques et mesures
    c1, c2, c3 = st.columns(3)
    
    with c1:
        pk_emp = st.text_input("Emplacement / PK", value="PK 14+250 - Voie 1")
        couche_elem = st.selectbox("Couche / Support", ["PFT3 (Couche de Forme)", "PST (Arase)", "Couche d'Assise", "Remblai"])
        
    with c2:
        # Données de base chargement Z1 et Z2
        z1 = st.number_input("Enfoncement 1er chargement Z1 (mm)", value=1.50, step=0.01, min_value=0.01)
        z2 = st.number_input("Enfoncement 2ème chargement Z2 (mm)", value=0.50, step=0.01, min_value=0.01)
        
    with c3:
        # Calculs automatiques des modules
        ev1 = round(112.5 / (z1 * 2.0), 2) if z1 > 0 else 0.0
        ev2 = round(90.0 / (z2 * 2.0), 2) if z2 > 0 else 0.0
        rapport_calc = round(ev2 / ev1, 2) if ev1 > 0 else 0.0
        
        # Affichage des métriques
        st.metric("EV1 (MPa)", value=ev1)
        st.metric("EV2 (MPa)", value=ev2)
        st.metric("Rapport k = EV2 / EV1", value=rapport_calc)
        
        is_conforme = (ev2 >= 50.0) and (rapport_calc <= 2.2)
        statut_auto = "✅ Conforme" if is_conforme else "⚠️ Non Conforme"
        st.info(f"Statut : **{statut_auto}**")

    obs_p = st.text_area("Observations", value="RAS - Sol bien compacté")
    
    # Bouton d'enregistrement
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
            # Sécurité de secours si les colonnes 'projet', 'client', 'z1' ou 'z2' ne sont pas encore créées dans Supabase
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
                st.error(f"Erreur lors de l'enregistrement : {e2}")

st.markdown("---")
st.subheader("📋 Historique des Essais à la Plaque")

if data_all_plaque:
    df_p = pd.DataFrame(data_all_plaque)
    df_p.index = range(1, len(df_p) + 1)
    st.dataframe(df_p, use_container_width=True)
else:
    st.info("Aucun essai à la plaque enregistré pour le moment.")
