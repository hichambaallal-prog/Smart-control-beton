import streamlit as st
import pandas as pd

# --- SECTION : ESSAI À LA PLAQUE ---
st.divider() # Ligne de séparation esthétique
st.title("🚧 Essai à la Plaque (Contrôle de portance)")

with st.form("form_essai_plaque"):
    st.subheader("Saisie des données de chargement")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Saisie de Z1 avec une valeur minimum pour éviter la division par zéro
        z1 = st.number_input("Enfoncement 1er chargement Z1 (mm)", min_value=0.01, value=1.50, step=0.05)
        
    with col2:
        # Saisie de Z2 avec une valeur minimum
        z2 = st.number_input("Enfoncement 2ème chargement Z2 (mm)", min_value=0.01, value=1.20, step=0.05)

    submit_plaque = st.form_submit_button("Calculer les modules")

# --- CALCUL ET AFFICHAGE DES RÉSULTATS ---
if submit_plaque:
    # Application des formules mathématiques
    ev1 = 112.5 / (z1 * 2)
    ev2 = 90 / (z2 * 2)
    k = ev2 / ev1
    
    st.success("Calculs effectués avec succès !")
    
    # Affichage des résultats sous forme de métriques visuelles
    col_res1, col_res2, col_res3 = st.columns(3)
    
    with col_res1:
        st.metric(label="Module EV1 (MPa)", value=f"{ev1:.2f}")
        
    with col_res2:
        st.metric(label="Module EV2 (MPa)", value=f"{ev2:.2f}")
        
    with col_res3:
        # Affiche le rapport K en mettant en évidence les éventuelles non-conformités (exemple de seuil indicatif K <= 2.0)
        st.metric(label="Rapport K (EV2/EV1)", value=f"{k:.2f}")
        
    # Interprétation visuelle rapide du résultat
    if k <= 2.0:
        st.info("✅ Le rapport de compactage K est satisfaisant (≤ 2).")
    else:
        st.warning("⚠️ Le rapport de compactage K est supérieur à 2. Vérifier la qualité du compactage.")
