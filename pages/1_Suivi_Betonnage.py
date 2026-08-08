import streamlit as st
import pandas as pd
from datetime import date
from supabase import create_client, Client

st.set_page_config(page_title="Suivi de Bétonnage - LPEE", layout="wide")

# Connection Supabase
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
# ⚠️ Remplacez ci-dessous par votre vraie clé ANON qui commence par eyJ...
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0" 

try:
    supabase: Client = create_client(URL, CLE)
except Exception as e:
    st.error(f"Erreur Supabase : {e}")
    st.stop()

st.title("🏗️ Suivi du Bétonnage")
st.markdown("### Laboratoire Public d'Essais et d'Études (LPEE)")

# Chargement global des données
try:
    resp = supabase.table("controles_beton").select("*").execute()
    data_all = resp.data or []
except Exception as e:
    st.error(f"Erreur de chargement : {e}")
    data_all = []

# Sélecteur de date
date_choisie = st.date_input("📅 Date du bétonnage :", value=date.today())
str_date = date_choisie.strftime("%d/%m/%Y")

# --- 1. FORMULAIRE DE SAISIE ---
with st.form("form_beton"):
    st.subheader(f"📝 Saisie d'une coulée ({str_date})")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        pk_emp = st.text_input("Emplacement / PK", value="PK 12+500")
        element = st.text_input("Élément / Support", value="Semelle S1")
        
    with col2:
        num_bon = st.text_input("N° Bon de Livraison", value="BL-1024")
        volume = st.number_input("Volume (m³)", value=8.0, step=0.5)
        
    with col3:
        temp = st.number_input("Température (°C)", value=22.0, step=0.5)
        # 🟢 CASE TECHNICIEN AJOUTÉE ICI :
        technicien = st.text_input("Technicien Contrôleur", value="Ismail / Mohamed")

    obs = st.text_area("Observations", value="Béton conforme au B.L.")
    
    submitted = st.form_submit_button("💾 Enregistrer la coulée", type="primary")
    
    if submitted:
        nouvelle_coulee = {
            "date_essai": str_date,
            "pk_emplacement": pk_emp,
            "couche_element": element,
            "num_bon": num_bon,
            "volume": float(volume),
            "temperature": float(temp),
            "technicien": technicien,  # Enregistré dans Supabase
            "observations": obs
        }
        
        try:
            supabase.table("controles_beton").insert(nouvelle_coulee).execute()
            st.success("✅ Coulée enregistrée avec succès !")
            st.rerun()
        except Exception as e:
            st.error(f"Erreur lors de l'enregistrement : {e}")

st.markdown("---")

# --- 2. RÉCAPITULATIF ET QUANTITÉ TOTALE ---
st.subheader("📊 Récapitulatif et Quantité Totale")

if data_all:
    df = pd.DataFrame(data_all)
    
    # Filtrer les coulées de la date sélectionnée
    df_jour = df[df["date_essai"] == str_date] if ("date_essai" in df.columns and not df.empty) else pd.DataFrame()
    
    # Calculs des métriques
    total_jour = df_jour["volume"].sum() if ("volume" in df_jour.columns and not df_jour.empty) else 0.0
    total_global = df["volume"].sum() if "volume" in df.columns else 0.0
    nb_coulees_jour = len(df_jour)
    
    # Affichage des cartes récapitulatives
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("🏗️ Béton Total du Jour", f"{total_jour:.2f} m³")
    with m2:
        st.metric("🔢 Nombre de Coulées (Jour)", f"{nb_coulees_jour}")
    with m3:
        st.metric("📈 Volume Cumulé Global", f"{total_global:.2f} m³")

    st.markdown("### 📋 Historique Général des Coulées")
    df.index = range(1, len(df) + 1)
    st.dataframe(df, use_container_width=True)
else:
    st.info("Aucune coulée enregistrée pour le moment.")
