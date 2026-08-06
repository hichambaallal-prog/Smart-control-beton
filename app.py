import streamlit as st
import pandas as pd
from supabase import create_client, Client

# Configuration de la page
st.set_page_config(page_title="Smart Control Béton - LPEE", layout="wide")
st.title("🚧 Supervision Smart Control Béton - LGV Casa Sud")

# Connexion à Supabase
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1"  # Gardez bien votre clé ici

try:
    supabase: Client = create_client(URL, CLE)
except Exception as e:
    st.error(f"Erreur de connexion : {e}")

# Récupération des données de la base
try:
    response = supabase.table("controles_beton").select("*").execute()
    data = response.data
except Exception as e:
    data = []

# --- SECTION 0 : INDICATEURS CLÉS (KPIs MÉTIER) ---
if data and len(data) > 0:
    df = pd.DataFrame(data)
    
    total_coulages = len(df)
    conformes = len(df[df["statut"].str.contains("Conforme", case=False, na=False)]) if "statut" in df.columns else 0
    taux_conf = (conformes / total_coulages) * 100 if total_coulages > 0 else 0
    temp_moy = df["temp_beton"].mean() if "temp_beton" in df.columns else 0

    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
    col_kpi1.metric("📊 Total Coulages Contrôlés", total_coulages)
    col_kpi2.metric("✅ Taux de Conformité", f"{taux_conf:.1f}%")
    col_kpi3.metric("🌡️ Température Moyenne", f"{temp_moy:.1f} °C" if pd.notnull(temp_moy) else "N/A")
    st.write("---")

# --- SECTION 1 : FORMULAIRE DE SAISIE TERRAIN ---
with st.expander("➕ Saisir un nouveau contrôle béton (Terrain)", expanded=False):
    with st.form("form_controle"):
        col1, col2 = st.columns(2)
        with col1:
            ouvrage = st.text_input("Ouvrage / Élément (ex: Barrette B-05)")
            heure_malaxage = st.text_input("Heure de Malaxage (ex: 08:30)")
            heure_arrivee = st.text_input("Heure d'Arrivée (ex: 09:15)")
        with col2:
            temp_beton = st.number_input("Température du Béton (°C)", value=25.0)
            temps_trajet = st.number_input("Temps de Trajet (minutes)", value=30.0)
            statut = st.selectbox("Statut du Contrôle", ["✅ Conforme", "⚠️ Non Conforme"])
        
        submit_button = st.form_submit_button("Envoyer le rapport")
        
        if submit_button:
            if ouvrage:
                data_to_insert = {
                    "ouvrage": ouvrage,
                    "heure_malaxage": heure_malaxage,
                    "heure_arrivee": heure_arrivee,
                    "temp_beton": temp_beton,
                    "temps_trajet": temps_trajet,
                    "statut": statut
                }
                supabase.table("controles_beton").insert(data_to_insert).execute()
                st.success("Rapport enregistré avec succès !")
                st.rerun()
            else:
                st.warning("Veuillez renseigner le nom de l'ouvrage.")

# --- SECTION 2 : REGISTRE NUMÉRIQUE & EXPLOITATION ---
st.subheader("📋 Registre Numérique de Réception du Béton")

if data and len(data) > 0:
    colonnes = ["ouvrage", "heure_malaxage", "heure_arrivee", "temp_beton", "temps_trajet", "statut"]
    df_affichage = df[[c for c in colonnes if c in df.columns]]
    df_affichage.columns = ["Ouvrage / Élément", "Heure Malaxage", "Heure Arrivée", "Température (°C)", "Temps Trajet (min)", "Statut"][:len(df_affichage.columns)]
    
    # Affichage du tableau principal
    st.dataframe(df_affichage, use_container_width=True)

    # Bouton d'export pour exploitation externe (Excel / Rapport LPEE)
    csv = df_affichage.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Télécharger le registre officiel (Format CSV / Excel)",
        data=csv,
        file_name="registre_controle_beton_lgv_casasud.csv",
        mime="text/csv",
    )
else:
    st.info("La base de données est vide pour le moment.")