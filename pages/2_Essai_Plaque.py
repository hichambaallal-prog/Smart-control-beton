import streamlit as st
import pandas as pd
from datetime import date
from supabase import create_client, Client

st.set_page_config(page_title="Essai à la Plaque - LPEE", layout="wide")

# Connection Supabase
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1"  # ⚠️ Remettez votre vraie clé Supabase

try:
    supabase: Client = create_client(URL, CLE)
except Exception as e:
    st.error(f"Erreur Supabase : {e}")
    st.stop()

st.title("🪨 Contrôle de Portance - Essai à la Plaque (NF P94-117-1)")
st.markdown("### Laboratoire Public d'Essais et d'Études (LPEE)")

# Chargement données
try:
    resp = supabase.table("essais_plaque").select("*").execute()
    data_all_plaque = resp.data or []
except Exception as e:
    st.error(f"Erreur de chargement : {e}")
    data_all_plaque = []

date_choisie_p = st.date_input("📅 Date de l'essai :", value=date.today())
str_date_p = date_choisie_p.strftime("%d/%m/%Y")

with st.form("form_plaque"):
    st.subheader(f"📝 Saisie Essai à la Plaque ({str_date_p})")
    c1, c2, c3 = st.columns(3)
    with c1:
        pk_emp = st.text_input("Emplacement / PK", value="PK 14+250 - Voie 1")
        couche_elem = st.selectbox("Couche / Support", ["PFT3 (Couche de Forme)", "PST (Arase)", "Couche d'Assise", "Remblai"])
    with c2:
        ev1 = st.number_input("Module EV1 (MPa)", value=38.5, step=0.1)
        ev2 = st.number_input("Module EV2 (MPa)", value=88.0, step=0.1)
    with c3:
        rapport_calc = round(ev2 / ev1, 2) if ev1 > 0 else 0
        st.metric("Rapport k = EV2 / EV1", value=rapport_calc)
        
        is_conforme = (ev2 >= 50.0) and (rapport_calc <= 2.2)
        statut_auto = "✅ Conforme" if is_conforme else "⚠️ Non Conforme"
        st.info(f"Statut : **{statut_auto}**")

    obs_p = st.text_area("Observations", value="RAS - Sol bien compacté")
    
    if st.form_submit_button("💾 Enregistrer l'essai à la plaque"):
        row_p = {
            "date_essai": str_date_p,
            "pk_emplacement": pk_emp,
            "couche_element": couche_elem,
            "ev1": float(ev1),
            "ev2": float(ev2),
            "rapport_ev2_ev1": float(rapport_calc),
            "statut": statut_auto,
            "observations": obs_p
        }
        supabase.table("essais_plaque").insert(row_p).execute()
        st.success("Essai enregistré avec succès !")
        st.rerun()

st.markdown("---")
st.subheader(f"📋 Historique des Essais à la Plaque")
if data_all_plaque:
    df_p = pd.DataFrame(data_all_plaque)
    df_p.index = range(1, len(df_p) + 1)
    st.dataframe(df_p, use_container_width=True)
