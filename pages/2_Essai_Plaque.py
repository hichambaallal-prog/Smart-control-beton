import streamlit as st
import pandas as pd
from datetime import date
from supabase import create_client, Client

st.set_page_config(page_title="Essai à la Plaque - LPEE", layout="wide")

# Connection Supabase
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
# ⚠️ Remplacez la chaîne ci-dessous par votre vraie clé ANON Supabase
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0" 

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

# Formulaire de Saisie
with st.form("form_plaque"):
    st.subheader(f"📝 Saisie Essai à la Plaque ({str_date_p})")
    
    # Répartition optimisée des colonnes pour intégrer Z1 et Z2
    c1, c2, c3 = st.columns(3)
    
    with c1:
        pk_emp = st.text_input("Emplacement / PK", value="PK 14+250 - Voie 1")
        couche_elem = st.selectbox("Couche / Support", ["PFT3 (Couche de Forme)", "PST (Arase)", "Couche d'Assise", "Remblai"])
        
    with c2:
        # Saisie des données de base Z1 et Z2 (min_value empêche la division par zéro)
        z1 = st.number_input("1er chargement Z1 (mm)", min_value=0.01, value=1.46, step=0.05)
        z2 = st.number_input("2ème chargement Z2 (mm)", min_value=0.01, value=1.02, step=0.05)
        
    with c3:
        # Calculs automatiques
        ev1 = 112.5 / (z1 * 2)
        ev2 = 90 / (z2 * 2)
        rapport_calc = round(ev2 / ev1, 2) if ev1 > 0 else 0.0
        
        # Affichage dynamique des résultats sous forme de métriques
        st.metric("Module EV1 (MPa)", value=f"{ev1:.1f}")
        st.metric("Module EV2 (MPa)", value=f"{ev2:.1f}")
        st.metric("Rapport k = EV2 / EV1", value=rapport_calc)
        
        is_conforme = (ev2 >= 50.0) and (rapport_calc <= 2.2)
        statut_auto = "✅ Conforme" if is_conforme else "⚠️ Non Conforme"
        st.info(f"Statut : **{statut_auto}**")

    obs_p = st.text_area("Observations", value="RAS - Sol bien compacté")
    
    # Bouton d'enregistrement propre au formulaire
    submitted = st.form_submit_button("💾 Enregistrer l'essai à la plaque", type="primary")
    
    if submitted:
        row_p = {
            "date_essai":Voici le code mis à jour avec l'intégration des enfoncements **Z1** et **Z2**, ainsi que le calcul automatique des modules $EV_1$, $EV_2$ et du rapport $K$ selon vos formules. 

J'ai structuré les nouvelles entrées dans la colonne `c2` et les résultats calculés dans la colonne `c3`.

### 💻 Code mis à jour

```python
import streamlit as st
import pandas as pd
from datetime import date
from supabase import create_client, Client

st.set_page_config(page_title="Essai à la Plaque - LPEE", layout="wide")

# Connection Supabase
URL = "[https://yqijsvxyrdymcnqluipa.supabase.co](https://yqijsvxyrdymcnqluipa.supabase.co)"
# ⚠️ Remplacez la chaîne ci-dessous par votre vraie clé ANON Supabase
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0" 

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

# Formulaire de Saisie
with st.form("form_plaque"):
    st.subheader(f"📝 Saisie Essai à la Plaque ({str_date_p})")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        pk_emp = st.text_input("Emplacement / PK", value="PK 14+250 - Voie 1")
        couche_elem = st.selectbox("Couche / Support", ["PFT3 (Couche de Forme)", "PST (Arase)", "Couche d'Assise", "Remblai"])
        
    with c2:
        # Nouvelles entrées pour les enfoncements Z1 et Z2
        z1 = st.number_input("Enfoncement 1er chargement Z1 (mm)", value=1.50, step=0.01)
        z2 = st.number_input("Enfoncement 2ème chargement Z2 (mm)", value=0.50, step=0.01)
        
    with c3:
        # Calculs automatiques (avec prévention de la division par zéro)
        ev1 = round(112.5 / (z1 * 2), 2) if z1 > 0 else 0.0
        ev2 = round(90.0 / (z2 * 2), 2) if z2 > 0 else 0.0
        rapport_calc = round(ev2 / ev1, 2) if ev1 > 0 else 0.0
        
        # Affichage des résultats calculés
        st.metric("EV1 (MPa)", value=ev1)
        st.metric("EV2 (MPa)", value=ev2)
        st.metric("Rapport k = EV2 / EV1", value=rapport_calc)
        
        is_conforme = (
