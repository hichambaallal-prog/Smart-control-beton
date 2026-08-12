import streamlit as st
import pandas as pd
from datetime import date, datetime

def show(supabase):
    st.title("🚜 Essai à la Plaque (NF P 94-117-1)")

    # ---------------------------------------------------------
    # 1. FORMULAIRE DE SAISIE
    # ---------------------------------------------------------
    st.subheader("📝 Saisie d'un nouvel essai")

    col1, col2, col3 = st.columns(3)

    with col1:
        date_essai = st.date_input("Date de l'essai", value=date.today(), key="plaque_date")
        client = st.text_input("Client / Organisme", value="TGCC", key="plaque_client")
        projet = st.text_input("Chantier / Projet", value="LGV CASA SUD", key="plaque_projet")
        
    with col2:
        emplacement = st.text_input("Emplacement / Zone", value="Voie B", key="plaque_empl")
        pk_profil = st.text_input("PK / Profil", value="PK 1+200", key="plaque_pk")
        couche = st.selectbox(
            "Couche testée", 
            ["Assise", "Remblai", "PST", "Couche de forme", "Autre"], 
            key="plaque_couche"
        )
        nature_materiau = st.text_input("Nature du matériau", value="GNT 0/31.5 Classée B2", key="plaque_mat")

    with col3:
        st.markdown("##### 📏 Données de Chargement (Enfoncements)")
        z1 = st.number_input("Z1 - 1er chargement (mm)", min_value=0.01, max_value=10.0, value=0.53, step=0.01, format="%.2f", key="plaque_z1")
        z2 = st.number_input("Z2 - 2ème chargement (mm)", min_value=0.01, max_value=10.0, value=0.52, step=0.01, format="%.2f", key="plaque_z2")
        technicien = st.text_input("Technicien LPEE", value="Agent LPEE", key="plaque_tech")

    # ---------------------------------------------------------
    # 2. CALCULS AUTOMATIQUES (NF P 94-117-1)
    # ---------------------------------------------------------
    # EV1 = 112.5 / (Z1 * 2)
    # EV2 = 90 / (Z2 * 2)
    # K = EV2 / EV1
    ev1 = round(112.5 / (z1 * 2), 2) if z1 > 0 else 0.0
    ev2 = round(90.0 / (z2 * 2), 2) if z2 > 0 else 0.0
    k_ratio = round(ev2 / ev1, 2) if ev1 > 0 else 0.0

    st.markdown("---")
    st.subheader("📈 Résultats Calculés Automatiquement")
    
    res_col1, res_col2, res_col3 = st.columns(3)
    res_col1.metric("EV1 (MPa)", f"{ev1:.2f}")
    res_col2.metric("EV2 (MPa)", f"{ev2:.2f}")
    
    k_delta = "Conforme (K ≤ 2.0)" if k_ratio <= 2.0 else "Attention (K > 2.0)"
    res_col3.metric("Coefficient K (EV2/EV1)", f"{k_ratio:.2f}", delta=k_delta, delta_color="normal" if k_ratio <= 2.0 else "inverse")

    observations = st.text_area("Observations / Remarques", value="Portance conforme aux exigences du CPT.", key="plaque_obs")

    # ---------------------------------------------------------
    # 3. ENREGISTREMENT SÉCURISÉ
    # ---------------------------------------------------------
    if st.button("💾 Enregistrer l'essai", key="btn_enregistrer_plaque", type="primary", use_container_width=True):
        payload = {
            "date_essai": str(date_essai),
            "client": client,
            "projet": projet,
            "emplacement": emplacement,
            "pk_profil": pk_profil,
            "pkl": pk_profil,
            "couche": couche,
            "nature_materiau": nature_materiau,
            "z1": float(z1),
            "z2": float(z2),
            "ev1": float(ev1),
            "ev2": float(ev2),
            "k": float(k_ratio),
            "k_ratio": float(k_ratio),
            "technicien": technicien,
            "observations": observations
        }

        try:
            sample_query = supabase.table("essais_plaque").select("*").limit(1).execute()
            
            if sample_query.data and len(sample_query.data) > 0:
                valid_columns = set(sample_query.data[0].keys())
                safe_payload = {k: v for k, v in payload.items() if k in valid_columns}
            else:
                safe_payload = payload

            supabase.table("essais_plaque").insert(safe_payload).execute()
            st.success("✅ Essai enregistré avec succès !")
            st.rerun()

        except Exception as e:
            st.error(f"Erreur lors de l'enregistrement : {e}")

    # ---------------------------------------------------------
    # 4. HISTORIQUE DES ESSAIS
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📋 Historique des Essais Enregistrés")

    try:
        res = supabase.table("essais_plaque").select("*").order("id", desc=True).execute()
        if res.data and len(res.data) > 0:
            df = pd.DataFrame(res.data)
            
            # Nettoyage des colonnes techniques d'affichage
            cols_to_drop = [c for c in ["created_at"] if c in df.columns]
            df_display = df.drop(columns=cols_to_drop)

            st.dataframe(df_display, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun essai à la plaque n'a encore été enregistré.")
    except Exception as e:
        st.warning(f"Impossible de charger l'historique : {e}")
