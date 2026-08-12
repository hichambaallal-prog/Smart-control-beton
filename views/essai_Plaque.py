import streamlit as st
import pandas as pd
from datetime import date, datetime

def show(supabase):
    st.title("🚜 Essai à la Plaque (NF P 94-117-1)")

    # ---------------------------------------------------------
    # 1. FORMULAIRE DE SAISIE
    # ---------------------------------------------------------
    st.subheader("Saisie d'un essai")

    col1, col2, col3 = st.columns(3)

    with col1:
        date_essai = st.date_input("Date de l'essai", value=date.today(), key="plaque_date")
        client = st.text_input("Client / Organisme", value="TGCC", key="plaque_client")
        projet = st.text_input("Chantier / Projet", value="Aménagement Boulevard Zerktouni", key="plaque_projet")
        emplacement = st.text_input("Emplacement / Zone", value="Voie B - PK 1+200", key="plaque_empl")

    with col2:
        norme = st.selectbox("Norme de référence", ["NF P 94-117-1", "LPEE-CTR-CSB"], key="plaque_norme")
        pk_profil = st.text_input("PK / Profil", value="PK 1+200", key="plaque_pk")
        couche = st.selectbox("Couche testée", ["Forme (PST)", "Fondation (GNT)", "Base", "Soustraitment"], key="plaque_couche")
        nature_materiau = st.text_input("Nature du matériau", value="GNT 0/31.5 Classée B2", key="plaque_mat")

    with col3:
        st.markdown("##### 📏 Données de Chargement (Enfoncements)")
        z1 = st.number_input("Z1 - 1er chargement (mm)", min_value=0.01, max_value=10.0, value=0.53, step=0.01, format="%.2f", key="plaque_z1")
        z2 = st.number_input("Z2 - 2ème chargement (mm)", min_value=0.01, max_value=10.0, value=0.52, step=0.01, format="%.2f", key="plaque_z2")
        technicien = st.text_input("Technicien LPEE", value="Agent LPEE", key="plaque_tech")

    # ---------------------------------------------------------
    # 2. CALCULS AUTOMATIQUES (Norme NF P 94-117-1 : EV = 112.5 / z)
    # ---------------------------------------------------------
    ev1 = round(112.5 / z1, 2) if z1 > 0 else 0.0
    ev2 = round(112.5 / z2, 2) if z2 > 0 else 0.0
    k_ratio = round(ev2 / ev1, 2) if ev1 > 0 else 0.0

    st.markdown("---")
    st.subheader("📈 Résultats Calculés Automatiquement")
    
    res_col1, res_col2, res_col3 = st.columns(3)
    res_col1.metric("EV1 (MPa)", f"{ev1:.2f}")
    res_col2.metric("EV2 (MPa)", f"{ev2:.2f}")
    res_col3.metric("Coefficient K (EV2/EV1)", f"{k_ratio:.2f}")

    observations = st.text_area("Observations / Remarques", value="Portance conforme aux exigences du CPT.", key="plaque_obs")

    # ---------------------------------------------------------
    # 3. ENREGISTREMENT SÉCURISÉ (FILTRAGE DYNAMIQUE DES COLONNES)
    # ---------------------------------------------------------
    if st.button("💾 Enregistrer l'essai", key="btn_enregistrer_plaque"):
        # Dictionnaire complet de toutes les données saisies
        payload = {
            "date_essai": str(date_essai),
            "client": client,
            "projet": projet,
            "emplacement": emplacement,
            "norme": norme,
            "pk_profil": pk_profil,
            "pkl": pk_profil, # Pour compatibilité
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
            # 1. On interroge Supabase avec 1 ligne pour détecter les colonnes réellement existantes
            sample_query = supabase.table("essais_plaque").select("*").limit(1).execute()
            
            # 2. Si la table a des colonnes retournées ou définies dans le premier enregistrement
            if sample_query.data and len(sample_query.data) > 0:
                valid_columns = set(sample_query.data[0].keys())
                # Filtrer le payload pour ne garder QUE les colonnes existantes
                safe_payload = {k: v for k, v in payload.items() if k in valid_columns}
            else:
                # Si la table est totalement vide, on fait un essai en retirant les clés problématiques récurrentes si besoin
                safe_payload = payload

            # 3. Insertion
            supabase.table("essais_plaque").insert(safe_payload).execute()
            st.success("✅ Essai enregistré avec succès !")
            st.rerun()

        except Exception as e:
            err_msg = str(e)
            st.error(f"Erreur d'enregistrement : {err_msg}")
            st.info("💡 Résolution rapide : Exécutez le script SQL ci-dessous dans Supabase SQL Editor pour ajouter toutes les colonnes manquantes.")

    # ---------------------------------------------------------
    # 4. HISTORIQUE & AFFICHER LA SYNTHÈSE
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📋 Historique des Essais à la Plaque Enregistrés")

    try:
        res = supabase.table("essais_plaque").select("*").order("id", desc=True).execute()
        if res.data and len(res.data) > 0:
            df = pd.DataFrame(res.data)
            
            # Nettoyage des colonnes techniques
            cols_to_drop = [c for c in ["id", "created_at"] if c in df.columns]
            df_display = df.drop(columns=cols_to_drop)

            # Re-numérotation lisible
            df_display.index = range(1, len(df_display) + 1)
            st.dataframe(df_display, use_container_width=True)
        else:
            st.info("Aucun essai à la plaque n'a encore été enregistré.")
    except Exception as e:
        st.warning(f"Impossible de charger l'historique : {e}")
