import streamlit as st
import pandas as pd
from datetime import date, datetime

def show(supabase):
    st.title("🚜 Essai à la Plaque (NF P 94-117-1)")

    # Vérification du rôle d'administrateur
    is_admin = st.session_state.get("is_admin", False) or st.session_state.get("role") == "admin"

    # ---------------------------------------------------------
    # 1. GESTION DU MOTEUR D'ÉDITION / MODIFICATION (ADMIN)
    # ---------------------------------------------------------
    editing_item = st.session_state.get("edit_plaque_item", None)

    if editing_item:
        st.info(f"✏️ **Mode Modification** - Essai ID #{editing_item['id']}")
        
        # Valeurs pré-remplies pour l'édition
        default_date = datetime.strptime(editing_item["date_essai"], "%Y-%m-%d").date() if isinstance(editing_item.get("date_essai"), str) else date.today()
        default_client = editing_item.get("client", "TGCC")
        default_projet = editing_item.get("projet", "LGV CASA SUD")
        default_empl = editing_item.get("emplacement", "")
        default_pk = editing_item.get("pk_profil", editing_item.get("pkl", ""))
        default_couche = editing_item.get("couche", "Assise")
        default_mat = editing_item.get("nature_materiau", "")
        default_z1 = float(editing_item.get("z1", 0.53))
        default_z2 = float(editing_item.get("z2", 0.52))
        default_tech = editing_item.get("technicien", "")
        default_obs = editing_item.get("observations", "")
    else:
        # Valeurs par défaut pour une nouvelle saisie
        default_date = date.today()
        default_client = "TGCC"
        default_projet = "LGV CASA SUD"
        default_empl = "Voie B"
        default_pk = "PK 1+200"
        default_couche = "Assise"
        default_mat = "GNT 0/31.5 Classée B2"
        default_z1 = 0.53
        default_z2 = 0.52
        default_tech = "Agent LPEE"
        default_obs = "Portance conforme aux exigences du CPT."

    # ---------------------------------------------------------
    # 2. FORMULAIRE DE SAISIE / ÉDITION
    # ---------------------------------------------------------
    st.subheader("📝 " + ("Modifier l'essai" if editing_item else "Saisie d'un nouvel essai"))

    col1, col2, col3 = st.columns(3)

    with col1:
        date_essai = st.date_input("Date de l'essai", value=default_date, key="plaque_date")
        client = st.text_input("Client / Organisme", value=default_client, key="plaque_client")
        projet = st.text_input("Chantier / Projet", value=default_projet, key="plaque_projet")
        
    with col2:
        emplacement = st.text_input("Emplacement / Zone", value=default_empl, key="plaque_empl")
        pk_profil = st.text_input("PK / Profil", value=default_pk, key="plaque_pk")
        couche_options = ["Assise", "Remblai", "PST", "Couche de forme", "Autre"]
        couche_idx = couche_options.index(default_couche) if default_couche in couche_options else 0
        couche = st.selectbox("Couche testée", couche_options, index=couche_idx, key="plaque_couche")
        nature_materiau = st.text_input("Nature du matériau", value=default_mat, key="plaque_mat")

    with col3:
        st.markdown("##### 📏 Données de Chargement (Enfoncements)")
        z1 = st.number_input("Z1 - 1er chargement (mm)", min_value=0.01, max_value=10.0, value=default_z1, step=0.01, format="%.2f", key="plaque_z1")
        z2 = st.number_input("Z2 - 2ème chargement (mm)", min_value=0.01, max_value=10.0, value=default_z2, step=0.01, format="%.2f", key="plaque_z2")
        technicien = st.text_input("Technicien LPEE", value=default_tech, key="plaque_tech")

    # Calculs automatiques (NF P 94-117-1)
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

    observations = st.text_area("Observations / Remarques", value=default_obs, key="plaque_obs")

    # ---------------------------------------------------------
    # 3. ENREGISTREMENT OU MISE À JOUR SÉCURISÉE
    # ---------------------------------------------------------
    btn_col1, btn_col2 = st.columns([3, 1])

    with btn_col1:
        button_label = "🔄 Mettre à jour l'essai" if editing_item else "💾 Enregistrer l'essai"
        if st.button(button_label, key="btn_enregistrer_plaque", type="primary", use_container_width=True):
            payload = {
                "date_essai": str(date_essai),
                "client": client,
                "projet": projet,
                "emplacement": emplacement,
                "pk_profil": pk_profil,
                "couche": couche,
                "nature_materiau": nature_materiau,
                "z1": float(z1),
                "z2": float(z2),
                "ev1": float(ev1),
                "ev2": float(ev2),
                "k_ratio": float(k_ratio),
                "technicien": technicien,
                "observations": observations
            }

            try:
                # Filtrage des colonnes valides
                sample_query = supabase.table("essais_plaque").select("*").limit(1).execute()
                if sample_query.data and len(sample_query.data) > 0:
                    valid_columns = set(sample_query.data[0].keys())
                    safe_payload = {k: v for k, v in payload.items() if k in valid_columns}
                else:
                    safe_payload = payload

                if editing_item:
                    supabase.table("essais_plaque").update(safe_payload).eq("id", editing_item["id"]).execute()
                    st.success(f"✅ Essai #{editing_item['id']} mis à jour avec succès !")
                    st.session_state["edit_plaque_item"] = None
                else:
                    supabase.table("essais_plaque").insert(safe_payload).execute()
                    st.success("✅ Essai enregistré avec succès !")

                st.rerun()

            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement : {e}")

    with btn_col2:
        if editing_item:
            if st.button("❌ Annuler l'édition", use_container_width=True):
                st.session_state["edit_plaque_item"] = None
                st.rerun()

    # ---------------------------------------------------------
    # 4. HISTORIQUE - COLONNES STRICTEMENT SÉLECTIONNÉES ET TRIÉES
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📋 Historique des Essais Enregistrés")

    try:
        res = supabase.table("essais_plaque").select("*").order("id", desc=True).execute()
        if res.data and len(res.data) > 0:
            
            # Reconstruction d'une liste de dictionnaires propre ligne par ligne
            clean_rows = []
            for row in res.data:
                # Récupération de la valeur K (gestion k_ratio ou k)
                k_val = row.get("k_ratio") if row.get("k_ratio") is not None else row.get("k")
                # Récupération de la valeur PK/Profil (gestion pk_profil ou pkl)
                pk_val = row.get("pk_profil") if row.get("pk_profil") is not None else row.get("pkl")

                clean_rows.append({
                    "ID": row.get("id"),
                    "Date d'essai": row.get("date_essai"),
                    "Client": row.get("client"),
                    "Projet": row.get("projet"),
                    "Emplacement": row.get("emplacement"),
                    "PK/profil": pk_val,
                    "Couche": row.get("couche"),
                    "Nature de matériaux": row.get("nature_materiau"),
                    "Z1": row.get("z1"),
                    "Z2": row.get("z2"),
                    "EV1": row.get("ev1"),
                    "EV2": row.get("ev2"),
                    "K": k_val,
                    "Technicien": row.get("technicien")
                })

            # Création directe du DataFrame avec l'ordre exact demandé
            df_display = pd.DataFrame(clean_rows)

            st.dataframe(df_display, use_container_width=True, hide_index=True)

            # --- ACTIONS RESERVÉES À L'ADMINISTRATEUR ---
            if is_admin:
                st.markdown("### ⚙️ Actions Administrateur (Modifier / Supprimer)")
                
                selected_id = st.selectbox(
                    "Sélectionnez un essai par son ID :", 
                    options=[item["id"] for item in res.data],
                    key="admin_select_plaque_id"
                )

                act_col1, act_col2 = st.columns(2)

                with act_col1:
                    if st.button("✏️ Modifier cet essai", type="secondary", use_container_width=True):
                        selected_item = next((item for item in res.data if item["id"] == selected_id), None)
                        if selected_item:
                            st.session_state["edit_plaque_item"] = selected_item
                            st.rerun()

                with act_col2:
                    if st.button("🗑️ Supprimer cet essai", type="primary", use_container_width=True):
                        try:
                            supabase.table("essais_plaque").delete().eq("id", selected_id).execute()
                            st.success(f"🗑️ Essai #{selected_id} supprimé avec succès.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur lors de la suppression : {e}")
            else:
                st.caption("🔒 *Connectez-vous en tant qu'administrateur pour modifier ou supprimer des enregistrements.*")

        else:
            st.info("Aucun essai à la plaque n'a encore été enregistré.")
    except Exception as e:
        st.warning(f"Impossible de charger l'historique : {e}")
