import streamlit as st
import pandas as pd
from datetime import date, datetime
from audit_log import enregistrer_modification, afficher_historique_modifications
import projets_config

def show(supabase):
    st.title("🚜 Essai à la Plaque (NF P 94-117-1)")

    # ---------------------------------------------------------
    # RÉCUPÉRATION DYNAMIQUE DU TECHNICIEN CONNECTÉ
    # ---------------------------------------------------------
    user_raw = (
        st.session_state.get("username") or 
        st.session_state.get("user") or 
        st.session_state.get("user_name") or 
        "Agent LPEE"
    )

    if isinstance(user_raw, dict):
        user_raw = user_raw.get("email") or user_raw.get("name") or "Agent LPEE"

    current_user = str(user_raw).upper()

    # Détection administrateur pour la suppression
    user_role = str(st.session_state.get("role", "")).upper()
    is_admin = st.session_state.get("is_admin", False) or user_role == "ADMIN"
    is_baallal_admin = current_user.strip() == "BAALLAL" and is_admin

    # Autoriser l'édition pour tous les utilisateurs connectés à l'application
    can_edit = True

    # Projet actif : toutes les lectures/écritures de cette page sont
    # limitées à ce projet, pour garantir l'étanchéité entre chantiers.
    user_info_projet = st.session_state.get("user") or {}
    projet_id_actif = projets_config.projet_actif(user_info_projet)
    if not projet_id_actif:
        st.error("⚠️ Aucun projet ne vous est autorisé. Contactez un administrateur.")
        return
    st.caption(f"📁 Projet actif : **{projets_config.nom_projet(projet_id_actif)}**")

    # ---------------------------------------------------------
    # 1. GESTION DU MOTEUR D'ÉDITION / MODIFICATION
    # ---------------------------------------------------------
    editing_item = st.session_state.get("edit_plaque_item", None)

    if editing_item:
        st.info(f"✏️ **Mode Modification** - Essai ID #{editing_item['id']}")
        
        default_date = datetime.strptime(editing_item["date_essai"], "%Y-%m-%d").date() if isinstance(editing_item.get("date_essai"), str) else date.today()
        default_client = editing_item.get("client", "TGCC")
        default_projet = editing_item.get("projet", "LGV CASA SUD")
        default_empl = editing_item.get("emplacement", "")
        default_pk = editing_item.get("pk_profil", editing_item.get("pkl", ""))
        default_couche = editing_item.get("couche", "Assise")
        default_mat = editing_item.get("nature_materiau", "")
        default_z1 = float(editing_item.get("z1", 0.53))
        default_z2 = float(editing_item.get("z2", 0.52))
        default_tech = editing_item.get("technicien", current_user)
        default_obs = editing_item.get("observations", "")
    else:
        default_date = date.today()
        default_client = "TGCC"
        default_projet = "LGV CASA SUD"
        default_empl = "Voie B"
        default_pk = "PK 1+200"
        default_couche = "Assise"
        default_mat = "GNT 0/31.5 Classée B2"
        default_z1 = 0.53
        default_z2 = 0.52
        default_tech = current_user
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
        couche_options = ["Arase", "Assise", "Remblai", "PST", "Couche de forme", "Autre"]
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
                sample_query = supabase.table("essais_plaque").select("*").limit(1).execute()
                if sample_query.data and len(sample_query.data) > 0:
                    valid_columns = set(sample_query.data[0].keys())
                    safe_payload = {k: v for k, v in payload.items() if k in valid_columns}
                else:
                    safe_payload = payload
                # Toujours inclus, même si absent des colonnes déjà vues par
                # l'échantillon ci-dessus (table encore vide par exemple).
                safe_payload["projet_id"] = projet_id_actif

                if editing_item:
                    anciennes_valeurs_plaque = {k: editing_item.get(k) for k in safe_payload}
                    supabase.table("essais_plaque").update(safe_payload).eq("id", editing_item["id"]).eq("projet_id", projet_id_actif).execute()
                    enregistrer_modification(
                        supabase,
                        table_concernee="essais_plaque",
                        enregistrement_id=editing_item["id"],
                        action="MODIFICATION",
                        anciennes_valeurs=anciennes_valeurs_plaque,
                        nouvelles_valeurs=safe_payload,
                    )
                    st.success(f"✅ Essai #{editing_item['id']} mis à jour avec succès !")
                    st.session_state["edit_plaque_item"] = None
                else:
                    res_ins_plaque = supabase.table("essais_plaque").insert(safe_payload).execute()
                    if res_ins_plaque.data:
                        nouvel_id_plaque = res_ins_plaque.data[0].get("id")
                        enregistrer_modification(
                            supabase,
                            table_concernee="essais_plaque",
                            enregistrement_id=nouvel_id_plaque,
                            action="CREATION",
                            nouvelles_valeurs=safe_payload,
                        )
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
    # 4. HISTORIQUE DES ESSAIS ET ACTIONS DE MODIFICATION
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📋 Historique des Essais Enregistrés")

    try:
        res = supabase.table("essais_plaque").select("*").eq("projet_id", projet_id_actif).order("id", desc=True).execute()
        if res.data and len(res.data) > 0:
            
            clean_rows = []
            for row in res.data:
                k_val = row.get("k_ratio") if row.get("k_ratio") is not None else row.get("k")
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

            df_display = pd.DataFrame(clean_rows)
            st.dataframe(df_display, use_container_width=True, hide_index=True)

            # --- ACTIONS DE SÉLECTION ET EDITION ---
            st.markdown("### ⚙️ Actions de Modification / Gestion")
            
            selected_id = st.selectbox(
                "Sélectionnez un essai par son ID :", 
                options=[item["id"] for item in res.data],
                key="admin_select_plaque_id"
            )

            if is_baallal_admin:
                afficher_historique_modifications(supabase, "essais_plaque", selected_id)

            if is_admin:
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
                            item_a_supprimer = next((item for item in res.data if item["id"] == selected_id), None)
                            enregistrer_modification(
                                supabase,
                                table_concernee="essais_plaque",
                                enregistrement_id=selected_id,
                                action="SUPPRESSION",
                                anciennes_valeurs={k: v for k, v in (item_a_supprimer or {}).items() if k != "id"},
                                commentaire="Suppression définitive de l'essai",
                            )
                            supabase.table("essais_plaque").delete().eq("id", selected_id).eq("projet_id", projet_id_actif).execute()
                            st.success(f"🗑️ Essai #{selected_id} supprimé avec succès.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur lors de la suppression : {e}")
            else:
                # Bouton de modification disponible pour les techniciens
                if st.button("✏️ Modifier cet essai", type="secondary", use_container_width=True):
                    selected_item = next((item for item in res.data if item["id"] == selected_id), None)
                    if selected_item:
                        st.session_state["edit_plaque_item"] = selected_item
                        st.rerun()

        else:
            st.info("Aucun essai à la plaque n'a encore été enregistré.")
    except Exception as e:
        st.warning(f"Impossible de charger l'historique : {e}")
