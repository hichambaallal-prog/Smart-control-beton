import pandas as pd
import streamlit as st

# =========================================================
# PHASE 2 : SAISIE DES ÉCRASEMENTS ET PV
# =========================================================
with tab_saisie:
    st.subheader("💥 2. Saisie Groupée & Édition des PV d'Écrasement")

    eprouvettes_en_attente = []
    try:
        res_att = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .order("id", desc=False)
            .execute()
        )
        if res_att.data:
            eprouvettes_en_attente = [
                e
                for e in res_att.data
                if e.get("force_kn") is None
                or float(e.get("force_kn") or 0) == 0
            ]
    except Exception as e:
        st.error(f"Erreur de chargement des essais en attente : {e}")

    if not eprouvettes_en_attente:
        st.info("👍 Aucune éprouvette en attente d'écrasement.")
    else:
        # Groupement par lot de bétonnage / échéance / ouvrage
        groupes_lots = {}
        for ep in eprouvettes_en_attente:
            b_id_ep = ep.get("betonnage_id")
            ech_ep = ep.get("echeance", "28 jours")
            ouv_ep = ep.get("ouvrage", "N/A")
            dt_ecras = ep.get("date_ecrasement", "N/A")

            cle_groupe = (
                f"Ouvrage: {ouv_ep} | Échéance: {ech_ep} (Date: {dt_ecras})"
                f" | Bétonnage ID #{b_id_ep}"
            )

            if cle_groupe not in groupes_lots:
                groupes_lots[cle_groupe] = []
            groupes_lots[cle_groupe].append(ep)

        choix_lot = st.selectbox(
            "📦 Sélectionner le lot d'éprouvettes :",
            list(groupes_lots.keys()),
            key="select_lot_saisie",
        )
        lot_selected = groupes_lots[choix_lot]
        sample = lot_selected[0]

        # Résumé visuel du lot
        col_l1, col_l2, col_l3, col_l4 = st.columns(4)
        col_l1.metric("Client", str(sample.get("client", "TGCC")))
        col_l2.metric("Projet", "LGV CASA")
        col_l3.metric("Ouvrage", str(sample.get("ouvrage", "N/A")))
        col_l4.metric("Échéance Visée", str(sample.get("echeance", "28 jours")))

        st.markdown("---")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            tech_global = st.text_input(
                "Technicien / Opérateur",
                value="Technicien LPEE",
                key="tech_global",
            )
        with col_g2:
            obs_globale = st.text_input(
                "Commentaire / Observation",
                value="PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES",
                key="obs_global",
            )

        st.markdown("##### 📝 Saisie des mesures pour le lot")

        lot_key = f"df_lot_{choix_lot}"

        if lot_key not in st.session_state:
            rows_list = []
            for ep in lot_selected:
                sec = float(ep.get("section") or 176.71)
                f_kn = float(ep.get("force_kn") or 0.0)
                fc = (
                    round((f_kn * 10.0) / sec, 1) if sec > 0 and f_kn > 0 else 0.0
                )

                rows_list.append({
                    "ID": ep["id"],
                    "Repère": ep.get(
                        "repere_eprouvette", f"EPR-{ep['id']}"
                    ),
                    "Forme d'éprouvette": str(
                        ep.get("forme") or "Cylindrique 150x300"
                    ),
                    "_section": sec,
                    "Force (kN)": f_kn,
                    "Résistance Fc (MPa)": fc,
                })
            st.session_state[lot_key] = pd.DataFrame(rows_list)

        def update_fc():
            changes = st.session_state.data_editor_ecrasement.get(
                "edited_rows", {}
            )
            for row_idx, updated_cols in changes.items():
                if "Force (kN)" in updated_cols:
                    new_force = float(updated_cols["Force (kN)"] or 0.0)
                    sec = float(
                        st.session_state[lot_key].at[row_idx, "_section"]
                    )
                    st.session_state[lot_key].at[row_idx, "Force (kN)"] = (
                        new_force
                    )
                    if sec > 0 and new_force > 0:
                        st.session_state[lot_key].at[
                            row_idx, "Résistance Fc (MPa)"
                        ] = round((new_force * 10.0) / sec, 1)
                    else:
                        st.session_state[lot_key].at[
                            row_idx, "Résistance Fc (MPa)"
                        ] = 0.0

        st.data_editor(
            st.session_state[lot_key],
            column_config={
                "ID": st.column_config.NumberColumn("ID", disabled=True),
                "Repère": st.column_config.TextColumn("Repère", disabled=True),
                "Forme d'éprouvette": st.column_config.TextColumn(
                    "Forme d'éprouvette", disabled=True
                ),
                "_section": None,  # Masqué dans l'affichage
                "Force (kN)": st.column_config.NumberColumn(
                    "⚡ Force (kN)",
                    help="Saisissez la force maximale à la rupture en kN",
                    min_value=0.0,
                    max_value=3000.0,
                    step=0.1,
                    format="%.1f",
                ),
                "Résistance Fc (MPa)": st.column_config.NumberColumn(
                    "💥 Résistance Fc (MPa)", disabled=True, format="%.1f"
                ),
            },
            use_container_width=True,
            hide_index=True,
            key="data_editor_ecrasement",
            on_change=update_fc,
        )

        df_actuel = st.session_state[lot_key]
        forces_valides = df_actuel[df_actuel["Force (kN)"] > 0]

        if not forces_valides.empty:
            fc_moy = round(forces_valides["Résistance Fc (MPa)"].mean(), 1)
            st.success(f"📈 **Résistance moyenne du lot : {fc_moy:.1f} MPa**")

        # --- PRÉPARATION DU FICHIER PV EXCEL ---
        export_data = []
        for _, row in df_actuel.iterrows():
            export_data.append({
                "repere_eprouvette": row["Repère"],
                "forme": row["Forme d'éprouvette"],
                "section": row["_section"],
                "force_kn": row["Force (kN)"],
                "fc_mpa": row["Résistance Fc (MPa)"],
                "date_essai": sample.get("date_ecrasement", "N/A"),
                "age": (
                    str(sample.get("echeance", "28"))
                    .replace(" jours", "")
                    .replace("j", "")
                ),
            })

        infos_header = {
            "re_num": sample.get("re_num", "25/260/LGV/ B/01"),
            "dossier": sample.get("dossier", "2025-260-05985-2025-0247"),
            "client": sample.get("client", "TGCC"),
            "num_bl": sample.get("num_bl", "15479"),
            "ouvrage": sample.get("ouvrage", "N/A"),
            "classe_beton": sample.get("classe_beton", "C30/37"),
            "date_coulee": sample.get("date_coulee", "02/06/2025"),
            "affaissement": sample.get("affaissement", "200"),
            "temperature": sample.get("temperature", "31"),
            "observations": obs_globale,
        }

        excel_file = generer_pv_excel(export_data, infos_header)
        filename = f"PV_Ecrasement_LPEE_{sample.get('num_bl', 'BL')}.xlsx"

        st.markdown("---")
        col_b1, col_b2 = st.columns(2)

        with col_b1:
            btn_enregistrer = st.button(
                "💾 Valider et Enregistrer Tout le Lot",
                type="primary",
                use_container_width=True,
            )

        with col_b2:
            st.download_button(
                label="📄 Télécharger le PV d'écrasement (Format LPEE)",
                data=excel_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if btn_enregistrer:
            if (df_actuel["Force (kN)"] == 0).any():
                st.error(
                    "❌ Une ou plusieurs forces d'écrasement sont à 0.0 kN."
                    " Veuillez saisir toutes les forces avant validation."
                )
            else:
                succes_lot = 0
                for _, row in df_actuel.iterrows():
                    update_payload = {
                        "force_kn": float(row["Force (kN)"]),
                        "fc_mpa": float(row["Résistance Fc (MPa)"]),
                        "technicien": tech_global,
                        "observations": obs_globale,
                    }
                    try:
                        supabase.table("suivi_controle_beton").update(
                            update_payload
                        ).eq("id", int(row["ID"])).execute()
                        succes_lot += 1
                    except Exception as e:
                        st.error(
                            f"Erreur sur l'éprouvette {row['Repère']} : {e}"
                        )

                if succes_lot == len(df_actuel):
                    st.balloons()
                    st.success(
                        f"✅ Lot de {succes_lot} éprouvettes enregistré avec succès !"
                    )
                    st.rerun()


# =========================================================
# HISTORIQUE & RE-TÉLÉCHARGEMENT DES PV
# =========================================================
with tab_hist:
    st.subheader("📋 Historique Général & Re-téléchargement des PV")
    try:
        res_all = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .order("id", desc=True)
            .execute()
        )
        if res_all.data:
            df_all = pd.DataFrame(res_all.data)

            # Extraction des essais validés (force_kn > 0)
            df_valides = df_all[
                (df_all["force_kn"].notnull()) & (df_all["force_kn"] > 0)
            ].copy()

            if not df_valides.empty:
                st.markdown("##### 📥 Re-télécharger un PV déjà validé")

                # Regroupement par lot d'écrasement
                groupes_valides = {}
                for _, row in df_valides.iterrows():
                    b_id_ep = row.get("betonnage_id")
                    ech_ep = row.get("echeance", "28 jours")
                    ouv_ep = row.get("ouvrage", "N/A")
                    dt_ecras = row.get("date_ecrasement", "N/A")

                    cle_pv = (
                        f"Ouvrage: {ouv_ep} | Échéance: {ech_ep}"
                        f" (Date: {dt_ecras}) | Lot ID #{b_id_ep}"
                    )

                    if cle_pv not in groupes_valides:
                        groupes_valides[cle_pv] = []
                    groupes_valides[cle_pv].append(row.to_dict())

                choix_pv_hist = st.selectbox(
                    "Sélectionnez le PV validé à re-télécharger :",
                    list(groupes_valides.keys()),
                    key="select_pv_hist",
                )

                lot_hist = groupes_valides[choix_pv_hist]
                sample_h = lot_hist[0]

                export_data_h = []
                for item in lot_hist:
                    sec = float(item.get("section") or 176.71)
                    f_kn = float(item.get("force_kn") or 0.0)
                    fc = float(item.get("fc_mpa") or 0.0)

                    export_data_h.append({
                        "repere_eprouvette": item.get(
                            "repere_eprouvette", f"EPR-{item['id']}"
                        ),
                        "forme": item.get("forme", "Cylindrique 150x300"),
                        "section": sec,
                        "force_kn": f_kn,
                        "fc_mpa": fc,
                        "date_essai": item.get("date_ecrasement", "N/A"),
                        "age": (
                            str(item.get("echeance", "28"))
                            .replace(" jours", "")
                            .replace("j", "")
                        ),
                    })

                infos_header_h = {
                    "re_num": sample_h.get("re_num", "25/260/LGV/ B/01"),
                    "dossier": sample_h.get(
                        "dossier", "2025-260-05985-2025-0247"
                    ),
                    "client": sample_h.get("client", "TGCC"),
                    "num_bl": sample_h.get("num_bl", "15479"),
                    "ouvrage": sample_h.get("ouvrage", "N/A"),
                    "classe_beton": sample_h.get("classe_beton", "C30/37"),
                    "date_coulee": sample_h.get("date_coulee", "N/A"),
                    "affaissement": sample_h.get("affaissement", "200"),
                    "temperature": sample_h.get("temperature", "31"),
                    "observations": sample_h.get(
                        "observations",
                        "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES",
                    ),
                }

                excel_pv_hist = generer_pv_excel(
                    export_data_h, infos_header_h
                )
                file_name_h = f"PV_Ecrasement_RE-EXPORT_{sample_h.get('num_bl', 'BL')}.xlsx"

                st.download_button(
                    label="📄 Télécharger le PV ré-énoncé (Format Excel LPEE)",
                    data=excel_pv_hist,
                    file_name=file_name_h,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_download_hist",
                )

            st.markdown("---")
            st.markdown("##### 📊 Base de données complète")
            st.dataframe(df_all, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun enregistrement d'écrasement trouvé dans la base.")
    except Exception as e:
        st.error(f"Erreur lors du chargement de l'historique : {e}")
