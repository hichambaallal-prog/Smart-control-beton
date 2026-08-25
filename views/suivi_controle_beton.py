# =========================================================
    # PHASE 2 : PLANNING & SAISIE DES ÉCRASEMENTS
    # =========================================================
    elif onglet_courant == OPTIONS_ONGLETS[2]:
        st.subheader("💥 2. Planning des Échéances & Saisie des Écrasements")

        # Banner d'accès direct QR Code
        scan_rec = st.session_state.get("pending_qr_rec")
        scan_b_id = st.session_state.get("pending_qr_bid")
        if scan_rec or scan_b_id:
            col_qr_info, col_qr_btn = st.columns([4, 1])
            col_qr_info.success(f"🎯 **Accès direct QR Code actif** (Réception / ID : `{scan_rec or scan_b_id}`)")
            if col_qr_btn.button("✖ Revenir au normal", use_container_width=True):
                st.session_state.pop("pending_qr_rec", None)
                st.session_state.pop("pending_qr_bid", None)
                st.rerun()

        today_date = date.today()
        today_str = str(today_date)

        date_filtre = st.date_input("📅 Choisir une date à consulter", value=today_date, key="filtre_date_planning")
        date_filtre_str = str(date_filtre)
        debut_semaine = date_filtre - timedelta(days=date_filtre.weekday())
        fin_semaine = debut_semaine + timedelta(days=6)

        # Retards & alertes échéances
        try:
            res_retards = (
                supabase.table("suivi_controle_beton")
                .select("*")
                .lte("date_ecrasement", today_str)
                .or_("force_kn.is.null,force_kn.eq.0")
                .order("date_ecrasement", desc=False)
                .execute()
            )
            retards_list = res_retards.data or []
        except Exception as e:
            retards_list = []
            st.warning(f"Note retards : {e}")

        if retards_list:
            st.error(f"🚨 **ATTENTION : {len(retards_list)} éprouvette(s) non écrasée(s) ont atteint ou dépassé leur date d'échéance !**")
            rows_retard = []
            for ep in retards_list:
                row = _format_ep_row(ep, date_ref=today_date)
                dt_e = datetime.strptime(str(ep.get("date_ecrasement"))[:10], "%Y-%m-%d").date() if ep.get("date_ecrasement") else today_date
                priorite = f"⚠️ En Retard ({(today_date - dt_e).days} jour(s))" if dt_e < today_date else "🔥 Prévu Aujourd'hui"
                rows_retard.append({
                    "Priorité": priorite,
                    "Date Écrasement Prévue": row["Date Écrasement Prévue"],
                    "Référence / Repère": row["Référence / Repère"],
                    "N° BL": row["N° BL"],
                    "Ouvrage": row["Ouvrage"],
                    "Classe Béton": row["Classe Béton"],
                    "Date Coulée": row["Date Coulée"],
                    "Échéance Visée": row["Échéance Visée"],
                    "Âge Actuel Réel": row["Âge Théorique"],
                })
            st.dataframe(pd.DataFrame(rows_retard), use_container_width=True, hide_index=True)
            st.markdown("---")

        # Planning Jour & Semaine
        try:
            eprouvettes_date_sel = supabase.table("suivi_controle_beton").select("*").eq("date_ecrasement", date_filtre_str).order("id", desc=False).execute().data or []
            eprouvettes_semaine = supabase.table("suivi_controle_beton").select("*").gte("date_ecrasement", str(debut_semaine)).lte("date_ecrasement", str(fin_semaine)).order("date_ecrasement", desc=False).execute().data or []
        except Exception as e:
            eprouvettes_date_sel, eprouvettes_semaine = [], []
            st.warning(f"Note chargement planning : {e}")

        df_sel = pd.DataFrame([_format_ep_row(ep) for ep in eprouvettes_date_sel])
        df_semaine = pd.DataFrame([_format_ep_row(ep) for ep in eprouvettes_semaine])

        with st.expander(f"📆 Éprouvettes programmées spécifiquement pour le : {date_filtre_str} ({len(eprouvettes_date_sel)} éprouvette(s))", expanded=True):
            if not df_sel.empty:
                st.dataframe(df_sel, use_container_width=True, hide_index=True)
            else:
                st.info(f"ℹ️ Aucune éprouvette programmée pour le {date_filtre_str}.")

            st.markdown("---")
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                if not df_sel.empty:
                    st.download_button(
                        f"📊 Télécharger la liste du jour ({date_filtre_str})",
                        exporter_dataframe_excel(df_sel, date_filtre_str),
                        file_name=f"Planning_Ecrasement_{date_filtre_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_download_planning_excel",
                    )
            with col_exp2:
                if not df_semaine.empty:
                    st.download_button(
                        f"📅 Télécharger la liste de la semaine ({debut_semaine} au {fin_semaine})",
                        exporter_dataframe_excel(df_semaine, f"Sem_{debut_semaine}"),
                        file_name=f"Planning_Semaine_{debut_semaine}_au_{fin_semaine}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_download_planning_semaine_excel",
                    )

        st.markdown("---")

        try:
            res_att = supabase.table("suivi_controle_beton").select("*").order("id", desc=False).execute()
            eprouvettes_en_attente = res_att.data if mode_admin else [e for e in (res_att.data or []) if e.get("force_kn") is None or float(e.get("force_kn") or 0) == 0]
        except Exception as e:
            eprouvettes_en_attente = []
            st.error(f"Erreur de chargement des essais : {e}")

        if not eprouvettes_en_attente:
            st.info("👍 Aucune éprouvette en attente de saisie.")
        else:
            groupes_lots = {}
            index_selectionne = 0
            cache_betonnage = {}  # Cache local pour limiter les requêtes Supabase

            for ep in eprouvettes_en_attente:
                b_id_ep = ep.get("betonnage_id")
                if b_id_ep not in cache_betonnage:
                    cache_betonnage[b_id_ep] = obtenir_infos_betonnage_parent(supabase, b_id_ep)
                info_b_temp = cache_betonnage[b_id_ep]

                ref_ctrl = determiner_ref_controle(supabase, b_id_ep, info_b_temp, ep)
                num_rec_parent = str((info_b_temp or {}).get("num_reception") or "").strip()
                classe_ep = ep.get("classe_beton") or (info_b_temp.get("classe_beton") if info_b_temp else "-")
                cle_groupe = f"Référence : {ref_ctrl} | Classe : {classe_ep} | Ouvrage : {ep.get('ouvrage', '-')} | Échéance : {ep.get('echeance', '28 jours')} (Date Prévue : {ep.get('date_ecrasement', '-')}) | Lot ID #{b_id_ep}"

                if cle_groupe not in groupes_lots:
                    if scan_rec and (str(scan_rec).strip().lower() in ref_ctrl.lower() or str(scan_rec).strip().lower() in num_rec_parent.lower()):
                        index_selectionne = len(groupes_lots)
                    elif scan_b_id and str(scan_b_id).strip() == str(b_id_ep).strip():
                        index_selectionne = len(groupes_lots)
                    groupes_lots[cle_groupe] = []

                groupes_lots[cle_groupe].append(ep)

            options_lots = list(groupes_lots.keys())
            index_defaut = min(index_selectionne, len(options_lots) - 1) if options_lots else 0

            choix_lot = st.selectbox("📦 Sélectionner le lot d'éprouvettes à écraser / modifier :", options_lots, index=index_defaut, key="select_lot_saisie")
            lot_selected = groupes_lots[choix_lot]
            sample = lot_selected[0]
            betonnage_id = sample.get("betonnage_id")

            info_betonnage = cache_betonnage.get(betonnage_id) or obtenir_infos_betonnage_parent(supabase, betonnage_id)
            historique_complet = obtenir_historique_betonnage(supabase, betonnage_id)
            exact_bl_phase1 = extraire_num_bl(sample, info_betonnage or {}, choix_lot)
            num_reception_affiche = sample.get("num_reception") or sample.get("n_reception") or ((info_betonnage or {}).get("num_reception") or "-")

            col_l1, col_l2, col_l3, col_l4 = st.columns(4)
            col_l1.metric("N° Réception", str(num_reception_affiche))
            col_l2.metric("N° Bon Livraison", exact_bl_phase1)
            col_l3.metric("Ouvrage", str(((info_betonnage or {}).get("ouvrage")) or sample.get("ouvrage") or "-"))
            col_l4.metric("Échéance Visée", str(sample.get("echeance", "-")))

            st.markdown("---")
            col_g1, col_g2 = st.columns(2)
            tech_global = col_g1.text_input("Technicien / Opérateur", value=sample.get("technicien", (info_betonnage or {}).get("technicien_prelevement") or "Technicien LPEE"), key="tech_global")
            obs_globale = col_g2.text_input("Commentaire / Observation", value=sample.get("observations", "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"), key="obs_global")

            st.markdown("##### 📝 Saisie / Modification des forces d'écrasement")
            ref_controle_courante = determiner_ref_controle(supabase, betonnage_id, info_betonnage, sample)
            lot_key = f"df_lot_{choix_lot}"

            if lot_key not in st.session_state or mode_admin:
                rows_list = []
                for ep in lot_selected:
                    sec = float(ep.get("section") or 176.71)
                    try:
                        f_kn = float(ep.get("force_kn") or 0.0)
                    except (ValueError, TypeError):
                        f_kn = 0.0
                    fc = round((f_kn * 10.0) / sec, 1) if sec > 0 and f_kn > 0 else 0.0
                    rows_list.append({
                        "ID": ep["id"],
                        "🏷️ Référence de Contrôle": str(ep.get("ref_controle") or ref_controle_courante).strip(),
                        "Repère": ep.get("repere_eprouvette", f"/{ep['id']}"),
                        "Forme d'éprouvette": str(ep.get("forme") or "Cylindrique 150x300"),
                        "_section": sec,
                        "Force (kN)": f_kn,
                        "Résistance Fc (MPa)": fc,
                        "Moyenne Resistance Fc (MPa)": 0.0,
                    })
                df_init = pd.DataFrame(rows_list)
                valides_init = df_init[df_init["Résistance Fc (MPa)"] > 0]
                df_init["Moyenne Resistance Fc (MPa)"] = round(valides_init["Résistance Fc (MPa)"].mean(), 1) if not valides_init.empty else 0.0
                st.session_state[lot_key] = df_init

            def update_fc():
                editor_state = st.session_state.get("data_editor_ecrasement", {})
                for row_idx, updated_cols in editor_state.get("edited_rows", {}).items():
                    if "Force (kN)" in updated_cols:
                        try:
                            new_force = float(updated_cols["Force (kN)"])
                        except (ValueError, TypeError):
                            new_force = 0.0
                        sec = float(st.session_state[lot_key].at[row_idx, "_section"])
                        st.session_state[lot_key].at[row_idx, "Force (kN)"] = new_force
                        st.session_state[lot_key].at[row_idx, "Résistance Fc (MPa)"] = round((new_force * 10.0) / sec, 1) if sec > 0 and new_force > 0 else 0.0

                    if "🏷️ Référence de Contrôle" in updated_cols:
                        nouvelle_ref = str(updated_cols["🏷️ Référence de Contrôle"] or "").strip()
                        st.session_state[lot_key].at[row_idx, "🏷️ Référence de Contrôle"] = nouvelle_ref
                        st.session_state[f"ref_controle_beton_{betonnage_id}"] = nouvelle_ref

                df_cur = st.session_state[lot_key]
                valides = df_cur[df_cur["Résistance Fc (MPa)"].astype(float) > 0]
                st.session_state[lot_key]["Moyenne Resistance Fc (MPa)"] = round(valides["Résistance Fc (MPa)"].astype(float).mean(), 1) if not valides.empty else 0.0

            st.data_editor(
                st.session_state[lot_key],
                column_config={
                    "ID": st.column_config.NumberColumn("ID", disabled=True),
                    "Repère": st.column_config.TextColumn("Repère", disabled=not mode_admin),
                    "Forme d'éprouvette": st.column_config.TextColumn("Forme d'éprouvette", disabled=True),
                    "_section": None,
                    "Force (kN)": st.column_config.NumberColumn("⚡ Force (kN)", min_value=0.0, max_value=3000.0, step=0.1, format="%.1f"),
                    "Résistance Fc (MPa)": st.column_config.NumberColumn("💥 Résistance Fc (MPa)", disabled=True, format="%.1f"),
                    "Moyenne Resistance Fc (MPa)": st.column_config.NumberColumn("📊 Moyenne Resistance Fc (MPa)", disabled=True, format="%.1f"),
                },
                use_container_width=True,
                hide_index=True,
                key="data_editor_ecrasement",
                on_change=update_fc,
            )

            df_actuel = st.session_state[lot_key]
            dict_actuel = {int(row["ID"]): row for _, row in df_actuel.iterrows()}
            export_data = []

            for ep_h in (historique_complet or lot_selected):
                ep_id, sec_h = ep_h["id"], float(ep_h.get("section") or 176.71)
                if ep_id in dict_actuel:
                    r_s = dict_actuel[ep_id]
                    f_kn, fc_mpa = float(r_s["Force (kN)"]), float(r_s["Résistance Fc (MPa)"])
                    ref_p, rep_s = str(r_s["🏷️ Référence de Contrôle"]).strip(), str(r_s["Repère"]).strip()
                else:
                    try:
                        f_kn = float(ep_h.get("force_kn") or 0.0)
                    except (ValueError, TypeError):
                        f_kn = 0.0
                    fc_mpa = float(ep_h.get("fc_mpa") or (round((f_kn * 10.0) / sec_h, 1) if f_kn > 0 else 0.0))
                    ref_p, rep_s = str(ep_h.get("ref_controle") or ref_controle_courante).strip(), str(ep_h.get("repere_eprouvette", f"/{ep_id}")).strip()

                export_data.append({
                    "repere_eprouvette": f"{ref_p}{rep_s}" if ref_p else rep_s,
                    "forme": ep_h.get("forme", "Cylindrique 150x300"),
                    "section": sec_h,
                    "force_kn": f_kn,
                    "fc_mpa": fc_mpa,
                    "date_essai": ep_h.get("date_ecrasement", "-"),
                    "age": str(ep_h.get("echeance", "28")).replace(" jours", "").replace("j", ""),
                    "statut": "En cours" if f_kn == 0 else "Réalisé",
                })

            infos_header = {
                "re_num": "25/260/LGV/ B/",
                "dossier": "2025-260-05985-2025-0247",
                "client": "TGCC",
                "num_bl": exact_bl_phase1,
                "ouvrage": (info_betonnage or {}).get("ouvrage") or sample.get("ouvrage"),
                "lieu_prelevement": (info_betonnage or {}).get("ouvrage") or sample.get("ouvrage"),
                "classe_beton": sample.get("classe_beton", "C35/45"),
                "date_coulee": extraire_date_coulee(info_betonnage or sample),
                "affaissement": (info_betonnage or {}).get("affaissement"),
                "temperature": (info_betonnage or {}).get("temperature"),
                "forme": sample.get("forme", "Cylindrique 150x300"),
                "centrale": (info_betonnage or {}).get("centrale") or sample.get("centrale"),
                "observations": obs_globale,
                "technicien_prelevement": (info_betonnage or {}).get("technicien_prelevement") or tech_global,
            }

            st.markdown("---")
            col_b1, col_b2 = st.columns(2)
            btn_enregistrer = col_b1.button("💾 Valider et Mettre à Jour Le Lot" if mode_admin else "💾 Valider et Enregistrer Le Lot", type="primary", use_container_width=True)
            col_b2.download_button(
                "📄 Télécharger le PV (Excel Modèle LPEE)",
                generer_pv_excel(export_data, infos_header),
                file_name=f"PV_Ecrasement_LPEE_{exact_bl_phase1 if exact_bl_phase1 != '-' else 'BL'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            if btn_enregistrer:
                if (df_actuel["Force (kN)"].astype(float) == 0).any() and not mode_admin:
                    st.error("❌ Les forces de rupture doivent toutes être saisies (> 0 kN).")
                else:
                    succes_lot = 0
                    ref_finale = df_actuel.iloc[0].get("🏷️ Référence de Contrôle")
                    try:
                        supabase.table("suivi_betonnage").update({"ref_controle": ref_finale}).eq("id", betonnage_id).execute()
                    except Exception:
                        pass

                    for _, row in df_actuel.iterrows():
                        upd = {
                            "ref_controle": row.get("🏷️ Référence de Contrôle"),
                            "repere_eprouvette": row.get("Repère"),
                            "force_kn": float(row["Force (kN)"]),
                            "fc_mpa": float(row["Résistance Fc (MPa)"]),
                            "technicien": tech_global,
                            "observations": obs_globale,
                        }
                        try:
                            supabase.table("suivi_controle_beton").update(upd).eq("id", int(row["ID"])).execute()
                            succes_lot += 1
                        except Exception as e:
                            st.error(f"Erreur sur l'éprouvette {row['Repère']} : {e}")

                    if succes_lot == len(df_actuel):
                        st.balloons()
                        st.success(f"✅ Lot de {succes_lot} éprouvettes mis à jour / validé !")


# =========================================================
# 4. INITIALISATION DE L'APPLICATION ET DU POINT D'ENTRÉE
# =========================================================
if __name__ == "__main__":
    query_params = st.query_params
    url_rec = query_params.get("rec") or query_params.get("num_reception")
    url_bid = query_params.get("beton_id") or query_params.get("id")

    if url_rec or url_bid:
        if url_rec:
            st.session_state["pending_qr_rec"] = str(url_rec).strip()
        if url_bid:
            st.session_state["pending_qr_bid"] = str(url_bid).strip()

        st.session_state["nav_segmented_phase"] = OPTIONS_ONGLETS[2]
        st.session_state["nav_radio_phase"] = OPTIONS_ONGLETS[2]
        st.session_state["onglet_actif"] = OPTIONS_ONGLETS[2]

    supabase_client = None
    if "supabase" in st.session_state:
        supabase_client = st.session_state["supabase"]
    else:
        try:
            url = st.secrets["SUPABASE_URL"]
            key = st.secrets["SUPABASE_KEY"]
            if Client:
                supabase_client = create_client(url, key)
                st.session_state["supabase"] = supabase_client
        except Exception as e:
            st.error("💡 Connexion Supabase : Assurez-vous d'avoir configuré SUPABASE_URL et SUPABASE_KEY dans vos secrets Streamlit.")

    if not st.session_state.get("user_logged", False):
        if supabase_client:
            afficher_ecran_connexion(supabase_client)
        else:
            st.error("Impossible d'afficher l'écran de connexion sans client Supabase actif.")
    else:
        if supabase_client:
            show(supabase_client)
