import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta

def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    # Création des onglets
    tab_prog, tab_saisie, tab_hist = st.tabs([
        "📅 Phase 1 : Programmation", 
        "💥 Phase 2 : Saisie des Écrasements (Par Lot)", 
        "📋 Historique Complet"
    ])

    # ---------------------------------------------------------
    # RÉCUPÉRATION DES BÉTONNAGES PRÉLEVÉS (OUI)
    # ---------------------------------------------------------
    betonnages_preleves = []
    try:
        res_beton = supabase.table("suivi_betonnage").select("*").order("id", desc=True).execute()
        if res_beton.data:
            betonnages_preleves = [
                item for item in res_beton.data
                if item.get("prelevement") and str(item.get("prelevement")).upper().startswith("OUI")
            ]
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des bétonnages : {e}")

    if not betonnages_preleves:
        st.info("ℹ️ Aucun suivi de bétonnage avec prélèvement d'éprouvettes (OUI) trouvé.")
        return

    options_beton = {
        f"ID #{b['id']} | BL: {b.get('num_bl', 'N/A')} | Ouvrage: {b.get('ouvrage', 'N/A')} | Date: {b.get('date_coulee', b.get('date_livraison', 'N/A'))} | Classe: {b.get('classe_beton', b.get('classe', 'N/A'))}": b
        for b in betonnages_preleves
    }

    # =========================================================
    # PHASE 1 : PROGRAMMATION DES ÉPROUVETTES
    # =========================================================
    with tab_prog:
        st.subheader("📅 1. Programmer les Échéances d'Écrasement")
        
        choix_label_p = st.selectbox("Sélectionner la fiche de bétonnage :", list(options_beton.keys()), key="prog_beton_select")
        beton_p = options_beton[choix_label_p]

        # Récupération dynamique des données du bétonnage
        b_id = beton_p.get("id")
        num_bl_p = str(beton_p.get("num_bl") or "N/A")
        ouvrage_p = str(beton_p.get("ouvrage") or "N/A")
        classe_beton_p = str(beton_p.get("classe_beton") or beton_p.get("classe") or "N/A")
        
        # Total d'éprouvettes prévu au suivi de bétonnage (ex: 12)
        raw_nb_ep = beton_p.get("nb_eprouvettes") or beton_p.get("nombre_eprouvettes")
        try:
            total_eprouvettes_prevues = int(raw_nb_ep) if raw_nb_ep is not None else 12
        except (ValueError, TypeError):
            total_eprouvettes_prevues = 12

        # Vérification des éprouvettes déjà programmées
        eprouvettes_deja_prog = 0
        try:
            res_deja = supabase.table("suivi_controle_beton").select("id").eq("betonnage_id", b_id).execute()
            if res_deja.data:
                eprouvettes_deja_prog = len(res_deja.data)
        except Exception:
            eprouvettes_deja_prog = 0

        # Calcul du solde disponible
        solde_disponible = max(0, total_eprouvettes_prevues - eprouvettes_deja_prog)

        affaissement_raw = str(beton_p.get("affaissement") or beton_p.get("slump") or "N/A")
        temp_beton_p = str(beton_p.get("temperature") or beton_p.get("temp_beton") or "N/A")
        affaissement_p = f"{affaissement_raw} mm" if affaissement_raw != "N/A" else "N/A"

        date_coulee_raw = beton_p.get("date_coulee") or beton_p.get("date_livraison") or str(date.today())
        try:
            date_coulee_p = datetime.strptime(str(date_coulee_raw), "%Y-%m-%d").date()
        except Exception:
            date_coulee_p = date.today()

        ref_controle_defaut = f"REF-{b_id}-{ouvrage_p}"

        st.markdown("---")
        st.info(
            f"📊 **Quota Éprouvettes :** Total prévu : **{total_eprouvettes_prevues}** | "
            f"Déjà programmée(s) : **{eprouvettes_deja_prog}** | "
            f"Reste disponible : **{solde_disponible}**"
        )

        ref_controle_p = st.text_input(
            "🏷️ Référence de Contrôle (Préfixe du repère)", 
            value=ref_controle_defaut, 
            key=f"p_ref_ctrl_{b_id}"
        )

        st.markdown("---")
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            st.text_input("N° Bon de Livraison (BL)", value=num_bl_p, disabled=True, key=f"p_bl_{b_id}")
        with col_p2:
            st.text_input("Ouvrage / Élément", value=ouvrage_p, disabled=True, key=f"p_ouv_{b_id}")
        with col_p3:
            st.text_input("Classe de Béton Spécifiée", value=classe_beton_p, disabled=True, key=f"p_classe_{b_id}")

        col_p4, col_p5 = st.columns(2)
        with col_p4:
            st.text_input("Affaissement / Slump (mm)", value=affaissement_p, disabled=True, key=f"p_aff_{b_id}")
        with col_p5:
            st.text_input("Température Béton Frais (°C)", value=f"{temp_beton_p} °C" if temp_beton_p != "N/A" else "N/A", disabled=True, key=f"p_temp_{b_id}")

        st.markdown("---")
        col_e1, col_e2, col_e3, col_e4 = st.columns(4)
        with col_e1:
            echeance_p = st.selectbox("Âge / Échéance visée", ["3 jours", "7 jours", "28 jours", "90 jours"], index=2, key=f"p_echeance_{b_id}")
        
        jours_dict = {"3 jours": 3, "7 jours": 7, "28 jours": 28, "90 jours": 90}
        nb_j = jours_dict.get(echeance_p, 28)
        date_prevue_auto = date_coulee_p + timedelta(days=nb_j)
        echeance_key_clean = echeance_p.replace(' ', '_')

        with col_e2:
            st.date_input("Date de Coulée", value=date_coulee_p, disabled=True, key=f"p_date_coul_{b_id}")
        with col_e3:
            date_ecrasement_prevue = st.date_input(
                "Date d'Écrasement Prévue", 
                value=date_prevue_auto, 
                key=f"p_date_ecras_{b_id}_{echeance_key_clean}"
            )

        min_val = 1 if solde_disponible > 0 else 0
        val_defaut = min(2, solde_disponible) if solde_disponible > 0 else 0

        with col_e4:
            if solde_disponible == 0:
                st.warning("⚠️ Quota atteint (12/12).")
                nb_eprouvettes_p = 0
            else:
                nb_eprouvettes_p = st.number_input(
                    "Nombre d'éprouvettes à programmer", 
                    min_value=min_val, 
                    max_value=solde_disponible, 
                    value=val_defaut, 
                    step=1, 
                    key=f"p_nb_ep_{b_id}_{echeance_key_clean}"
                )

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            forme_p = st.selectbox(
                "Type / Forme d'éprouvette", 
                ["Cylindrique 150x300", "Cylindrique 160x320", "Cylindrique 100x200"], 
                key=f"p_forme_{b_id}"
            )
        
        if "150x300" in forme_p:
            sect_def = 176.71
        elif "160x320" in forme_p:
            sect_def = 201.06
        elif "100x200" in forme_p:
            sect_def = 78.54
        else:
            sect_def = 176.71

        with col_f2:
            forme_key_clean = forme_p.replace(' ', '_').replace('x', '_')
            st.number_input(
                "Section Théorique (cm²)", 
                value=sect_def, 
                format="%.2f", 
                disabled=True, 
                key=f"p_section_{b_id}_{forme_key_clean}"
            )

        if int(nb_eprouvettes_p) > 0:
            st.markdown("##### 🏷️ Repères codés des éprouvettes")
            reperes_p = []
            cols_rep = st.columns(min(int(nb_eprouvettes_p), 6))
            for i in range(int(nb_eprouvettes_p)):
                col_idx = i % 6
                with cols_rep[col_idx]:
                    num_ep = eprouvettes_deja_prog + i + 1
                    rep_defaut = f"{ref_controle_p}/{num_ep}"
                    rep_val = st.text_input(
                        f"Repère #{num_ep}", 
                        value=rep_defaut, 
                        key=f"prog_rep_{b_id}_{echeance_key_clean}_{i}"
                    )
                    reperes_p.append(rep_val)

            if st.button("📌 Enregistrer la Programmation", type="primary", use_container_width=True, key=f"btn_save_prog_{b_id}"):
                succes_cnt = 0
                for rep in reperes_p:
                    payload_prog = {
                        "betonnage_id": b_id,
                        "num_bl": num_bl_p,
                        "ouvrage": ouvrage_p,
                        "classe_beton": classe_beton_p,
                        "date_coulee": str(date_coulee_p),
                        "echeance": echeance_p,
                        "date_ecrasement": str(date_ecrasement_prevue),
                        "repere_eprouvette": rep,
                        "forme": forme_p,
                        "section": float(sect_def)
                    }
                    try:
                        res = supabase.table("suivi_controle_beton").insert(payload_prog).execute()
                        if res.data:
                            succes_cnt += 1
                    except Exception as err:
                        st.error(f"Erreur lors de la programmation de {rep} : {err}")

                if succes_cnt > 0:
                    st.success(f"✅ {succes_cnt} éprouvette(s) programmée(s) pour le {date_ecrasement_prevue} ({echeance_p}) !")
                    st.rerun()

    # =========================================================
    # PHASE 2 : SAISIE GROUPÉE PAR ÉCHÉANCE / AGE
    # =========================================================
    with tab_saisie:
        st.subheader("💥 2. Saisie Groupée des Résultats d'Écrasement")

        # Récupérer toutes les éprouvettes non encore écrasées
        eprouvettes_en_attente = []
        try:
            res_att = supabase.table("suivi_controle_beton").select("*").order("id", desc=False).execute()
            if res_att.data:
                eprouvettes_en_attente = [
                    e for e in res_att.data 
                    if e.get("force_kn") is None or float(e.get("force_kn") or 0) == 0
                ]
        except Exception as e:
            st.error(f"Erreur de chargement des essais en attente : {e}")

        if not eprouvettes_en_attente:
            st.info("👍 Aucune éprouvette en attente d'écrasement.")
        else:
            # Regroupement par (betonnage_id + echeance)
            groupes_lots = {}
            for ep in eprouvettes_en_attente:
                b_id_ep = ep.get("betonnage_id")
                ech_ep = ep.get("echeance", "28 jours")
                ouv_ep = ep.get("ouvrage", "N/A")
                dt_ecras = ep.get("date_ecrasement", "N/A")
                
                cle_groupe = f"Ouvrage: {ouv_ep} | Échéance: {ech_ep} (Date: {dt_ecras}) | Bétonnage ID #{b_id_ep}"
                
                if cle_groupe not in groupes_lots:
                    groupes_lots[cle_groupe] = []
                groupes_lots[cle_groupe].append(ep)

            # Sélection du lot à écraser
            choix_lot = st.selectbox("📦 Sélectionner le lot d'éprouvettes à écraser :", list(groupes_lots.keys()), key="select_lot_saisie")
            lot_selected = groupes_lots[choix_lot]

            # Informations générales du lot
            sample = lot_selected[0]
            col_l1, col_l2, col_l3, col_l4 = st.columns(4)
            col_l1.metric("Ouvrage", str(sample.get("ouvrage")))
            col_l2.metric("Classe Béton", str(sample.get("classe_beton")))
            col_l3.metric("Échéance Visée", str(sample.get("echeance")))
            col_l4.metric("Nombre dans le lot", f"{len(lot_selected)} éprouvettes")

            st.markdown("---")
            
            # Formulaire global pour Technicien et Observations
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                tech_global = st.text_input("Technicien / Opérateur", value="Technicien LPEE", key="tech_global")
            with col_g2:
                obs_globale = st.text_input("Observations générales", value="Rupture satisfaisante (NF EN 12390-3).", key="obs_global")

            st.markdown("##### 📝 Saisie des mesures pour le lot")

            # Construction du DataFrame
            rows_list = []
            for ep in lot_selected:
                sec = float(ep.get("section") or 176.71)
                f_kn = float(ep.get("force_kn") or 0.0)
                fc = round((f_kn * 10.0) / sec, 2) if sec > 0 and f_kn > 0 else 0.0
                
                rows_list.append({
                    "ID": ep["id"],
                    "Repère": ep.get("repere_eprouvette", f"EP-{ep['id']}"),
                    "Forme d'éprouvette": str(ep.get("forme") or "Cylindrique 150x300"),
                    "_section": sec,  # Champ masqué pour le calcul
                    "Force (kN)": f_kn,
                    "Résistance Fc (MPa)": fc
                })

            df_saisie = pd.DataFrame(rows_list)

            # Éditeur de données avec exclusion explicite de la section
            edited_df = st.data_editor(
                df_saisie,
                column_config={
                    "ID": st.column_config.NumberColumn("ID", disabled=True),
                    "Repère": st.column_config.TextColumn("Repère", disabled=True),
                    "Forme d'éprouvette": st.column_config.TextColumn("Forme d'éprouvette", disabled=True),
                    "_section": None,  # Masqué de l'interface
                    "Force (kN)": st.column_config.NumberColumn(
                        "⚡ Force (kN)", 
                        help="Saisissez la force de rupture lue sur la presse", 
                        min_value=0.0, 
                        max_value=3000.0, 
                        step=0.1, 
                        format="%.1f"
                    ),
                    "Résistance Fc (MPa)": st.column_config.NumberColumn(
                        "💥 Résistance Fc (MPa)", 
                        disabled=True, 
                        format="%.1f"
                    ),
                },
                use_container_width=True,
                hide_index=True,
                key="data_editor_ecrasement"
            )

            # Recalcul dynamique de Fc
            edited_df["Résistance Fc (MPa)"] = edited_df.apply(
                lambda row: round((row["Force (kN)"] * 10.0) / row["_section"], 2) if row["_section"] > 0 and row["Force (kN)"] > 0 else 0.0,
                axis=1
            )

            # Résumé rapide des résultats
            forces_valides = edited_df[edited_df["Force (kN)"] > 0]
            if not forces_valides.empty:
                fc_moy = round(forces_valides["Résistance Fc (MPa)"].mean(), 2)
                st.success(f"📈 **Résistance moyenne calculée pour les éprouvettes saisies : {fc_moy:.1f} MPa**")
            else:
                st.warning("👈 Veuillez remplir la colonne **Force (kN)** pour chaque éprouvette.")

            # Bouton d'enregistrement pour TOUT LE LOT
            if st.button("💾 Valider et Enregistrer Tout le Lot", type="primary", use_container_width=True):
                if (edited_df["Force (kN)"] == 0).any():
                    st.error("❌ Attention : Une ou plusieurs éprouvettes ont encore une force de 0.0 kN. Veuillez compléter les saisies.")
                else:
                    succes_lot = 0
                    for _, row in edited_df.iterrows():
                        update_payload = {
                            "force_kn": float(row["Force (kN)"]),
                            "fc_mpa": float(row["Résistance Fc (MPa)"]),
                            "technicien": tech_global,
                            "observations": obs_globale
                        }
                        try:
                            supabase.table("suivi_controle_beton").update(update_payload).eq("id", int(row["ID"])).execute()
                            succes_lot += 1
                        except Exception as e:
                            st.error(f"Erreur sur l'éprouvette {row['Repère']} : {e}")

                    if succes_lot == len(edited_df):
                        st.success(f"✅ Lot de {succes_lot} éprouvettes enregistré avec succès !")
                        st.rerun()

    # =========================================================
    # HISTORIQUE ET SUIVI GLOBAL
    # =========================================================
    with tab_hist:
        st.subheader("📋 Historique Général des Contrôles de Béton")

        try:
            res_all = supabase.table("suivi_controle_beton").select("*").order("id", desc=True).execute()
            if res_all.data:
                df_all = pd.DataFrame(res_all.data)
                cols_display = [
                    "id", "num_bl", "ouvrage", "classe_beton", "date_coulee", 
                    "echeance", "date_ecrasement", "repere_eprouvette", "forme",
                    "force_kn", "fc_mpa", "technicien"
                ]
                cols_valid = [c for c in cols_display if c in df_all.columns]
                st.dataframe(df_all[cols_valid], use_container_width=True, hide_index=True)
            else:
                st.info("Aucun enregistrement d'écrasement dans la base de données.")
        except Exception as e:
            st.error(f"Erreur lors du chargement de l'historique : {e}")
