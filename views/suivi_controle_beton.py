import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta

def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    # Création des onglets
    tab_prog, tab_saisie, tab_hist = st.tabs([
        "📅 Phase 1 : Programmation", 
        "💥 Phase 2 : Saisie des Écrasements", 
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
        
        # Récupération du nombre d'éprouvettes prélevées dans le suivi de bétonnage (par défaut 2 si non spécifié)
        nb_ep_suivi = int(beton_p.get("nb_eprouvettes") or beton_p.get("nombre_eprouvettes") or 2)

        # Récupération automatique de l'affaissement et de la température
        affaissement_raw = str(beton_p.get("affaissement") or beton_p.get("slump") or "N/A")
        temp_beton_p = str(beton_p.get("temperature") or beton_p.get("temp_beton") or "N/A")

        # Mise en forme de l'affaissement en mm
        affaissement_p = f"{affaissement_raw} mm" if affaissement_raw != "N/A" else "N/A"

        date_coulee_raw = beton_p.get("date_coulee") or beton_p.get("date_livraison") or str(date.today())
        
        try:
            date_coulee_p = datetime.strptime(str(date_coulee_raw), "%Y-%m-%d").date()
        except Exception:
            date_coulee_p = date.today()

        # Génération automatique de la Référence de Contrôle
        # Exemple: REF-5-PRA 500-2026-08-01
        ref_controle_defaut = f"REF-{b_id}-{ouvrage_p}-{date_coulee_p}"

        st.markdown("---")
        
        # Ligne Référence de contrôle
        ref_controle_p = st.text_input(
            "🏷️ Référence de Contrôle (Identifiant unique du prélèvement)", 
            value=ref_controle_defaut, 
            key=f"p_ref_ctrl_{b_id}"
        )

        st.markdown("---")
        
        # Ligne 1 : BL, Ouvrage, Classe de béton
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            st.text_input("N° Bon de Livraison (BL)", value=num_bl_p, disabled=True, key=f"p_bl_{b_id}")
        with col_p2:
            st.text_input("Ouvrage / Élément", value=ouvrage_p, disabled=True, key=f"p_ouv_{b_id}")
        with col_p3:
            st.text_input("Classe de Béton Spécifiée", value=classe_beton_p, disabled=True, key=f"p_classe_{b_id}")

        # Ligne 2 : Affaissement (mm) et Température (°C) remplis automatiquement
        col_p4, col_p5 = st.columns(2)
        with col_p4:
            st.text_input("Affaissement / Slump (mm)", value=affaissement_p, disabled=True, key=f"p_aff_{b_id}")
        with col_p5:
            st.text_input("Température Béton Frais (°C)", value=f"{temp_beton_p} °C" if temp_beton_p != "N/A" else "N/A", disabled=True, key=f"p_temp_{b_id}")

        st.markdown("---")
        
        # Ligne 3 : Programmation des dates & échéances
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
        with col_e4:
            nb_eprouvettes_p = st.number_input(
                "Nombre d'éprouvettes", 
                min_value=1, 
                max_value=12, 
                value=nb_ep_suivi, 
                step=1, 
                key=f"p_nb_ep_{b_id}"
            )

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            forme_p = st.selectbox(
                "Type / Forme d'éprouvette", 
                ["Cylindrique 150x300", "Cylindrique 160x320", "Cylindrique 100x200"], 
                key=f"p_forme_{b_id}"
            )
        
        # Calcul strict de la section théorique (en cm²)
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

        st.markdown("##### 🏷️ Repères codés des éprouvettes")
        reperes_p = []
        cols_rep = st.columns(min(int(nb_eprouvettes_p), 6))
        for i in range(int(nb_eprouvettes_p)):
            # Index d'affichage dynamique pour la grille
            col_idx = i % 6
            with cols_rep[col_idx]:
                # Génération automatique sous le format REF/1, REF/2, ...
                rep_defaut = f"{ref_controle_p}/{i+1}"
                rep_val = st.text_input(
                    f"Repère #{i+1}", 
                    value=rep_defaut, 
                    key=f"prog_rep_{b_id}_{echeance_key_clean}_{i}"
                )
                reperes_p.append(rep_val)

        if st.button("📌 Enregistrer la Programmation", type="primary", use_container_width=True, key=f"btn_save_prog_{b_id}"):
            succes_cnt = 0
            for rep in reperes_p:
                payload_prog = {
                    "betonnage_id": b_id,
                    "reference_controle": ref_controle_p,
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
                st.success(f"✅ {succes_cnt} éprouvette(s) programmée(s) pour le {date_ecrasement_prevue} ({echeance_p}) avec la référence {ref_controle_p} !")
                st.rerun()

    # =========================================================
    # PHASE 2 : SAISIE DES RÉSULTATS D'ÉCRASEMENT
    # =========================================================
    with tab_saisie:
        st.subheader("💥 2. Saisie des Résultats d'Écrasement")

        eprouvettes_en_attente = []
        try:
            res_att = supabase.table("suivi_controle_beton").select("*").order("id", desc=True).execute()
            if res_att.data:
                eprouvettes_en_attente = [
                    e for e in res_att.data 
                    if e.get("force_kn") is None or float(e.get("force_kn", 0)) == 0
                ]
        except Exception as e:
            st.error(f"Erreur de chargement des essais en attente : {e}")

        if not eprouvettes_en_attente:
            st.info("👍 Aucune éprouvette en attente d'écrasement.")
        else:
            options_saisie = {
                f"ID #{e['id']} | Ref: {e.get('reference_controle', 'N/A')} | Repère: {e.get('repere_eprouvette')} | Échéance: {e.get('echeance')} ({e.get('date_ecrasement')})": e
                for e in eprouvettes_en_attente
            }

            selected_saisie_label = st.selectbox("Sélectionner l'éprouvette à écraser :", list(options_saisie.keys()), key="saisie_select")
            item_saisie = options_saisie[selected_saisie_label]
            item_id = item_saisie["id"]

            st.markdown("---")
            col_s1, col_s2, col_s3, col_s4 = st.columns(4)
            col_s1.metric("Ouvrage", str(item_saisie.get("ouvrage")))
            col_s2.metric("Classe Béton", str(item_saisie.get("classe_beton")))
            col_s3.metric("Échéance", str(item_saisie.get("echeance")))
            col_s4.metric("Date Écrasement", str(item_saisie.get("date_ecrasement")))

            st.markdown("---")
            col_in1, col_in2, col_in3 = st.columns(3)

            with col_in1:
                section_val = float(item_saisie.get("section", 176.71))
                # Section mesurée désactivée
                section_reel = st.number_input("Section mesurée (cm²)", value=section_val, format="%.2f", disabled=True, key=f"s_section_{item_id}")
                masse_g = st.number_input("Masse de l'éprouvette (g)", value=12500.0, step=10.0, key=f"s_masse_{item_id}")

            with col_in2:
                force_kn = st.number_input("Force de rupture (kN)", value=500.0, step=5.0, format="%.1f", key=f"s_force_{item_id}")
                tech_saisie = st.text_input("Technicien / Opérateur", value="Technicien LPEE", key=f"s_tech_{item_id}")

            fc_calc = round((force_kn * 10.0) / section_reel, 2) if section_reel > 0 else 0.0

            with col_in3:
                st.markdown("##### Résultat du calcul")
                st.metric("Résistance Fc", f"{fc_calc:.2f} MPa")

            obs_saisie = st.text_area("Observations / Mode de rupture", value="Rupture satisfaisante (NF EN 12390-3).", key=f"s_obs_{item_id}")

            if st.button("💾 Valider et Enregistrer l'Écrasement", type="primary", use_container_width=True, key=f"btn_save_saisie_{item_id}"):
                update_payload = {
                    "section": section_reel,
                    "masse": masse_g,
                    "force_kn": force_kn,
                    "fc_mpa": fc_calc,
                    "technicien": tech_saisie,
                    "observations": obs_saisie
                }

                try:
                    supabase.table("suivi_controle_beton").update(update_payload).eq("id", item_id).execute()
                    st.success(f"✅ Écrasement validé pour {item_saisie.get('repere_eprouvette')} : Fc = {fc_calc} MPa !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur lors de la mise à jour de l'écrasement : {e}")

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
                    "id", "reference_controle", "num_bl", "ouvrage", "classe_beton", "date_coulee", 
                    "echeance", "date_ecrasement", "repere_eprouvette",
                    "masse", "force_kn", "fc_mpa", "technicien"
                ]
                cols_valid = [c for c in cols_display if c in df_all.columns]
                st.dataframe(df_all[cols_valid], use_container_width=True, hide_index=True)
            else:
                st.info("Aucun enregistrement d'écrasement dans la base de données.")
        except Exception as e:
            st.error(f"Erreur lors du chargement de l'historique : {e}")
