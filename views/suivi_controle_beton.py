import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta

def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    # Vérification du rôle administrateur
    is_admin = st.session_state.get("is_admin", False) or st.session_state.get("role") == "admin"

    # ---------------------------------------------------------
    # 1. RÉCUPÉRATION DES BÉTONNAGES AVEC PRÉLÈVEMENT "OUI"
    # ---------------------------------------------------------
    betonnages_preleves = []
    try:
        res_beton = supabase.table("suivi_betonnage").select("*").order("id", desc=True).execute()
        if res_beton.data:
            # Filtrer les enregistrements où un prélèvement a été effectué (Prélèvement commence par OUI)
            betonnages_preleves = [
                item for item in res_beton.data 
                if item.get("prelevement") and str(item.get("prelevement")).upper().startswith("OUI")
            ]
    except Exception as e:
        st.error(f"Erreur lors du chargement des suivis de bétonnage : {e}")

    if not betonnages_preleves:
        st.warning("⚠️ Aucun suivi de bétonnage avec prélèvement (OUI) n'a été trouvé dans la base de données.")
        return

    # ---------------------------------------------------------
    # 2. SELECTION DU BETONNAGE CONCERNÉ
    # ---------------------------------------------------------
    st.subheader("📌 1. Sélection du Prélèvement à Tester")

    # Format de la liste déroulante pour identifier facilement le bétonnage
    options_beton = {
        f"ID #{b['id']} | Date: {b.get('date_livraison', b.get('date_betonnage', 'N/A'))} | BL: {b.get('num_bl', 'N/A')} | Ouvrage: {b.get('ouvrage', 'N/A')} | Classe: {b.get('classe', 'N/A')}": b
        for b in betonnages_preleves
    }

    selected_label = st.selectbox(
        "Choisissez l'enregistrement de bétonnage prélevé :",
        options=list(options_beton.keys()),
        key="select_beton_preleve"
    )

    selected_beton = options_beton[selected_label]

    # Extraction des infos du bétonnage sélectionné
    beton_id = selected_beton["id"]
    date_coulee_str = selected_beton.get("date_livraison") or selected_beton.get("date_betonnage") or str(date.today())
    try:
        date_coulee = datetime.strptime(date_coulee_str, "%Y-%m-%d").date()
    except Exception:
        date_coulee = date.today()

    classe_beton = selected_beton.get("classe", "C25/30")
    num_bl = selected_beton.get("num_bl", "N/A")
    ouvrage = selected_beton.get("ouvrage", "N/A")

    # Affichage récapitulatif
    info_col1, info_col2, info_col3, info_col4 = st.columns(4)
    info_col1.metric("Date de Coulée", str(date_coulee))
    info_col2.metric("N° Bon Livraison", str(num_bl))
    info_col3.metric("Ouvrage / Élément", str(ouvrage))
    info_col4.metric("Classe Spécifiée", str(classe_beton))

    st.markdown("---")

    # ---------------------------------------------------------
    # 3. FORMULAIRE DE SAISIE DE L'ÉCRASEMENT
    # ---------------------------------------------------------
    st.subheader("💥 2. Saisie de l'Écrasement d'Éprouvette")

    # Mode Édition vs Nouvelle saisie
    editing_item = st.session_state.get("edit_controle_beton_item", None)
    if editing_item:
        st.info(f"✏️ **Mode Modification Écrasement** - ID #{editing_item['id']}")

    f_col1, f_col2, f_col3 = st.columns(3)

    with f_col1:
        echeance_options = ["3 jours", "7 jours", "28 jours", "90 jours"]
        echeance = st.selectbox(
            "Échéance d'écrasement", 
            options=echeance_options,
            index=echeance_options.index(editing_item["echeance"]) if editing_item and editing_item.get("echeance") in echeance_options else 2,
            key="ctrl_echeance"
        )

        # Calcul automatique de la date d'écrasement théorique
        jours_dict = {"3 jours": 3, "7 jours": 7, "28 jours": 28, "90 jours": 90}
        nb_jours = jours_dict.get(echeance, 28)
        date_ecrasement_prevue = date_coulee + timedelta(days=nb_jours)

        date_ecrasement = st.date_input(
            "Date d'écrasement effective", 
            value=datetime.strptime(editing_item["date_ecrasement"], "%Y-%m-%d").date() if editing_item and editing_item.get("date_ecrasement") else date_ecrasement_prevue,
            key="ctrl_date_ecrasement"
        )

        st.caption(f"📅 Date théorique à {echeance} : **{date_ecrasement_prevue}**")

    with f_col2:
        repere_eprouvette = st.text_input(
            "Repère Éprouvette (ex: E1, E2...)", 
            value=editing_item.get("repere_eprouvette", "E1") if editing_item else "E1",
            key="ctrl_repere"
        )
        
        forme = st.selectbox(
            "Forme de l'éprouvette", 
            ["Cylindrique (16x32 cm)", "Cubique (15x15 cm)"],
            index=0 if not editing_item or "Cylindrique" in editing_item.get("forme", "") else 1,
            key="ctrl_forme"
        )

        # Section selon la forme (cm²)
        section_standard = 201.06 if "Cylindrique" in forme else 225.0
        section = st.number_input(
            "Section de l'éprouvette (cm²)", 
            value=float(editing_item.get("section", section_standard)) if editing_item else section_standard,
            step=0.1,
            format="%.2f",
            key="ctrl_section"
        )

    with f_col3:
        masse = st.number_input(
            "Masse de l'éprouvette (g)", 
            value=float(editing_item.get("masse", 12500)) if editing_item else 12500.0,
            step=50.0,
            key="ctrl_masse"
        )

        force_kn = st.number_input(
            "Force de rupture (kN)", 
            value=float(editing_item.get("force_kn", 600.0)) if editing_item else 600.0,
            step=5.0,
            format="%.1f",
            key="ctrl_force"
        )

        technicien = st.text_input(
            "Opérateur / Technicien", 
            value=editing_item.get("technicien", "Agent LPEE") if editing_item else "Agent LPEE",
            key="ctrl_tech"
        )

    # Calcul de la résistance Fc (MPa) = Force (kN) * 10 / Section (cm²)
    fc_mpa = round((force_kn * 10.0) / section, 2) if section > 0 else 0.0

    # Résistance caractéristique requise à 28j (extraite de la classe, ex C25/30 -> 25 MPa)
    fck_req = 25.0
    if "/" in classe_beton:
        try:
            fck_req = float(classe_beton.split("/")[0].replace("C", "").strip())
        except Exception:
            fck_req = 25.0

    st.markdown("---")
    res_c1, res_c2, res_c3 = st.columns(3)
    res_c1.metric("Force de Rupture", f"{force_kn:.1f} kN")
    
    # Indicateur visuel selon échéance
    delta_text = f"Cible 28j : {fck_req} MPa"
    res_c2.metric("Résistance Fc (MPa)", f"{fc_mpa:.2f} MPa", delta=delta_text)

    observations = st.text_area(
        "Observations / Mode de rupture", 
        value=editing_item.get("observations", "Rupture satisfaisante (NF EN 12390-3).") if editing_item else "Rupture satisfaisante (NF EN 12390-3).",
        key="ctrl_obs"
    )

    # ---------------------------------------------------------
    # 4. ENREGISTREMENT SOUFFLÉ DANS SUPABASE
    # ---------------------------------------------------------
    btn_c1, btn_c2 = st.columns([3, 1])

    with btn_c1:
        label_btn = "🔄 Mettre à jour le contrôle" if editing_item else "💾 Enregistrer l'écrasement"
        if st.button(label_btn, type="primary", use_container_width=True, key="btn_save_controle"):
            payload = {
                "betonnage_id": beton_id,
                "num_bl": num_bl,
                "ouvrage": ouvrage,
                "classe_beton": classe_beton,
                "date_coulee": str(date_coulee),
                "echeance": echeance,
                "date_ecrasement": str(date_ecrasement),
                "repere_eprouvette": repere_eprouvette,
                "forme": forme,
                "section": float(section),
                "masse": float(masse),
                "force_kn": float(force_kn),
                "fc_mpa": float(fc_mpa),
                "technicien": technicien,
                "observations": observations
            }

            try:
                # Vérification / filtrage des colonnes valides
                sample = supabase.table("suivi_controle_beton").select("*").limit(1).execute()
                if sample.data and len(sample.data) > 0:
                    valid_cols = set(sample.data[0].keys())
                    safe_payload = {k: v for k, v in payload.items() if k in valid_cols}
                else:
                    safe_payload = payload

                if editing_item:
                    supabase.table("suivi_controle_beton").update(safe_payload).eq("id", editing_item["id"]).execute()
                    st.success(f"✅ Écrasement ID #{editing_item['id']} mis à jour !")
                    st.session_state["edit_controle_beton_item"] = None
                else:
                    supabase.table("suivi_controle_beton").insert(safe_payload).execute()
                    st.success("✅ Contrôle d'écrasement enregistré avec succès !")

                st.rerun()
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement dans la table 'suivi_controle_beton' : {e}")

    with btn_c2:
        if editing_item:
            if st.button("❌ Annuler", use_container_width=True):
                st.session_state["edit_controle_beton_item"] = None
                st.rerun()

    # ---------------------------------------------------------
    # 5. HISTORIQUE DES ÉCRASEMENTS DE BÉTON
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📋 Historique du Suivi de Contrôle du Béton")

    try:
        res_ctrl = supabase.table("suivi_controle_beton").select("*").order("id", desc=True).execute()

        if res_ctrl.data and len(res_ctrl.data) > 0:
            clean_rows = []
            for item in res_ctrl.data:
                clean_rows.append({
                    "ID": item.get("id"),
                    "N° BL": item.get("num_bl"),
                    "Ouvrage": item.get("ouvrage"),
                    "Classe": item.get("classe_beton"),
                    "Date Coulée": item.get("date_coulee"),
                    "Échéance": item.get("echeance"),
                    "Date Écrasement": item.get("date_ecrasement"),
                    "Repère": item.get("repere_eprouvette"),
                    "Masse (g)": item.get("masse"),
                    "Force (kN)": item.get("force_kn"),
                    "Fc (MPa)": item.get("fc_mpa"),
                    "Technicien": item.get("technicien")
                })

            df_display = pd.DataFrame(clean_rows)
            st.dataframe(df_display, use_container_width=True, hide_index=True)

            # Actions Administrateur
            if is_admin:
                st.markdown("### ⚙️ Actions Administrateur (Modifier / Supprimer)")
                selected_ctrl_id = st.selectbox(
                    "Sélectionner un essai d'écrasement par son ID :",
                    options=[item["id"] for item in res_ctrl.data],
                    key="admin_select_ctrl_id"
                )

                act_col1, act_col2 = st.columns(2)
                with act_col1:
                    if st.button("✏️ Modifier cet écrasement", use_container_width=True):
                        target = next((i for i in res_ctrl.data if i["id"] == selected_ctrl_id), None)
                        if target:
                            st.session_state["edit_controle_beton_item"] = target
                            st.rerun()

                with act_col2:
                    if st.button("🗑️ Supprimer cet écrasement", type="primary", use_container_width=True):
                        try:
                            supabase.table("suivi_controle_beton").delete().eq("id", selected_ctrl_id).execute()
                            st.success(f"🗑️ Contrôle #{selected_ctrl_id} supprimé.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur de suppression : {e}")
        else:
            st.info("Aucun essai d'écrasement de béton n'a encore été enregistré.")

    except Exception as e:
        st.warning(f"Impossible de charger l'historique des écrasements : {e}")
