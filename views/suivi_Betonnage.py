import streamlit as st
import pandas as pd
from datetime import datetime, date

def show(supabase):
    st.title("🏗️ Suivi et Contrôle Qualité Béton")
    
    # Récupération des informations de session
    user_info = st.session_state.get("user", {})
    username = user_info.get("username", "")
    role = st.session_state.get("role", "")
    can_edit = st.session_state.get("can_edit", False) or (user_info.get("can_edit", False))
    is_admin = (role == "admin")
    
    # ---------------------------------------------------------
    # 1. FORMULAIRE DE SAISIE
    # ---------------------------------------------------------
    st.subheader("Saisie d'un contrôle")
    
    date_livraison = st.date_input("Date de livraison", value=date.today(), key="saisie_date")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        technicien = st.text_input("Nom du Technicien LPEE", value="Agent LPEE", key="saisie_tech")
        bl_num = st.text_input("N° BL", value="BL-2026-001", key="saisie_bl")
        ouvrage = st.text_input("Ouvrage", value="Voile / Semelle", key="saisie_ouvrage")
        quantite_m3 = st.number_input("Quantité (m³)", min_value=0.0, value=8.0, step=0.5, key="saisie_qte")
        
    with col2:
        client = st.text_input("Client", value="TGCC", disabled=True, key="saisie_client")
        
        # Saisie des heures
        heure_fin = st.time_input("Heure de fin de production", value=datetime.strptime("08:00", "%H:%M").time(), key="saisie_h_fin")
        heure_arrivee = st.time_input("Heure d'arrivée au chantier", value=datetime.strptime("08:35", "%H:%M").time(), key="saisie_h_arr")
        
        # Calcul sécurisé de la durée en minutes
        dt_fin = datetime.combine(date.today(), heure_fin)
        dt_arr = datetime.combine(date.today(), heure_arrivee)
        duree_minutes = int((dt_arr - dt_fin).total_seconds() / 60)
        if duree_minutes < 0:
            duree_minutes += 1440  # Passage de minuit si besoin
        
        st.text_input("Durée de transport / attente (min)", value=f"{duree_minutes} min", disabled=True, key="saisie_duree")
        
        classe_beton = st.selectbox(
            "Classe", 
            ["C25/30", "C30/37", "C35/45", "C40/50", "C45/55"],
            key="saisie_classe"
        )
        
    with col3:
        centrale = st.text_input("Centrale à Béton", value="TG PREFA", key="saisie_centrale")
        meteo = st.selectbox("Météo", ["Ensoleillé ☀️", "Nuageux ☁️", "Pluie 🌧️"], key="saisie_meteo")
        
        temp_beton = st.number_input("Température du Béton (°C)", value=20.0, step=0.1, format="%.1f", key="saisie_t_beton")
        temp_ambiante = st.number_input("Température Ambiante (°C)", value=25.0, step=0.1, format="%.1f", key="saisie_t_amb")
        affaissement = st.number_input("Affaissement (mm)", min_value=0, value=150, step=10, key="saisie_aff")
        
        prelevement = st.selectbox(
            "Prélèvement", 
            ["OUI - Conforme (NF EN 12350-2)", "NON"],
            key="saisie_prel"
        )
        
        is_non_prelevement = "NON" in prelevement
        
        nb_eprouvettes = st.number_input(
            "Nb d'éprouvettes", 
            min_value=0, 
            value=0 if is_non_prelevement else 6,
            disabled=is_non_prelevement,
            key="saisie_eprov"
        )

    observations = st.text_area("Observations", value="Béton conforme", key="saisie_obs")

    # Bouton Enregistrer
    if st.button("💾 Enregistrer", key="btn_enregistrer"):
        data = {
            "date_livraison": str(date_livraison),
            "bl_num": bl_num,
            "ouvrage": ouvrage,
            "quantite_m3": float(quantite_m3),
            "client": client,
            "classe_beton": classe_beton,
            "centrale_beton": centrale,
            "meteo": meteo,
            "heure_fin_coulage": heure_fin.strftime("%H:%M"),
            "heure_arrivee": heure_arrivee.strftime("%H:%M"),
            "temperature": float(temp_beton),
            "temperature_ambiante": float(temp_ambiante),
            "affaissement": int(affaissement),
            "prelevement": prelevement,
            "nb_eprouvettes": int(nb_eprouvettes),
            "observations": observations,
            "technicien": technicien
        }
        
        try:
            supabase.table("suivi_betonnage").insert(data).execute()
            st.success("Enregistrement réussi !")
            st.rerun()
        except Exception as e:
            st.error(f"Erreur d'enregistrement : {e}")

    # ---------------------------------------------------------
    # 2. AFFICHAGE DE L'HISTORIQUE ET ESPACE DE MODIFICATION
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📊 Historique")
    
    try:
        res = supabase.table("suivi_betonnage").select("*").order("id", desc=True).execute()
        if res.data:
            df = pd.DataFrame(res.data)
            
            # 1. Calcul de la colonne "Durée de transport"
            if "heure_fin_coulage" in df.columns and "heure_arrivee" in df.columns:
                def calculer_duree(row):
                    try:
                        str_fin = str(row["heure_fin_coulage"]).split(".")[0]
                        str_arr = str(row["heure_arrivee"]).split(".")[0]
                        
                        fmt = "%H:%M:%S" if len(str_fin.split(":")) == 3 else "%H:%M"
                        h_fin = datetime.strptime(str_fin, fmt)
                        
                        fmt_arr = "%H:%M:%S" if len(str_arr.split(":")) == 3 else "%H:%M"
                        h_arr = datetime.strptime(str_arr, fmt_arr)
                        
                        diff = int((h_arr - h_fin).total_seconds() / 60)
                        if diff < 0:
                            diff += 1440
                        return f"{diff} min"
                    except:
                        return "-"
                
                df["Durée de transport"] = df.apply(calculer_duree, axis=1)

            # 2. Suppression stricte des colonnes secondaires (EXCLUSION de 'id')
            cols_to_drop = [
                col for col in ["created_at", "created", "heure_fin_coulage", "heure_fin", "client", "centrale_beton"] 
                if col in df.columns
            ]
            if cols_to_drop:
                df = df.drop(columns=cols_to_drop)

            # 3. Renommage explicite des colonnes avant réordonnancement
            df = df.rename(columns={
                "id": "ID",
                "date_livraison": "Date Livraison",
                "heure_arrivee": "Heure d'arrivée",
                "bl_num": "N° BL",
                "ouvrage": "Ouvrage",
                "quantite_m3": "Quantité (m³)",
                "classe_beton": "Classe",
                "temperature": "Temp. Béton",
                "temperature_ambiante": "Temp. Ambiante",
                "affaissement": "Affaissement",
                "prelevement": "Prélèvement",
                "nb_eprouvettes": "Nb Éprouvettes",
                "observations": "Observations",
                "technicien": "Technicien",
                "meteo": "Météo"
            })

            # 4. Forcer la colonne 'ID' en toute première position
            all_cols = list(df.columns)
            if "ID" in all_cols:
                all_cols.remove("ID")
                final_cols = ["ID"] + all_cols
                df = df[final_cols]

            # 5. Affichage du DataFrame en masquant l'index Pandas résiduel pour placer ID en 1re position
            st.dataframe(df, use_container_width=True, hide_index=True)

            # --- BLOC MODIFICATION (AUTORISÉ POUR AMINA ET ADMIN) ---
            if is_admin or can_edit:
                st.markdown("---")
                st.subheader("🛠️ Espace Administration - Suivi Béton")
                
                record_options = {f"ID {r['id']} - BL: {r.get('bl_num', 'N/A')} - Ouvrage: {r.get('ouvrage', '')}": r for r in res.data}
                selected_key = st.selectbox("Sélectionner l'enregistrement à gérer", list(record_options.keys()), key="admin_select_record")
                selected_item = record_options[selected_key]
                
                # Disposition des colonnes selon le rôle
                if is_admin:
                    col_ed, col_del = st.columns([2, 1])
                else:
                    col_ed = st.container()

                # --- BLOC DE MODIFICATION ---
                with col_ed:
                    with st.expander("📝 Modifier ce contrôle (Tous les champs)"):
                        with st.form("edit_form_beton_complet"):
                            try:
                                def_date = datetime.strptime(str(selected_item.get("date_livraison", date.today())), "%Y-%m-%d").date()
                            except:
                                def_date = date.today()

                            def parse_heure_safe(val_str, default_str):
                                try:
                                    clean_str = str(val_str).split(".")[0]
                                    fmt = "%H:%M:%S" if len(clean_str.split(":")) == 3 else "%H:%M"
                                    return datetime.strptime(clean_str, fmt).time()
                                except:
                                    return datetime.strptime(default_str, "%H:%M").time()

                            def_h_fin = parse_heure_safe(selected_item.get("heure_fin_coulage"), "08:00")
                            def_h_arr = parse_heure_safe(selected_item.get("heure_arrivee"), "08:35")

                            new_date_livraison = st.date_input("Date de livraison", value=def_date, key="edit_date")
                            new_technicien = st.text_input("Nom du Technicien LPEE", value=selected_item.get("technicien", "Agent LPEE"), key="edit_tech")
                            new_bl = st.text_input("N° BL", value=selected_item.get("bl_num", ""), key="edit_bl")
                            new_ouvrage = st.text_input("Ouvrage", value=selected_item.get("ouvrage", ""), key="edit_ouvrage")
                            new_quantite = st.number_input("Quantité (m³)", value=float(selected_item.get("quantite_m3", 0.0)), key="edit_qte")
                            
                            new_heure_fin = st.time_input("Heure de fin de production", value=def_h_fin, key="edit_h_fin")
                            new_heure_arrivee = st.time_input("Heure d'arrivée au chantier", value=def_h_arr, key="edit_h_arr")
                            
                            classes_list = ["C25/30", "C30/37", "C35/45", "C40/50", "C45/55"]
                            current_classe = selected_item.get("classe_beton", "C25/30")
                            idx_classe = classes_list.index(current_classe) if current_classe in classes_list else 0
                            new_classe = st.selectbox("Classe", classes_list, index=idx_classe, key="edit_classe")
                            
                            new_centrale = st.text_input("Centrale à Béton", value=selected_item.get("centrale_beton", "TG PREFA"), key="edit_centrale")
                            
                            meteo_list = ["Ensoleillé ☀️", "Nuageux ☁️", "Pluie 🌧️"]
                            current_meteo = selected_item.get("meteo", "Ensoleillé ☀️")
                            idx_meteo = meteo_list.index(current_meteo) if current_meteo in meteo_list else 0
                            new_meteo = st.selectbox("Météo", meteo_list, index=idx_meteo, key="edit_meteo")
                            
                            new_temp_beton = st.number_input("Température du Béton (°C)", value=float(selected_item.get("temperature", 20.0)), step=0.1, format="%.1f", key="edit_t_beton")
                            new_temp_amb = st.number_input("Température Ambiante (°C)", value=float(selected_item.get("temperature_ambiante", 25.0)), step=0.1, format="%.1f", key="edit_t_amb")
                            new_affaissement = st.number_input("Affaissement (mm)", value=int(selected_item.get("affaissement", 150)), step=10, key="edit_aff")
                            
                            prelevement_list = ["OUI - Conforme (NF EN 12350-2)", "NON"]
                            current_prel = selected_item.get("prelevement", "NON")
                            idx_prel = prelevement_list.index(current_prel) if current_prel in prelevement_list else 1
                            new_prelevement = st.selectbox("Prélèvement", prelevement_list, index=idx_prel, key="edit_prel")
                            
                            new_nb_eprouvettes = st.number_input("Nb d'éprouvettes", value=int(selected_item.get("nb_eprouvettes", 0)), min_value=0, key="edit_eprov")
                            new_observations = st.text_area("Observations", value=selected_item.get("observations", ""), key="edit_obs")
                            
                            if st.form_submit_button("💾 Enregistrer toutes les modifications"):
                                try:
                                    supabase.table("suivi_betonnage").update({
                                        "date_livraison": str(new_date_livraison),
                                        "technicien": new_technicien,
                                        "bl_num": new_bl,
                                        "ouvrage": new_ouvrage,
                                        "quantite_m3": float(new_quantite),
                                        "heure_fin_coulage": new_heure_fin.strftime("%H:%M"),
                                        "heure_arrivee": new_heure_arrivee.strftime("%H:%M"),
                                        "classe_beton": new_classe,
                                        "centrale_beton": new_centrale,
                                        "meteo": new_meteo,
                                        "temperature": float(new_temp_beton),
                                        "temperature_ambiante": float(new_temp_amb),
                                        "affaissement": int(new_affaissement),
                                        "prelevement": new_prelevement,
                                        "nb_eprouvettes": int(new_nb_eprouvettes),
                                        "observations": new_observations
                                    }).eq("id", selected_item["id"]).execute()
                                    
                                    st.success("Modifications enregistrées avec succès !")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erreur de mise à jour : {e}")

                # --- BLOC DE SUPPRESSION (RESERVÉ UNIQUEMENT À L'ADMINISTRATEUR) ---
                if is_admin:
                    with col_del:
                        st.markdown("##### ⚠️ Suppression")
                        if st.button("🗑️ Supprimer définitivement ce contrôle", type="primary", key="btn_supprimer_admin"):
                            try:
                                supabase.table("suivi_betonnage").delete().eq("id", selected_item["id"]).execute()
                                st.success("Enregistrement supprimé avec succès.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erreur de suppression : {e}")

        else:
            st.info("Aucune donnée enregistrée pour le moment.")
            
    except Exception as e:
        st.error(f"Erreur lors de la récupération de l'historique : {e}")
