import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta

def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    # 1. RÉCUPÉRATION DES BÉTONNAGES AVEC PRÉLÈVEMENT "OUI"
    betonnages_preleves = []
    try:
        res_beton = supabase.table("suivi_betonnage").select("*").order("id", desc=True).execute()
        if res_beton.data:
            # Filtrer les enregistrements où un prélèvement a été effectué
            betonnages_preleves = [
                item for item in res_beton.data
                if item.get("prelevement") and str(item.get("prelevement")).upper().startswith("OUI")
            ]
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des bétonnages : {e}")

    if not betonnages_preleves:
        st.info("ℹ️ Aucun bétonnage avec prélèvement d'éprouvettes enregistré pour le moment.")
        return

    # Option d'affichage dans la liste déroulante
    options_beton = {
        f"BL: {b.get('num_bl', 'N/A')} | Ouvrage: {b.get('ouvrage', 'N/A')} | Date: {b.get('date_coulee', 'N/A')} | Classe: {b.get('classe_beton', 'N/A')}": b
        for b in betonnages_preleves
    }

    st.subheader("📋 1. Sélection du Prélèvement")
    choix_beton_label = st.selectbox("Sélectionner une fiche de bétonnage :", list(options_beton.keys()))
    beton_selectionne = options_beton[choix_beton_label]

    # --- Récupération dynamique de la Classe de Béton ---
    classe_beton_auto = beton_selectionne.get("classe_beton", "Non spécifiée")
    date_coulee_str = beton_selectionne.get("date_coulee")
    
    # Conversion de la date de coulée
    if date_coulee_str:
        if isinstance(date_coulee_str, str):
            date_coulee_obj = datetime.strptime(date_coulee_str, "%Y-%m-%d").date()
        else:
            date_coulee_obj = date_coulee_str
    else:
        date_coulee_obj = date.today()

    st.markdown("---")
    st.subheader("🗓️ 2. Programmation et Saisie de l'Écrasement")

    col_info1, col_info2, col_info3 = st.columns(3)
    with col_info1:
        st.text_input("N° Bon de Livraison (BL)", value=str(beton_selectionne.get("num_bl", "")), disabled=True)
    with col_info2:
        st.text_input("Ouvrage / Elément", value=str(beton_selectionne.get("ouvrage", "")), disabled=True)
    with col_info3:
        # Affichage dynamique de la Classe de Béton provenant du suivi de bétonnage
        st.text_input("Classe de Béton Spécifiée", value=str(classe_beton_auto), disabled=True)

    col_echeance, col_nbr, col_date_c, col_date_e = st.columns(4)

    with col_echeance:
        echeance = st.selectbox("Échéance / Âge du Béton", ["3 jours", "7 jours", "28 jours", "90 jours"], index=2)

    # Calcul automatique des jours à ajouter
    jours_dict = {"3 jours": 3, "7 jours": 7, "28 jours": 28, "90 jours": 90}
    nb_jours = jours_dict.get(echeance, 28)

    # Calcul automatique de la date d'écrasement prévue
    date_ecrasement_auto = date_coulee_obj + timedelta(days=nb_jours)

    with col_nbr:
        nb_eprouvettes = st.number_input("Nombre d'éprouvettes à écraser", min_value=1, max_value=6, value=2, step=1)

    with col_date_c:
        st.date_input("Date de Coulée", value=date_coulee_obj, disabled=True)

    with col_date_e:
        # La date d'écrasement est calculée automatiquement (3j, 7j, 28j, 90j)
        date_ecrasement = st.date_input("Date d'Écrasement Effective", value=date_ecrasement_auto)

    st.markdown("---")
    st.markdown("### 🧪 Résultats d'Écrasement")

    # Formulaire de saisie pour chaque éprouvette
    eprouvettes_data = []
    
    # Forme et Section par défaut (Cylindrique 15x30 -> ~176.7 cm²)
    col_forme, col_section = st.columns(2)
    with col_forme:
        forme = st.selectbox("Type d'éprouvette", ["Cylindrique 15x30", "Cubique 15x15", "Cylindrique 11x22"])
    with col_section:
        if "15x30" in forme:
            section_defaut = 176.7
        elif "15x15" in forme:
            section_defaut = 225.0
        else:
            section_defaut = 95.03
        section_cm2 = st.number_input("Section (cm²)", value=section_defaut, format="%.2f")

    st.caption("Saisissez la force de rupture (kN) et la masse (g) pour chaque éprouvette :")

    cols_eprouvettes = st.columns(int(nb_eprouvettes))

    for idx in range(int(nb_eprouvettes)):
        with cols_eprouvettes[idx]:
            st.markdown(f"**Éprouvette N° {idx + 1}**")
            repere = st.text_input(f"Repère #{idx + 1}", value=f"E{idx+1}-{echeance.replace(' ', '')}", key=f"rep_{idx}")
            masse = st.number_input(f"Masse (g) #{idx + 1}", value=12500.0, step=10.0, key=f"masse_{idx}")
            force_kn = st.number_input(f"Force (kN) #{idx + 1}", value=500.0, step=5.0, key=f"force_{idx}")

            # Calcul de la résistance Fc = (Force en N) / (Section en mm²) = (Force in kN * 10) / (Section in cm²)
            fc_mpa = (force_kn * 10.0) / section_cm2 if section_cm2 > 0 else 0.0
            st.metric(label=f"Résistance Fc #{idx + 1}", value=f"{fc_mpa:.2f} MPa")

            eprouvettes_data.append({
                "repere": repere,
                "masse": masse,
                "force_kn": force_kn,
                "fc_mpa": round(fc_mpa, 2)
            })

    observations = st.text_area("Observations / Mode de rupture", value="Rupture satisfaisante (NF EN 12390-3).")
    technicien = st.text_input("Technicien / Opérateur", value=st.session_state.get("username", "Technicien LPEE"))

    # Enregistrement dans Supabase
    if st.button("💾 Enregistrer l'écrasement", use_container_width=True, type="primary"):
        enregistrements_succes = 0
        
        for ep in eprouvettes_data:
            data_to_insert = {
                "betonnage_id": beton_selectionne.get("id"),
                "num_bl": beton_selectionne.get("num_bl"),
                "ouvrage": beton_selectionne.get("ouvrage"),
                "classe_beton": classe_beton_auto,  # Classe dynamique récupérée automatiquement
                "date_coulee": str(date_coulee_obj),
                "echeance": echeance,
                "date_ecrasement": str(date_ecrasement),
                "repere_eprouvette": ep["repere"],
                "forme": forme,
                "section": section_cm2,
                "masse": ep["masse"],
                "force_kn": ep["force_kn"],
                "fc_mpa": ep["fc_mpa"],
                "technicien": technicien,
                "observations": observations
            }

            try:
                res = supabase.table("suivi_controle_beton").insert(data_to_insert).execute()
                if res.data:
                    enregistrements_succes += 1
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement de l'éprouvette {ep['repere']} : {e}")

        if enregistrements_succes > 0:
            st.success(f"✅ {enregistrements_succes} éprouvette(s) enregistrée(s) avec succès pour l'échéance {echeance} !")
            st.rerun()

    # --- 3. HISTORIQUE DES ÉCRASEMENTS ---
    st.markdown("---")
    st.subheader("📜 Historique des Écrasements pour cet Ouvrage")
    try:
        res_hist = supabase.table("suivi_controle_beton")\
            .select("*")\
            .eq("betonnage_id", beton_selectionne.get("id"))\
            .order("date_ecrasement", desc=True)\
            .execute()

        if res_hist.data:
            df_hist = pd.DataFrame(res_hist.data)
            col_order = ["echeance", "date_coulee", "date_ecrasement", "repere_eprouvette", "classe_beton", "masse", "force_kn", "fc_mpa", "technicien"]
            cols_existantes = [c for c in col_order if c in df_hist.columns]
            st.dataframe(df_hist[cols_existantes], use_container_width=True)
        else:
            st.info("Aucun essai d'écrasement enregistré pour cette fiche de bétonnage.")
    except Exception as e:
        st.error(f"Erreur lors du chargement de l'historique : {e}")
