# ==============================================================================
# 4. APPLICATION STREAMLIT (CODE COMPLET MIS À JOUR)
# ==============================================================================
def show(supabase):
  st.title("📋 Historique & Procès-Verbaux d'Écrasement (NF EN 12390)")

  user_info = st.session_state.get("user", {})
  role = str(
      st.session_state.get("user_role")
      or st.session_state.get("role")
      or user_info.get("role", "")
  ).lower()
  can_edit = st.session_state.get("can_edit", False) or bool(
      user_info.get("can_edit", False)
  )
  is_admin = (
      role in ["admin", "responsable_labo"]
      or st.session_state.get("is_admin", False)
  )

  if (
      role not in ["laboratoire", "labo", "admin", "responsable_labo", "qualite"]
      and not is_admin
      and not can_edit
  ):
    st.error("⛔ **Accès Restreint**")
    st.warning(
        "Ce module est réservé exclusivement au personnel du **Laboratoire de"
        " Contrôle**."
    )
    return

  st.subheader("📋 Historique Général & Consultation des PVs")

  try:
    res_all = (
        supabase.table("suivi_controle_beton")
        .select("*")
        .order("id", desc=True)
        .execute()
    )
    if not res_all.data:
      st.info("ℹ️ Aucun historique disponible dans la base de données.")
      return

    df_all = pd.DataFrame(res_all.data)

    # Récupération de l'ensemble des fiches parents
    unique_b_ids = [
        b_id for b_id in df_all["betonnage_id"].unique() if pd.notnull(b_id)
    ]
    unique_parents = {
        b_id: obtenir_infos_betonnage_parent(supabase, b_id)
        for b_id in unique_b_ids
    }

    # -------------------------------------------------------------------------
    # FILTRE : Conserver uniquement les lignes dont le bétonnage est VALIDÉ & SIGNÉ
    # -------------------------------------------------------------------------
    def verifier_pv_valide(row):
      b_id = row.get("betonnage_id")
      parent = unique_parents.get(b_id) or {}

      # Vérification de la force écrasée
      f_kn = row.get("force_kn")
      a_force = pd.notnull(f_kn) and float(f_kn) > 0

      # Vérification du statut de validation (table parent ou ligne)
      statut_valide = (
          parent.get("statut_pv") == "Validé"
          or parent.get("validation_admin") is True
          or row.get("statut_pv") == "Validé"
          or row.get("validation_admin") is True
      )

      # Vérification de la présence des signatures / visas
      a_visa = bool(
          parent.get("visa_chef")
          or parent.get("visa_admin")
          or row.get("visa_chef")
          or row.get("visa_admin")
      )

      return a_force and statut_valide and a_visa

    # Filtrage des lignes validées
    mask_valides = df_all.apply(verifier_pv_valide, axis=1)
    df_valides = df_all[mask_valides].copy()

    st.markdown("##### 📥 Re-télécharger un Procès-Verbal")

    if df_valides.empty:
      st.info(
          "ℹ️ Aucun Procès-Verbal **validé et signé** n'est disponible pour"
          " le téléchargement."
      )
    else:
      c_r1, c_r2 = st.columns(2)
      recherche_pv = c_r1.text_input(
          "🔍 Rechercher (réf, ouvrage, classe...)",
          placeholder="Ex: gare casa sud, B/394...",
          key="search_input_pv",
      )
      recherche_date_pv = c_r2.text_input(
          "📅 Rechercher par Date d'écrasement",
          placeholder="Ex: 2026-08-08",
          key="search_date_pv",
      )

      groupes_valides = {}
      for _, row in df_valides.iterrows():
        b_id = row.get("betonnage_id")
        info_b = unique_parents.get(b_id) or {}
        ref_ctrl = determiner_ref_controle(
            supabase, b_id, info_b, row.to_dict()
        )
        classe = (
            row.get("classe_beton")
            or (info_b.get("classe_beton") or info_b.get("classe") if info_b else "-")
            or "-"
        )

        cle = (
            f"Référence : {ref_ctrl} | Classe : {classe} | Ouvrage :"
            f" {row.get('ouvrage', '-')} | Échéance : {row.get('echeance', '28 jours')}"
            f" (Date : {row.get('date_ecrasement', '-')}) | Lot ID #{b_id}"
        )
        groupes_valides.setdefault(cle, []).append(row.to_dict())

      pvs_filtrés = [
          k
          for k in groupes_valides.keys()
          if (not recherche_pv or recherche_pv.lower() in k.lower())
          and (not recherche_date_pv or recherche_date_pv in k)
      ]

      if not pvs_filtrés:
        st.warning(
            "Aucun PV validé et signé ne correspond à votre recherche."
        )
      else:
        choix_pv = st.selectbox(
            "Sélectionnez le PV à consulter :",
            pvs_filtrés,
            key="select_pv_hist",
        )
        lot_hist = groupes_valides[choix_pv]
        sample_h = lot_hist[0]
        b_id_h = sample_h.get("betonnage_id")

        info_b_h = unique_parents.get(b_id_h) or {}
        essais_h = obtenir_historique_betonnage(supabase, b_id_h) or lot_hist

        export_data_h = []
        for item in essais_h:
          sec = float(item.get("section") or 176.71)
          f_kn = float(item.get("force_kn") or 0.0)
          fc = float(
              item.get("fc_mpa")
              or (round((f_kn * 10.0) / sec, 1) if f_kn > 0 else 0.0)
          )
          ref_p = str(item.get("ref_controle") or "").strip()
          rep_s = str(
              item.get("repere_eprouvette", f"/{item['id']}")
          ).strip()

          export_data_h.append({
              "repere_eprouvette": f"{ref_p}{rep_s}" if ref_p else rep_s,
              "forme": item.get("forme", "Cylindrique 150x300"),
              "section": sec,
              "force_kn": f_kn,
              "fc_mpa": fc,
              "date_essai": item.get("date_ecrasement", "-"),
              "age": (
                  str(item.get("age", "28"))
                  .replace(" jours", "")
                  .replace("j", "")
              ),
              "statut": "En cours" if f_kn == 0 else "Réalisé",
          })

        num_bl_h = extraire_num_bl(sample_h, info_b_h, choix_pv)
        ouv_h = info_b_h.get("ouvrage") or sample_h.get("ouvrage")

        infos_header_h = {
            "re_num": "25/260/LGV/ B/",
            "dossier": "2025-260-05985-2025-0247",
            "client": "TGCC",
            "num_bl": num_bl_h,
            "ouvrage": ouv_h,
            "lieu_prelevement": ouv_h,
            "classe_beton": sample_h.get("classe_beton", "C35/45"),
            "date_coulee": info_b_h.get("date_coulee")
            or sample_h.get("date_coulee"),
            "affaissement": (
                info_b_h.get("affaissement") or info_b_h.get("slump")
            ),
            "temperature": (
                info_b_h.get("temperature") or info_b_h.get("temp_beton")
            ),
            "forme": sample_h.get("forme", "Cylindrique 150x300"),
            "centrale": (
                info_b_h.get("centrale")
                or info_b_h.get("centrale_beton")
                or sample_h.get("centrale")
            ),
            "observations": (
                info_b_h.get("observations_admin")
                or sample_h.get("observations")
                or "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"
            ),
            "technicien_prelevement": (
                info_b_h.get("technicien_prelevement")
                or info_b_h.get("preleve_par")
                or info_b_h.get("technicien")
                or sample_h.get("technicien")
            ),
        }

        # Bouton de téléchargement accessible uniquement pour ces PVs validés
        st.download_button(
            label="📄 Télécharger le PV (Excel Format LPEE)",
            data=generer_pv_excel(export_data_h, infos_header_h),
            file_name=(
                f"PV_Ecrasement_LPEE_{num_bl_h if num_bl_h != '-' else 'BL'}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
            key="btn_download_hist",
        )

    # -------------------------------------------------------------------------
    # TABLEAU GLOBAL DE TOUTE LA BASE DE DONNÉES
    # -------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("##### 📊 Base de données globale")

    df_all["affaissement_mm"] = df_all["betonnage_id"].map(
        lambda b: (unique_parents.get(b) or {}).get("affaissement")
        or (unique_parents.get(b) or {}).get("slump")
        or "-"
    )
    df_all["temp_beton_C"] = df_all["betonnage_id"].map(
        lambda b: (unique_parents.get(b) or {}).get("temperature")
        or (unique_parents.get(b) or {}).get("temp_beton")
        or "-"
    )

    df_all["statut_validation"] = df_all.apply(
        lambda r: (
            "✅ Validé & Signé"
            if verifier_pv_valide(r)
            else "⏳ En attente de validation"
        ),
        axis=1,
    )

    cols_ordre = [
        "id",
        "betonnage_id",
        "repere_eprouvette",
        "num_bl",
        "ouvrage",
        "classe_beton",
        "statut_validation",
        "date_coulee",
        "affaissement_mm",
        "temp_beton_C",
        "echeance",
        "date_ecrasement",
        "fc_mpa",
        "technicien",
    ]
    exclus = {
        "forme",
        "section",
        "force_kn",
        "observations",
        "masse",
        "ref_controle",
        "reference_controle",
        "refernce_controle",
        "num_reception",
    }

    cols_finales = [
        c
        for c in cols_ordre
        + [c for c in df_all.columns if c not in cols_ordre]
        if c not in exclus
    ]
    df_final = df_all[cols_finales]

    c_s1, c_s2 = st.columns(2)
    search_ref = c_s1.text_input(
        "🔍 Recherche par Réf. Contrôle",
        placeholder="Ex: REF-123-GARE CASA SUD",
    )
    search_date = c_s2.text_input(
        "📅 Recherche par Date de coulée", placeholder="Ex: 2026-08-24"
    )

    if search_ref:
      ref_col = next(
          (
              c
              for c in [
                  "ref_controle",
                  "reference_controle",
                  "refernce_controle",
              ]
              if c in df_all.columns
          ),
          None,
      )
      if ref_col:
        df_final = df_final[
            df_all[ref_col]
            .astype(str)
            .str.contains(search_ref, case=False, na=False)
        ]
    if search_date:
      df_final = df_final[
          df_final["date_coulee"]
          .astype(str)
          .str.contains(search_date, case=False, na=False)
      ]

    st.dataframe(df_final, use_container_width=True, hide_index=True)

    st.download_button(
        label="📊 Télécharger la base de données globale (Excel)",
        data=exporter_dataframe_excel(df_final, "Historique_Global"),
        file_name=f"Historique_Global_Beton_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="btn_download_hist_global",
    )

  except Exception as e:
    st.error(f"Erreur lors du chargement de l'historique global : {e}")
