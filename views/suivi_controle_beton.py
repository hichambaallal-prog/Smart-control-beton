import io
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st

# Importations openpyxl pour le style et la mise en page A4 Excel
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def generer_pv_excel(lot_data, infos_header):
    """Génère un fichier Excel mis en page au format A4 Portrait pour impression.

    - Client : TGCC
    - Projet : LGV CASA
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PV Écrasement"

    # --- CONFIGURATION IMPRESSION A4 PORTRAIT ---
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Marges réduites
    ws.margins.left = 0.5
    ws.margins.right = 0.5
    ws.margins.top = 0.6
    ws.margins.bottom = 0.6

    # --- STYLES ---
    font_title = Font(name="Arial", size=14, bold=True, color="1F497D")
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_norm = Font(name="Arial", size=10)

    fill_navy = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    fill_sub_header = PatternFill(
        start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"
    )
    fill_zebra = PatternFill(
        start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
    )

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thick_bottom_side = Side(border_style="medium", color="1F497D")

    border_all = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )

    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # --- ENTÊTE DU DOCUMENT / TITRE ---
    ws.merge_cells("A1:F1")
    ws["A1"] = "LABORATOIRE DE CONTRÔLE DE QUALITÉ DES BÉTONS"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:F2")
    ws["A2"] = (
        "PROCÈS-VERBAL D'ESSAI D'ÉCRASEMENT D'ÉPROUVETTES DE BÉTON (NF EN"
        " 12390-3)"
    )
    ws["A2"].font = Font(name="Arial", size=11, bold=True, color="333333")
    ws["A2"].alignment = align_center

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 18

    # Ligne de séparation
    for col in range(1, 7):
        ws.cell(row=3, column=col).border = Border(bottom=thick_bottom_side)

    # --- BLOC INFORMATIONS PROJET & CLIENT ---
    ws.merge_cells("A5:F5")
    ws["A5"] = " 📌 INFORMATIONS GÉNÉRALES DU PROJET & PRÉLÈVEMENT"
    ws["A5"].font = font_header
    ws["A5"].fill = fill_navy
    ws["A5"].alignment = align_left
    ws.row_dimensions[5].height = 20

    info_grid = [
        (
            "Client :",
            infos_header.get("client", "TGCC"),
            "Projet :",
            infos_header.get("projet", "LGV CASA"),
        ),
        (
            "N° Bon de Livraison (BL) :",
            infos_header.get("num_bl", "N/A"),
            "Ouvrage / Élément :",
            infos_header.get("ouvrage", "N/A"),
        ),
        (
            "Classe de Béton Spécifiée :",
            infos_header.get("classe_beton", "N/A"),
            "Échéance Visée :",
            infos_header.get("echeance", "N/A"),
        ),
        (
            "Date de Coulée :",
            str(infos_header.get("date_coulee", "N/A")),
            "Date d'Écrasement :",
            str(infos_header.get("date_ecrasement", "N/A")),
        ),
        (
            "Affaissement / Slump :",
            str(infos_header.get("affaissement", "N/A")),
            "Température Béton :",
            str(infos_header.get("temperature", "N/A")),
        ),
        (
            "Opérateur / Technicien :",
            infos_header.get("technicien", "Technicien LPEE"),
            "Observations :",
            infos_header.get("observations", "-"),
        ),
    ]

    row_idx = 6
    for label1, val1, label2, val2 in info_grid:
        ws.cell(row=row_idx, column=1, value=label1).font = font_bold
        ws.cell(row=row_idx, column=1).alignment = align_left
        ws.cell(row=row_idx, column=1).fill = fill_sub_header

        ws.cell(row=row_idx, column=2, value=val1).font = font_norm
        ws.cell(row=row_idx, column=2).alignment = align_left

        ws.cell(row=row_idx, column=4, value=label2).font = font_bold
        ws.cell(row=row_idx, column=4).alignment = align_left
        ws.cell(row=row_idx, column=4).fill = fill_sub_header

        ws.merge_cells(
            start_row=row_idx, start_column=5, end_row=row_idx, end_column=6
        )
        ws.cell(row=row_idx, column=5, value=val2).font = font_norm
        ws.cell(row=row_idx, column=5).alignment = align_left

        for c in range(1, 7):
            ws.cell(row=row_idx, column=c).border = border_all

        ws.row_dimensions[row_idx].height = 18
        row_idx += 1

    row_idx += 1  # Espace

    # --- TABLEAU DES RÉSULTATS D'ÉCRASEMENT ---
    ws.merge_cells(
        start_row=row_idx, start_column=1, end_row=row_idx, end_column=6
    )
    ws.cell(
        row=row_idx,
        column=1,
        value=" 💥 RÉSULTATS DES ESSAIS DE RUPTURE SUR ÉPROUVETTES",
    ).font = font_header
    ws.cell(row=row_idx, column=1).fill = fill_navy
    ws.cell(row=row_idx, column=1).alignment = align_left
    ws.row_dimensions[row_idx].height = 20
    row_idx += 1

    headers = [
        "N°",
        "Repère Éprouvette",
        "Type / Forme",
        "Section (cm²)",
        "Force F (kN)",
        "Résistance Fc (MPa)",
    ]
    for col_i, h in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_i, value=h)
        cell.font = font_header
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = border_all

    ws.row_dimensions[row_idx].height = 22
    row_idx += 1

    start_data_row = row_idx
    for i, item in enumerate(lot_data, 1):
        sec = float(item.get("section") or 176.71)
        f_kn = float(item.get("force_kn") or 0.0)
        fc_mpa = float(item.get("fc_mpa") or 0.0)

        ws.cell(row=row_idx, column=1, value=i).alignment = align_center
        ws.cell(
            row=row_idx,
            column=2,
            value=str(item.get("repere_eprouvette", "")),
        ).alignment = align_left
        ws.cell(
            row=row_idx, column=3, value=str(item.get("forme", ""))
        ).alignment = align_center

        c_sec = ws.cell(row=row_idx, column=4, value=sec)
        c_sec.number_format = "0.00"
        c_sec.alignment = align_right

        c_f = ws.cell(row=row_idx, column=5, value=f_kn)
        c_f.number_format = "0.0"
        c_f.alignment = align_right

        c_fc = ws.cell(row=row_idx, column=6, value=fc_mpa)
        c_fc.number_format = "0.0"
        c_fc.alignment = align_right

        # Style zebra & bordures
        current_fill = (
            fill_zebra if i % 2 == 0 else PatternFill(fill_type=None)
        )
        for c in range(1, 7):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = font_norm
            cell.border = border_all
            if current_fill.fill_type:
                cell.fill = current_fill

        ws.row_dimensions[row_idx].height = 18
        row_idx += 1

    end_data_row = row_idx - 1

    # --- LIGNE STATISTIQUE : MOYENNE ---
    if end_data_row >= start_data_row:
        ws.merge_cells(
            start_row=row_idx, start_column=1, end_row=row_idx, end_column=5
        )
        ws.cell(
            row=row_idx,
            column=1,
            value="RÉSISTANCE MOYENNE DU LOT Fc,moy (MPa) :",
        ).font = font_bold
        ws.cell(row=row_idx, column=1).alignment = align_right
        ws.cell(row=row_idx, column=1).fill = fill_sub_header

        cell_avg = ws.cell(
            row=row_idx,
            column=6,
            value=f"=AVERAGE(F{start_data_row}:F{end_data_row})",
        )
        cell_avg.font = Font(name="Arial", size=11, bold=True, color="1F497D")
        cell_avg.number_format = "0.0"
        cell_avg.alignment = align_right
        cell_avg.fill = fill_sub_header

        for c in range(1, 7):
            ws.cell(row=row_idx, column=c).border = border_all

        ws.row_dimensions[row_idx].height = 22
        row_idx += 2

    # --- ZONE DE SIGNATURE ---
    ws.cell(row=row_idx, column=1, value="Le Technicien d'Essai :").font = (
        font_bold
    )
    ws.merge_cells(
        start_row=row_idx, start_column=4, end_row=row_idx, end_column=6
    )
    ws.cell(
        row=row_idx,
        column=4,
        value="Le Responsable de Laboratoire / Validation :",
    ).font = font_bold
    ws.cell(row=row_idx, column=4).alignment = align_right

    # LARGOUR AUTOMATIQUE DES COLONNES AJUSTÉE AU A4 PORTRAIT
    col_widths = {
        "A": 6,  # N°
        "B": 24,  # Repère Éprouvette
        "C": 22,  # Forme
        "D": 15,  # Section
        "E": 15,  # Force
        "F": 20,  # Fc
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    # Création des onglets
    tab_prog, tab_saisie, tab_hist = st.tabs([
        "📅 Phase 1 : Programmation",
        "💥 Phase 2 : Saisie des Écrasements (Par Lot)",
        "📋 Historique Complet",
    ])

    # ---------------------------------------------------------
    # RÉCUPÉRATION DES BÉTONNAGES PRÉLEVÉS (OUI)
    # ---------------------------------------------------------
    betonnages_preleves = []
    try:
        res_beton = (
            supabase.table("suivi_betonnage")
            .select("*")
            .order("id", desc=True)
            .execute()
        )
        if res_beton.data:
            betonnages_preleves = [
                item
                for item in res_beton.data
                if item.get("prelevement")
                and str(item.get("prelevement")).upper().startswith("OUI")
            ]
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des bétonnages : {e}")

    if not betonnages_preleves:
        st.info(
            "ℹ️ Aucun suivi de bétonnage avec prélèvement d'éprouvettes (OUI)"
            " trouvé."
        )
        return

    options_beton = {
        (
            f"ID #{b['id']} | BL: {b.get('num_bl', 'N/A')} | Ouvrage:"
            f" {b.get('ouvrage', 'N/A')} | Date:"
            f" {b.get('date_coulee', b.get('date_livraison', 'N/A'))} | Classe:"
            f" {b.get('classe_beton', b.get('classe', 'N/A'))}"
        ): b
        for b in betonnages_preleves
    }

    # =========================================================
    # PHASE 1 : PROGRAMMATION DES ÉPROUVETTES
    # =========================================================
    with tab_prog:
        st.subheader("📅 1. Programmer les Échéances d'Écrasement")

        choix_label_p = st.selectbox(
            "Sélectionner la fiche de bétonnage :",
            list(options_beton.keys()),
            key="prog_beton_select",
        )
        beton_p = options_beton[choix_label_p]

        b_id = beton_p.get("id")
        num_bl_p = str(beton_p.get("num_bl") or "N/A")
        ouvrage_p = str(beton_p.get("ouvrage") or "N/A")
        classe_beton_p = str(
            beton_p.get("classe_beton") or beton_p.get("classe") or "N/A"
        )

        raw_nb_ep = beton_p.get("nb_eprouvettes") or beton_p.get(
            "nombre_eprouvettes"
        )
        try:
            total_eprouvettes_prevues = (
                int(raw_nb_ep) if raw_nb_ep is not None else 12
            )
        except (ValueError, TypeError):
            total_eprouvettes_prevues = 12

        eprouvettes_deja_prog = 0
        try:
            res_deja = (
                supabase.table("suivi_controle_beton")
                .select("id")
                .eq("betonnage_id", b_id)
                .execute()
            )
            if res_deja.data:
                eprouvettes_deja_prog = len(res_deja.data)
        except Exception:
            eprouvettes_deja_prog = 0

        solde_disponible = max(
            0, total_eprouvettes_prevues - eprouvettes_deja_prog
        )

        affaissement_raw = str(
            beton_p.get("affaissement") or beton_p.get("slump") or "N/A"
        )
        temp_beton_p = str(
            beton_p.get("temperature") or beton_p.get("temp_beton") or "N/A"
        )
        affaissement_p = (
            f"{affaissement_raw} mm" if affaissement_raw != "N/A" else "N/A"
        )

        date_coulee_raw = (
            beton_p.get("date_coulee")
            or beton_p.get("date_livraison")
            or str(date.today())
        )
        try:
            date_coulee_p = datetime.strptime(
                str(date_coulee_raw), "%Y-%m-%d"
            ).date()
        except Exception:
            date_coulee_p = date.today()

        ref_controle_defaut = f"REF-{b_id}-{ouvrage_p}"

        st.markdown("---")
        st.info(
            f"📊 **Quota Éprouvettes :** Total prévu :"
            f" **{total_eprouvettes_prevues}** | Déjà programmée(s) :"
            f" **{eprouvettes_deja_prog}** | Reste disponible :"
            f" **{solde_disponible}**"
        )

        ref_controle_p = st.text_input(
            "🏷️ Référence de Contrôle (Préfixe du repère)",
            value=ref_controle_defaut,
            key=f"p_ref_ctrl_{b_id}",
        )

        st.markdown("---")
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            st.text_input(
                "N° Bon de Livraison (BL)",
                value=num_bl_p,
                disabled=True,
                key=f"p_bl_{b_id}",
            )
        with col_p2:
            st.text_input(
                "Ouvrage / Élément",
                value=ouvrage_p,
                disabled=True,
                key=f"p_ouv_{b_id}",
            )
        with col_p3:
            st.text_input(
                "Classe de Béton Spécifiée",
                value=classe_beton_p,
                disabled=True,
                key=f"p_classe_{b_id}",
            )

        col_p4, col_p5 = st.columns(2)
        with col_p4:
            st.text_input(
                "Affaissement / Slump (mm)",
                value=affaissement_p,
                disabled=True,
                key=f"p_aff_{b_id}",
            )
        with col_p5:
            st.text_input(
                "Température Béton Frais (°C)",
                value=(
                    f"{temp_beton_p} °C" if temp_beton_p != "N/A" else "N/A"
                ),
                disabled=True,
                key=f"p_temp_{b_id}",
            )

        st.markdown("---")
        col_e1, col_e2, col_e3, col_e4 = st.columns(4)
        with col_e1:
            echeance_p = st.selectbox(
                "Âge / Échéance visée",
                ["3 jours", "7 jours", "28 jours", "90 jours"],
                index=2,
                key=f"p_echeance_{b_id}",
            )

        jours_dict = {
            "3 jours": 3,
            "7 jours": 7,
            "28 jours": 28,
            "90 jours": 90,
        }
        nb_j = jours_dict.get(echeance_p, 28)
        date_prevue_auto = date_coulee_p + timedelta(days=nb_j)
        echeance_key_clean = echeance_p.replace(" ", "_")

        with col_e2:
            st.date_input(
                "Date de Coulée",
                value=date_coulee_p,
                disabled=True,
                key=f"p_date_coul_{b_id}",
            )
        with col_e3:
            date_ecrasement_prevue = st.date_input(
                "Date d'Écrasement Prévue",
                value=date_prevue_auto,
                key=f"p_date_ecras_{b_id}_{echeance_key_clean}",
            )

        min_val = 1 if solde_disponible > 0 else 0
        val_defaut = min(2, solde_disponible) if solde_disponible > 0 else 0

        with col_e4:
            if solde_disponible == 0:
                st.warning("⚠️ Quota atteint.")
                nb_eprouvettes_p = 0
            else:
                nb_eprouvettes_p = st.number_input(
                    "Nombre d'éprouvettes à programmer",
                    min_value=min_val,
                    max_value=solde_disponible,
                    value=val_defaut,
                    step=1,
                    key=f"p_nb_ep_{b_id}_{echeance_key_clean}",
                )

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            forme_p = st.selectbox(
                "Type / Forme d'éprouvette",
                [
                    "Cylindrique 150x300",
                    "Cylindrique 160x320",
                    "Cylindrique 100x200",
                ],
                key=f"p_forme_{b_id}",
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
            forme_key_clean = forme_p.replace(" ", "_").replace("x", "_")
            st.number_input(
                "Section Théorique (cm²)",
                value=sect_def,
                format="%.2f",
                disabled=True,
                key=f"p_section_{b_id}_{forme_key_clean}",
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
                        key=f"prog_rep_{b_id}_{echeance_key_clean}_{i}",
                    )
                    reperes_p.append(rep_val)

            if st.button(
                "📌 Enregistrer la Programmation",
                type="primary",
                use_container_width=True,
                key=f"btn_save_prog_{b_id}",
            ):
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
                        "section": float(sect_def),
                    }
                    try:
                        res = (
                            supabase.table("suivi_controle_beton")
                            .insert(payload_prog)
                            .execute()
                        )
                        if res.data:
                            succes_cnt += 1
                    except Exception as err:
                        st.error(
                            f"Erreur lors de la programmation de {rep} : {err}"
                        )

                if succes_cnt > 0:
                    st.success(
                        f"✅ {succes_cnt} éprouvette(s) programmée(s) pour le"
                        f" {date_ecrasement_prevue} ({echeance_p}) !"
                    )
                    st.rerun()

    # =========================================================
    # PHASE 2 : SAISIE GROUPÉE PAR ÉCHÉANCE / AGE
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
            col_l1, col_l2, col_l3, col_l4 = st.columns(4)
            col_l1.metric("Client", "TGCC")
            col_l2.metric("Projet", "LGV CASA")
            col_l3.metric("Ouvrage", str(sample.get("ouvrage")))
            col_l4.metric("Échéance Visée", str(sample.get("echeance")))

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
                    "Observations générales",
                    value="Rupture satisfaisante (NF EN 12390-3).",
                    key="obs_global",
                )

            st.markdown("##### 📝 Saisie des mesures pour le lot")

            lot_key = f"df_lot_{choix_lot}"

            # Initialisation de st.session_state pour ce lot précis
            if lot_key not in st.session_state:
                rows_list = []
                for ep in lot_selected:
                    sec = float(ep.get("section") or 176.71)
                    f_kn = float(ep.get("force_kn") or 0.0)
                    fc = (
                        round((f_kn * 10.0) / sec, 1)
                        if sec > 0 and f_kn > 0
                        else 0.0
                    )

                    rows_list.append({
                        "ID": ep["id"],
                        "Repère": ep.get(
                            "repere_eprouvette", f"EP-{ep['id']}"
                        ),
                        "Forme d'éprouvette": str(
                            ep.get("forme") or "Cylindrique 150x300"
                        ),
                        "_section": sec,
                        "Force (kN)": f_kn,
                        "Résistance Fc (MPa)": fc,
                    })
                st.session_state[lot_key] = pd.DataFrame(rows_list)

            # Callback corrigé pour la mise à jour de la résistance Fc
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

            edited_df = st.data_editor(
                st.session_state[lot_key],
                column_config={
                    "ID": st.column_config.NumberColumn("ID", disabled=True),
                    "Repère": st.column_config.TextColumn(
                        "Repère", disabled=True
                    ),
                    "Forme d'éprouvette": st.column_config.TextColumn(
                        "Forme d'éprouvette", disabled=True
                    ),
                    "_section": None,
                    "Force (kN)": st.column_config.NumberColumn(
                        "⚡ Force (kN)",
                        help=(
                            "Saisissez la force de rupture lue sur la presse"
                        ),
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
                st.success(
                    "📈 **Résistance moyenne calculée pour les éprouvettes"
                    f" saisies : {fc_moy:.1f} MPa**"
                )
            else:
                st.warning(
                    "👈 Veuillez remplir la colonne **Force (kN)** pour chaque"
                    " éprouvette."
                )

            col_b1, col_b2 = st.columns(2)

            with col_b1:
                if st.button(
                    "💾 Valider et Enregistrer Tout le Lot",
                    type="primary",
                    use_container_width=True,
                ):
                    if (df_actuel["Force (kN)"] == 0).any():
                        st.error(
                            "❌ Attention : Une ou plusieurs éprouvettes ont"
                            " encore une force de 0.0 kN."
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
                                    "Erreur sur l'éprouvette"
                                    f" {row['Repère']} : {e}"
                                )

                        if succes_lot == len(df_actuel):
                            del st.session_state[lot_key]
                            st.success(
                                f"✅ Lot de {succes_lot} éprouvettes"
                                " enregistré avec succès !"
                            )
                            st.rerun()

            with col_b2:
                # Préparation des données pour le PV Excel
                export_data = []
                for _, row in df_actuel.iterrows():
                    export_data.append({
                        "repere_eprouvette": row["Repère"],
                        "forme": row["Forme d'éprouvette"],
                        "section": row["_section"],
                        "force_kn": row["Force (kN)"],
                        "fc_mpa": row["Résistance Fc (MPa)"],
                    })

                infos_header = {
                    "client": "TGCC",
                    "projet": "LGV CASA",
                    "num_bl": sample.get("num_bl", "N/A"),
                    "ouvrage": sample.get("ouvrage", "N/A"),
                    "classe_beton": sample.get("classe_beton", "N/A"),
                    "echeance": sample.get("echeance", "N/A"),
                    "date_coulee": sample.get("date_coulee", "N/A"),
                    "date_ecrasement": sample.get("date_ecrasement", "N/A"),
                    "affaissement": sample.get("affaissement", "N/A"),
                    "temperature": sample.get("temperature", "N/A"),
                    "technicien": tech_global,
                    "observations": obs_globale,
                }

                excel_file = generer_pv_excel(export_data, infos_header)
                filename = (
                    f"PV_Ecrasement_TGCC_LGV_CASA_{sample.get('num_bl', 'BL')}_{sample.get('echeance', '28j')}.xlsx"
                )

                st.download_button(
                    label="📄 Télécharger le PV d'écrasement (Excel A4)",
                    data=excel_file,
                    file_name=filename,
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

    # =========================================================
    # HISTORIQUE ET SUIVI GLOBAL
    # =========================================================
    with tab_hist:
        st.subheader("📋 Historique Général des Contrôles de Béton")

        try:
            res_all = (
                supabase.table("suivi_controle_beton")
                .select("*")
                .order("id", desc=True)
                .execute()
            )
            if res_all.data:
                df_all = pd.DataFrame(res_all.data)
                cols_display = [
                    "id",
                    "num_bl",
                    "ouvrage",
                    "classe_beton",
                    "date_coulee",
                    "echeance",
                    "date_ecrasement",
                    "repere_eprouvette",
                    "forme",
                    "force_kn",
                    "fc_mpa",
                    "technicien",
                ]
                cols_valid = [c for c in cols_display if c in df_all.columns]
                st.dataframe(
                    df_all[cols_valid],
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info(
                    "Aucun enregistrement d'écrasement dans la base de"
                    " données."
                )
        except Exception as e:
            st.error(f"Erreur lors du chargement de l'historique : {e}")
