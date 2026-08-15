from datetime import date, datetime, timedelta
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st


# =========================================================
# GÉNÉRATION DU PROCES-VERBAL EXCEL (FORMAT EXACT LPEE)
# =========================================================
def generer_pv_excel(export_data, infos_header):
    """Génère un Procès-Verbal (PV) d'écrasement de béton répliquant exactement le modèle LPEE / CTR CSB."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PV Écrasement LPEE"

    # Configuration A4 Portrait
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.page_margins = PageMargins(
        left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.2, footer=0.2
    )

    # Styles Typographiques
    font_bold = Font(name="Calibri", size=9, bold=True)
    font_regular = Font(name="Calibri", size=8.5)
    font_small = Font(name="Calibri", size=8)

    # Alignements
    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(
        horizontal="right", vertical="center", wrap_text=True
    )

    # Bordures
    thin_side = Side(border_style="thin", color="000000")
    border_cell = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )

    # 1. ENTÊTE DU LABORATOIRE ET REFERENCES
    ws.merge_cells("A1:D1")
    ws["A1"] = "LPEE / CTR CSB"
    ws["A1"].font = font_bold
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:D3")
    ws["A2"] = "Laboratoire de Contrôle Externe"
    ws["A2"].font = font_bold
    ws["A2"].alignment = align_center

    ws["E1"] = "RE N° :"
    ws["E1"].font = font_bold
    ws.merge_cells("F1:H1")
    ws["F1"] = infos_header.get("re_num", "25/260/LGV/ B/01")
    ws["F1"].font = font_regular

    ws["E2"] = "DOSSIER :"
    ws["E2"].font = font_bold
    ws.merge_cells("F2:H2")
    ws["F2"] = infos_header.get("dossier", "2025-260-05985-2025-0247")
    ws["F2"].font = font_regular

    ws["E3"] = "CLIENT :"
    ws["E3"].font = font_bold
    ws.merge_cells("F3:H3")
    ws["F3"] = infos_header.get("client", "TGCC")
    ws["F3"].font = font_bold

    for r in range(1, 4):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # 2. TITRE DE L'ESSAI ET TYPE D'ESSAI
    ws.merge_cells("A4:H4")
    ws["A4"] = "ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
    ws["A4"].font = Font(name="Calibri", size=11, bold=True)
    ws["A4"].alignment = align_center
    ws["A4"].border = border_cell

    ws.merge_cells("A5:D5")
    ws["A5"] = "[X]  COMPRESSION NF EN 12390-3 (2019)"
    ws["A5"].font = font_bold
    ws["A5"].alignment = align_center

    ws.merge_cells("E5:H5")
    ws["E5"] = "[  ]  TRACTION PAR FENDAGE NF EN 12390-6 (2019)"
    ws["E5"].font = font_bold
    ws["E5"].alignment = align_center

    for c in range(1, 9):
        ws.cell(row=5, column=c).border = border_cell

    # Ligne de la Presse
    ws.merge_cells("A6:F6")
    ws["A6"] = "Presse : Marque: Controls"
    ws["A6"].font = font_bold
    ws["A6"].alignment = align_left

    ws.merge_cells("G6:H6")
    ws["G6"] = "Classe : A"
    ws["G6"].font = font_bold
    ws["G6"].alignment = align_center

    for c in range(1, 9):
        ws.cell(row=6, column=c).border = border_cell

    # 3. FICHE TECHNIQUE DE PRELEVEMENT ET CHANTIER
    ws["A7"] = "Date de\nprélèvement"
    ws["A7"].font = font_bold
    ws["A7"].alignment = align_center
    ws["B7"] = str(infos_header.get("date_coulee", "02/06/2025"))
    ws["B7"].font = font_bold
    ws["B7"].alignment = align_center

    ws.merge_cells("C7:D7")
    ws["C7"] = "Lieu de\nprélèvement"
    ws["C7"].font = font_bold
    ws["C7"].alignment = align_center

    ws.merge_cells("E7:H7")
    ws["E7"] = infos_header.get(
        "lieu_prelevement",
        "Gros béton de la semelle C1 S2 Pro 745 bis Côté Marrakech 1° Partie",
    )
    ws["E7"].font = font_regular
    ws["E7"].alignment = align_center

    ws.merge_cells("A8:A9")
    ws["A8"] = "Chantier"
    ws["A8"].font = font_bold
    ws["A8"].alignment = align_center

    ws.merge_cells("B8:D9")
    ws["B8"] = infos_header.get(
        "chantier",
        "Augmentation de la capacité ferroviaire entre Kenitra et Marrakech et au niveau du hub de Casablanca\nTravaux d'exécution de terrassement, ouvrages d'art et rétablissement de communication entre PK 5+450 et PK 10+000-GARE CASA SUD",
    )
    ws["B8"].font = font_small
    ws["B8"].alignment = align_center

    ws.merge_cells("E8:F8")
    ws["E8"] = "Type de béton"
    ws["E8"].font = font_bold
    ws["E8"].alignment = align_center

    ws.merge_cells("G8:H8")
    ws["G8"] = infos_header.get("classe_beton", "C30/37")
    ws["G8"].font = font_bold
    ws["G8"].alignment = align_center

    ws.merge_cells("E9:H9")
    ws["E9"] = "EPROUVETTES"
    ws["E9"].font = font_bold
    ws["E9"].alignment = align_center

    ws.merge_cells("A10:B10")
    ws["A10"] = infos_header.get("centrale", "TG Prefa Oulad Saleh")
    ws["A10"].font = font_bold
    ws["A10"].alignment = align_center

    ws["C10"] = "- Dimensions"
    ws["C10"].font = font_regular
    ws["D10"] = "Φ"
    ws["D10"].font = font_bold
    ws["D10"].alignment = align_center
    ws["E10"] = "15"
    ws["E10"].alignment = align_center
    ws.merge_cells("F10:H10")
    ws["F10"] = "30"
    ws["F10"].alignment = align_center

    ws.merge_cells("A11:B11")
    ws["A11"] = "Affaissement au cône d'abrams NF EN 12350-2"
    ws["A11"].font = font_small
    ws["A11"].alignment = align_center

    ws["C11"] = str(infos_header.get("affaissement", "200"))
    ws["C11"].font = font_bold
    ws["C11"].alignment = align_center

    ws["D11"] = "- Mode confection"
    ws["D11"].font = font_regular

    ws.merge_cells("E11:H11")
    ws["E11"] = "Par vibration  NF EN 12390-2 (2019)"
    ws["E11"].font = font_bold
    ws["E11"].alignment = align_center

    ws.merge_cells("A12:B12")
    ws["A12"] = "Température °C"
    ws["A12"].font = font_regular
    ws["A12"].alignment = align_center

    ws["C12"] = str(infos_header.get("temperature", "31"))
    ws["C12"].font = font_bold
    ws["C12"].alignment = align_center

    ws["D12"] = "- Mode conservation"
    ws["D12"].font = font_regular

    ws.merge_cells("E12:H12")
    ws["E12"] = (
        "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C ± 2°C"
    )
    ws["E12"].font = font_bold
    ws["E12"].alignment = align_center

    ws.merge_cells("A13:C13")
    ws["A13"] = "Densité du béton durci NF EN 12390-7(2019)"
    ws["A13"].font = font_small
    ws["A13"].alignment = align_center

    ws.merge_cells("D13:E13")
    ws["D13"] = "N° de bon de livraison"
    ws["D13"].font = font_regular
    ws["D13"].alignment = align_center

    ws.merge_cells("F13:H13")
    ws["F13"] = str(infos_header.get("num_bl", "15479"))
    ws["F13"].font = font_bold
    ws["F13"].alignment = align_center

    for r in range(7, 14):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # 4. TABLEAU DES RÉSULTATS D'ÉCRASEMENT
    ws.merge_cells("A14:A15")
    ws["A14"] = "Réf,"
    ws["A14"].font = font_bold
    ws["A14"].alignment = align_center

    ws.merge_cells("B14:C14")
    ws["B14"] = "Date"
    ws["B14"].font = font_bold
    ws["B14"].alignment = align_center

    ws["B15"] = "Fabri"
    ws["B15"].font = font_regular
    ws["B15"].alignment = align_center

    ws["C15"] = "Essai"
    ws["C15"].font = font_regular
    ws["C15"].alignment = align_center

    ws.merge_cells("D14:D15")
    ws["D14"] = "Age (jours)"
    ws["D14"].font = font_bold
    ws["D14"].alignment = align_center

    ws.merge_cells("E14:E15")
    ws["E14"] = "Charge rupture(KN)"
    ws["E14"].font = font_bold
    ws["E14"].alignment = align_center

    ws.merge_cells("F14:H14")
    ws["F14"] = "Résistance (MPa)"
    ws["F14"].font = font_bold
    ws["F14"].alignment = align_center

    ws["F15"] = "Compression"
    ws["F15"].font = font_regular
    ws["F15"].alignment = align_center

    ws["G15"] = "Traction"
    ws["G15"].font = font_regular
    ws["G15"].alignment = align_center

    ws["H15"] = "Moyenne"
    ws["H15"].font = font_regular
    ws["H15"].alignment = align_center

    for r in range(14, 16):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    row_start = 16
    nb_total = len(export_data)

    if nb_total > 0:
        ws.merge_cells(f"A{row_start}:A{row_start + nb_total - 1}")
        ws[f"A{row_start}"] = "B/01"
        ws[f"A{row_start}"].font = font_bold
        ws[f"A{row_start}"].alignment = align_center

        ws.merge_cells(f"B{row_start}:B{row_start + nb_total - 1}")
        ws[f"B{row_start}"] = str(infos_header.get("date_coulee", "02/06/2025"))
        ws[f"B{row_start}"].font = font_bold
        ws[f"B{row_start}"].alignment = align_center

    for idx, item in enumerate(export_data):
        curr_row = row_start + idx

        ws.cell(row=curr_row, column=3, value=str(item.get("date_essai", "09/06/2025"))).alignment = align_center
        ws.cell(row=curr_row, column=4, value=item.get("age", 7)).alignment = align_center

        f_kn = float(item.get("force_kn", 0.0))
        ws.cell(row=curr_row, column=5, value=f"{f_kn:.1f}".replace(".", ",")).alignment = align_right

        fc_mpa = float(item.get("fc_mpa", 0.0))
        ws.cell(row=curr_row, column=6, value=f"{fc_mpa:.1f}".replace(".", ",")).alignment = align_right

        ws.cell(row=curr_row, column=7, value="-").alignment = align_center

        for c in range(1, 9):
            ws.cell(row=curr_row, column=c).font = font_regular
            ws.cell(row=curr_row, column=c).border = border_cell

    # Fusions conditionnelles pour la présentation de la moyenne
    if nb_total >= 3:
        ws.merge_cells(f"C16:C18")
        ws.merge_cells(f"D16:D18")
        ws.merge_cells(f"H16:H18")
        ws["H16"] = "=ROUND(AVERAGE(F16:F18),1)"
        ws["H16"].alignment = align_center
        ws["H16"].font = font_bold

    if nb_total >= 12:
        ws.merge_cells(f"C19:C27")
        ws.merge_cells(f"D19:D27")

        ws.merge_cells(f"H19:H21")
        ws["H19"] = "=ROUND(AVERAGE(F19:F21),1)"
        ws["H19"].alignment = align_center
        ws["H19"].font = font_bold

        ws.merge_cells(f"H22:H27")
        ws["H22"] = "=ROUND(AVERAGE(F22:F27),1)"
        ws["H22"].alignment = align_center
        ws["H22"].font = font_bold

    # 5. PIED DE PAGE ET COMMENTAIRE
    last_row = row_start + max(nb_total, 1)

    ws.cell(row=last_row, column=1, value="COMMENTAIRE :").font = font_bold
    ws.cell(row=last_row, column=1).alignment = align_center

    ws.merge_cells(start_row=last_row, start_column=2, end_row=last_row, end_column=8)
    ws.cell(
        row=last_row,
        column=2,
        value=infos_header.get(
            "observations", "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"
        ),
    ).font = font_bold
    ws.cell(row=last_row, column=2).alignment = align_left

    for c in range(1, 9):
        ws.cell(row=last_row, column=c).border = border_cell

    # Adjusting Column Widths
    col_widths = {
        "A": 8,
        "B": 12,
        "C": 12,
        "D": 10,
        "E": 18,
        "F": 14,
        "G": 12,
        "H": 12,
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================================================
# APPLICATION PRINCIPALE STREAMLIT
# =========================================================
def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    tab_prog, tab_saisie, tab_hist = st.tabs([
        "📅 Phase 1 : Programmation",
        "💥 Phase 2 : Saisie des Écrasements (Par Lot)",
        "📋 Historique Complet",
    ])

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

    # ---------------------------------------------------------
    # PHASE 1 : PROGRAMMATION
    # ---------------------------------------------------------
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

        sect_def = 176.71 if "150x300" in forme_p else 201.06

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
                    rep_defaut = f"/{num_ep}"
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

    # ---------------------------------------------------------
    # ---------------------------------------------------------
# PHASE 2 : SAISIE DES ÉCRASEMENTS ET PV (CORRIGÉE)
    # ---------------------------------------------------------
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
                    "Commentaire / Observation",
                    value=(
                        "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"
                    ),
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
                        round((f_kn * 10.0) / sec, 1)
                        if sec > 0 and f_kn > 0
                        else 0.0
                    )

                    rows_list.append({
                        "ID": ep["id"],
                        "Repère": ep.get(
                            "repere_eprouvette", f"/{ep['id']}"
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
                    "Repère": st.column_config.TextColumn(
                        "Repère", disabled=True
                    ),
                    "Forme d'éprouvette": st.column_config.TextColumn(
                        "Forme d'éprouvette", disabled=True
                    ),
                    "_section": None,
                    "Force (kN)": st.column_config.NumberColumn(
                        "⚡ Force (kN)",
                        help="Saisissez la force lue",
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
                    f"📈 **Résistance moyenne du lot : {fc_moy:.1f} MPa**"
                )

            # --- PRÉPARATION DES DONNÉES DU PV ---
            export_data = []
            for _, row in df_actuel.iterrows():
                export_data.append({
                    "repere_eprouvette": row["Repère"],
                    "forme": row["Forme d'éprouvette"],
                    "section": row["_section"],
                    "force_kn": row["Force (kN)"],
                    "fc_mpa": row["Résistance Fc (MPa)"],
                    "date_essai": sample.get("date_ecrasement", "N/A"),
                    "age": sample.get("echeance", "28")
                    .replace(" jours", "")
                    .replace("j", ""),
                })

            infos_header = {
                "re_num": "25/260/LGV/ B/01",
                "dossier": "2025-260-05985-2025-0247",
                "client": "TGCC",
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
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

            # TRAITEMENT DE LA SAUVEGARDE (SANS ST.RERUN SOUDAIN)
            if btn_enregistrer:
                if (df_actuel["Force (kN)"] == 0).any():
                    st.error(
                        "❌ Une ou plusieurs forces d'écrasement sont à 0.0"
                        " kN. Veuillez saisir toutes les forces."
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
                            f"✅ Lot de {succes_lot} éprouvettes enregistré dans"
                            " Supabase ! Vous pouvez télécharger le PV"
                            " ci-dessus ou actualiser la page."
                        )
                        "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"
                    ),
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
                        round((f_kn * 10.0) / sec, 1)
                        if sec > 0 and f_kn > 0
                        else 0.0
                    )

                    rows_list.append({
                        "ID": ep["id"],
                        "Repère": ep.get(
                            "repere_eprouvette", f"/{ep['id']}"
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
                    "Repère": st.column_config.TextColumn(
                        "Repère", disabled=True
                    ),
                    "Forme d'éprouvette": st.column_config.TextColumn(
                        "Forme d'éprouvette", disabled=True
                    ),
                    "_section": None,
                    "Force (kN)": st.column_config.NumberColumn(
                        "⚡ Force (kN)",
                        help="Saisissez la force lue",
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
                    f"📈 **Résistance moyenne du lot : {fc_moy:.1f} MPa**"
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
                            "❌ Une ou plusieurs forces d'écrasement sont à 0.0"
                            " kN."
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
                                f"✅ Lot de {succes_lot} éprouvettes enregistré"
                                " !"
                            )
                            st.rerun()

            with col_b2:
                export_data = []
                for _, row in df_actuel.iterrows():
                    export_data.append({
                        "repere_eprouvette": row["Repère"],
                        "forme": row["Forme d'éprouvette"],
                        "section": row["_section"],
                        "force_kn": row["Force (kN)"],
                        "fc_mpa": row["Résistance Fc (MPa)"],
                        "date_essai": sample.get("date_ecrasement", "N/A"),
                        "age": sample.get("echeance", "28")
                        .replace(" jours", "")
                        .replace("j", ""),
                    })

                infos_header = {
                    "re_num": "25/260/LGV/ B/01",
                    "dossier": "2025-260-05985-2025-0247",
                    "client": "TGCC",
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

                st.download_button(
                    label="📄 Télécharger le PV d'écrasement (Format LPEE)",
                    data=excel_file,
                    file_name=filename,
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

    # ---------------------------------------------------------
    # HISTORIQUE
    # ---------------------------------------------------------
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
                st.dataframe(df_all, use_container_width=True, hide_index=True)
            else:
                st.info("Aucun enregistrement d'écrasement dans la base.")
        except Exception as e:
            st.error(f"Erreur lors du chargement : {e}")
