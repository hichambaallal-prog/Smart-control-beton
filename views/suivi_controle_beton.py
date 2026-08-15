import io
from datetime import date, datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st


# =========================================================
# 1. GÉNÉRATION DU PROCÈS-VERBAL EXCEL (FORMAT EXACT LPEE)
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
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # Bordures
    thin_side = Side(border_style="thin", color="000000")
    border_cell = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )

    # Extraction sécurisée du N° BL pour remplacer les éventuelles valeurs "N/A"
    default_bl = str(infos_header.get("num_bl") or "N/A")

    def remplacer_na(valeur):
        if str(valeur).strip().upper() in ["N/A", "NONE", "NAN", ""]:
            return default_bl
        return valeur

    # --- ENTÊTE DU LABORATOIRE ET RÉFÉRENCES ---
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
    ws["F1"] = remplacer_na(infos_header.get("re_num", "25/260/LGV/ B/01"))
    ws["F1"].font = font_regular

    ws["E2"] = "DOSSIER :"
    ws["E2"].font = font_bold
    ws.merge_cells("F2:H2")
    ws["F2"] = remplacer_na(infos_header.get("dossier", "2025-260-05985-2025-0247"))
    ws["F2"].font = font_regular

    ws["E3"] = "CLIENT :"
    ws["E3"].font = font_bold
    ws.merge_cells("F3:H3")
    ws["F3"] = remplacer_na(infos_header.get("client", "TGCC"))
    ws["F3"].font = font_bold

    for r in range(1, 4):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # --- TITRE DE L'ESSAI ET TYPE D'ESSAI ---
    ws.merge_cells("A4:H4")
    ws["A4"] = "ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
    ws["A4"].font = Font(name="Calibri", size=11, bold=True)
    ws["A4"].alignment = align_center
    ws["A4"].border = border_cell

    ws.merge_cells("A5:D5")
    ws["A5"] = "[X] COMPRESSION NF EN 12390-3 (2019)"
    ws["A5"].font = font_bold
    ws["A5"].alignment = align_center

    ws.merge_cells("E5:H5")
    ws["E5"] = "[ ] TRACTION PAR FENDAGE NF EN 12390-6 (2019)"
    ws["E5"].font = font_bold
    ws["E5"].alignment = align_center

    for c in range(1, 9):
        ws.cell(row=5, column=c).border = border_cell

    # MODIFICATION : Ligne 6 - Alignement à droite pour "Presse : Marque: Controls"
    ws.merge_cells("A6:F6")
    ws["A6"] = "Presse : Marque: Controls"
    ws["A6"].font = font_bold
    ws["A6"].alignment = align_right

    ws.merge_cells("G6:H6")
    ws["G6"] = "Classe : A"
    ws["G6"].font = font_bold
    ws["G6"].alignment = align_center

    for c in range(1, 9):
        ws.cell(row=6, column=c).border = border_cell

    # --- FICHE TECHNIQUE DE PRÉLÈVEMENT ET CHANTIER ---
    ws["A7"] = "Date de\nprélèvement"
    ws["A7"].font = font_bold
    ws["A7"].alignment = align_center
    ws["B7"] = str(remplacer_na(infos_header.get("date_coulee")))
    ws["B7"].font = font_bold
    ws["B7"].alignment = align_center

    ws.merge_cells("C7:D7")
    ws["C7"] = "Lieu de\nprélèvement"
    ws["C7"].font = font_bold
    ws["C7"].alignment = align_center

    ws.merge_cells("E7:H7")
    ws["E7"] = remplacer_na(
        infos_header.get("lieu_prelevement", infos_header.get("ouvrage"))
    )
    ws["E7"].font = font_regular
    ws["E7"].alignment = align_center

    # MODIFICATION : Hauteur de la ligne 8 à 54 (Suppression de la ligne 9)
    ws["A8"] = "Chantier"
    ws["A8"].font = font_bold
    ws["A8"].alignment = align_center

    ws.merge_cells("B8:D8")
    ws["B8"] = remplacer_na(
        infos_header.get(
            "chantier",
            "Augmentation de la capacité ferroviaire entre Kenitra et Marrakech et au niveau du hub de Casablanca\nTravaux d'exécution de terrassement, ouvrages d'art et rétablissement de communication entre PK 5+450 et PK 10+000-GARE CASA SUD",
        )
    )
    ws["B8"].font = font_small
    ws["B8"].alignment = align_center

    ws.merge_cells("E8:F8")
    ws["E8"] = "Type de béton"
    ws["E8"].font = font_bold
    ws["E8"].alignment = align_center

    ws.merge_cells("G8:H8")
    ws["G8"] = remplacer_na(infos_header.get("classe_beton", "C30/37"))
    ws["G8"].font = font_bold
    ws["G8"].alignment = align_center

    # MODIFICATION : Remplacement de "TG Prefa Oulad Saleh" par la saisie dynamique de la Centrale
    centrale_saisie = remplacer_na(infos_header.get("centrale", "Centrale à Béton"))
    ws.merge_cells("A9:B9")
    ws["A9"] = centrale_saisie
    ws["A9"].font = font_bold
    ws["A9"].alignment = align_center

    ws["C9"] = "- Dimensions"
    ws["C9"].font = font_regular

    # Inscription dynamique du Type / Forme de l'éprouvette
    ws.merge_cells("D9:H9")
    ws["D9"] = remplacer_na(infos_header.get("forme", "Cylindrique 150x300"))
    ws["D9"].font = font_bold
    ws["D9"].alignment = align_center

    ws.merge_cells("A10:B10")
    ws["A10"] = "Affaissement au cône d'abrams NF EN 12350-2"
    ws["A10"].font = font_small
    ws["A10"].alignment = align_center

    ws["C10"] = str(remplacer_na(infos_header.get("affaissement")))
    ws["C10"].font = font_bold
    ws["C10"].alignment = align_center

    ws["D10"] = "- Mode confection"
    ws["D10"].font = font_regular
    ws["D10"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("E10:H10")
    ws["E10"] = "Par vibration NF EN 12390-2 (2019)"
    ws["E10"].font = font_bold
    ws["E10"].alignment = align_center

    ws.merge_cells("A11:B11")
    ws["A11"] = "Température °C"
    ws["A11"].font = font_regular
    ws["A11"].alignment = align_center

    ws["C11"] = str(remplacer_na(infos_header.get("temperature")))
    ws["C11"].font = font_bold
    ws["C11"].alignment = align_center

    ws["D11"] = "- Mode conservation"
    ws["D11"].font = font_regular
    ws["D11"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("E11:H11")
    ws["E11"] = "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C ± 2°C"
    ws["E11"].font = font_bold
    ws["E11"].alignment = align_center

    ws.merge_cells("A12:C12")
    ws["A12"] = "Densité du béton durci NF EN 12390-7(2019)"
    ws["A12"].font = font_small
    ws["A12"].alignment = align_center

    ws.merge_cells("D12:E12")
    ws["D12"] = "N° de bon de livraison"
    ws["D12"].font = font_regular
    ws["D12"].alignment = align_center

    ws.merge_cells("F12:H12")
    ws["F12"] = default_bl
    ws["F12"].font = font_bold
    ws["F12"].alignment = align_center

    for r in range(7, 13):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # --- TABLEAU DES RÉSULTATS D'ÉCRASEMENT ---
    ws.merge_cells("A13:A14")
    ws["A13"] = "Réf,"
    ws["A13"].font = font_bold
    ws["A13"].alignment = align_center

    ws.merge_cells("B13:C13")
    ws["B13"] = "Date"
    ws["B13"].font = font_bold
    ws["B13"].alignment = align_center

    ws["B14"] = "Fabri"
    ws["B14"].font = font_regular
    ws["B14"].alignment = align_center

    ws["C14"] = "Essai"
    ws["C14"].font = font_regular
    ws["C14"].alignment = align_center

    ws.merge_cells("D13:D14")
    ws["D13"] = "Age (jours)"
    ws["D13"].font = font_bold
    ws["D13"].alignment = align_center

    ws.merge_cells("E13:E14")
    ws["E13"] = "Charge rupture(KN)"
    ws["E13"].font = font_bold
    ws["E13"].alignment = align_center

    ws.merge_cells("F13:H13")
    ws["F13"] = "Résistance (MPa)"
    ws["F13"].font = font_bold
    ws["F13"].alignment = align_center

    ws["F14"] = "Compression"
    ws["F14"].font = font_regular
    ws["F14"].alignment = align_center

    ws["G14"] = "Traction"
    ws["G14"].font = font_regular
    ws["G14"].alignment = align_center

    ws["H14"] = "Moyenne"
    ws["H14"].font = font_regular
    ws["H14"].alignment = align_center

    for r in range(13, 15):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    row_start = 15
    nb_total = len(export_data)

    if nb_total > 0:
        ws.merge_cells(f"A{row_start}:A{row_start + nb_total - 1}")
        ws[f"A{row_start}"] = "B/01"
        ws[f"A{row_start}"].font = font_bold
        ws[f"A{row_start}"].alignment = align_center

        ws.merge_cells(f"B{row_start}:B{row_start + nb_total - 1}")
        ws[f"B{row_start}"] = str(remplacer_na(infos_header.get("date_coulee")))
        ws[f"B{row_start}"].font = font_bold
        ws[f"B{row_start}"].alignment = align_center

    for idx, item in enumerate(export_data):
        curr_row = row_start + idx

        ws.cell(
            row=curr_row,
            column=3,
            value=str(remplacer_na(item.get("date_essai"))),
        ).alignment = align_center

        ws.cell(
            row=curr_row, column=4, value=item.get("age", 7)
        ).alignment = align_center

        f_kn = float(item.get("force_kn", 0.0))
        ws.cell(
            row=curr_row, column=5, value=f"{f_kn:.1f}".replace(".", ",")
        ).alignment = align_right

        fc_mpa = float(item.get("fc_mpa", 0.0))
        ws.cell(
            row=curr_row, column=6, value=f"{fc_mpa:.1f}".replace(".", ",")
        ).alignment = align_right

        ws.cell(row=curr_row, column=7, value="-").alignment = align_center

        for c in range(1, 9):
            ws.cell(row=curr_row, column=c).font = font_regular
            ws.cell(row=curr_row, column=c).border = border_cell

    # MODIFICATION : Remplacement de la case moyenne par la formule dynamique de moyenne du lot (colonne H)
    row_end = row_start + max(nb_total, 1) - 1
    if nb_total > 0:
        ws.merge_cells(f"H{row_start}:H{row_end}")
        ws[f"H{row_start}"] = f"=ROUND(AVERAGE(F{row_start}:F{row_end}), 1)"
        ws[f"H{row_start}"].alignment = align_center
        ws[f"H{row_start}"].font = font_bold

    # MODIFICATION : Gestion de la ligne 22 pour le commentaire de conformité
    # Suppression de l'ancienne ligne 22 et insertion d'une ligne 22 propre
    ws.delete_rows(22)
    ws.insert_rows(22)

    ws.cell(row=22, column=1, value="Commentaire :").font = font_bold
    ws.cell(row=22, column=1).alignment = align_left

    ws.merge_cells("B22:H22")
    
    # Formule dynamique de conformité en fonction de la classe de béton et de la moyenne à 28 jours (Cellule H15/H_moyenne)
    # Vérifie si le béton respecte le seuil MPa du projet
    moyenne_cell = f"H{row_start}"
    formule_commentaires = (
        f'=IF(ISBLANK({moyenne_cell}), "", '
        f'IF(OR('
            f'AND(ISNUMBER(SEARCH("C25/30", G8)), {moyenne_cell}>=25), '
            f'AND(ISNUMBER(SEARCH("C30/37", G8)), {moyenne_cell}>=30), '
            f'AND(ISNUMBER(SEARCH("C35/45", G8)), {moyenne_cell}>=35), '
            f'AND(ISNUMBER(SEARCH("C40/50", G8)), {moyenne_cell}>=40)'
        f'), '
        f'"PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES", '
        f'"PERFORMANCES MECANIQUES A 28 JOURS NE SONT PAS CONFORMES"))'
    )
    
    ws.cell(row=22, column=2, value=formule_commentaires).font = font_bold
    ws.cell(row=22, column=2).alignment = align_left

    for c in range(1, 9):
        ws.cell(row=22, column=c).border = border_cell

    # MODIFICATION : Configuration des hauteurs de lignes (Hauteur ligne 8 = 54)
    ws.row_dimensions[8].height = 54
    for r in range(1, 23):
        if r != 8:
            if 1 <= r <= 6:
                ws.row_dimensions[r].height = 16
            else:
                ws.row_dimensions[r].height = 31

    # Largeurs des colonnes
    col_widths = {
        "A": 10,
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
# FONCTION AUXILIAIRE : HISTORIQUE D'ÉCRASEMENT D'UN BÉTON
# =========================================================
def obtenir_historique_betonnage(supabase, betonnage_id):
    """Récupère l'intégralité des écrasements déjà enregistrés pour un même béton (betonnage_id)."""
    if not betonnage_id:
        return []
    try:
        res = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .eq("betonnage_id", betonnage_id)
            .not_.is_("force_kn", "null")
            .gt("force_kn", 0)
            .order("id", desc=False)
            .execute()
        )
        return res.data if res.data else []
    except Exception as e:
        st.warning(
            f"Note : Historique du bétonnage #{betonnage_id} non disponible : {e}"
        )
        return []


def obtenir_infos_betonnage_parent(supabase, betonnage_id):
    """Récupère les détails saisis initialement au niveau de la table suivi_betonnage."""
    if not betonnage_id:
        return {}
    try:
        res = (
            supabase.table("suivi_betonnage")
            .select("*")
            .eq("id", betonnage_id)
            .execute()
        )
        if res.data:
            return res.data[0]
    except Exception as e:
        st.warning(
            f"Note : Impossible de charger la fiche parent de bétonnage #{betonnage_id} : {e}"
        )
    return {}


# =========================================================
# 2. APPLICATION PRINCIPALE STREAMLIT
# =========================================================
def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    tab_prog, tab_saisie, tab_hist = st.tabs([
        "📅 Phase 1 : Programmation",
        "💥 Phase 2 : Saisie des Écrasements (Par Lot)",
        "📋 Historique Complet & PVs",
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

    # ---------------------------------------------------------
    # PHASE 1 : PROGRAMMATION
    # ---------------------------------------------------------
    with tab_prog:
        st.subheader("📅 1. Programmer les Échéances d'Écrasement")

        if not betonnages_preleves:
            st.info(
                "ℹ️ Aucun suivi de bétonnage avec prélèvement d'éprouvettes"
                " (OUI) trouvé."
            )
        else:
            options_beton = {
                (
                    f"ID #{b['id']} | BL: {b.get('num_bl', 'N/A')} | Ouvrage:"
                    f" {b.get('ouvrage', 'N/A')} | Date:"
                    f" {b.get('date_coulee', b.get('date_livraison', 'N/A'))} |"
                    f" Classe: {b.get('classe_beton', b.get('classe', 'N/A'))}"
                ): b
                for b in betonnages_preleves
            }

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
                                f"Erreur lors de la programmation de {rep} :"
                                f" {err}"
                            )

                    if succes_cnt > 0:
                        st.success(
                            f"✅ {succes_cnt} éprouvette(s) programmée(s) pour"
                            f" le {date_ecrasement_prevue} ({echeance_p}) !"
                        )
                        st.rerun()

    # ---------------------------------------------------------
    # PHASE 2 : SAISIE DES ÉCRASEMENTS PAR LOT
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
                    f" | Lot ID #{b_id_ep}"
                )

                if cle_groupe not in groupes_lots:
                    groupes_lots[cle_groupe] = []
                groupes_lots[cle_groupe].append(ep)

            choix_lot = st.selectbox(
                "📦 Sélectionner le lot d'éprouvettes à écraser :",
                list(groupes_lots.keys()),
                key="select_lot_saisie",
            )
            lot_selected = groupes_lots[choix_lot]

            sample = lot_selected[0]
            betonnage_id = sample.get("betonnage_id")

            # Récupération des informations initiales saisies au niveau du suivi de bétonnage
            info_betonnage = obtenir_infos_betonnage_parent(
                supabase, betonnage_id
            )

            # Récupération automatique des essais déjà effectués (3j, 7j)
            essais_anterieurs = obtenir_historique_betonnage(
                supabase, betonnage_id
            )

            col_l1, col_l2, col_l3, col_l4 = st.columns(4)
            col_l1.metric("Client", "TGCC")
            col_l2.metric("N° Bon Livraison", str(sample.get("num_bl", "N/A")))
            col_l3.metric("Ouvrage", str(sample.get("ouvrage", "N/A")))
            col_l4.metric("Échéance Visée", str(sample.get("echeance", "N/A")))

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

            st.markdown("##### 📝 Saisie des forces d'écrasement")

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
                        "Repère": ep.get("repere_eprouvette", f"/{ep['id']}"),
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
                        st.session_state[lot_key].at[
                            row_idx, "Force (kN)"
                        ] = new_force
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
                        help="Saisissez la force lue sur la presse",
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
                    f"📈 **Résistance moyenne du lot actuel : {fc_moy:.1f} MPa**"
                )

            # Construction des données complètes (Historique + Lot Actuel)
            export_data = []

            # 1. Éprouvettes antérieures du même lot
            if essais_anterieurs:
                st.info(
                    f"ℹ️ {len(essais_anterieurs)} essai(s) antérieur(s)"
                    f" répertorié(s) pour ce béton (Bétonnage ID #{betonnage_id})"
                    " et inclus dans l'impression."
                )
                for ep_ant in essais_anterieurs:
                    sec_a = float(ep_ant.get("section") or 176.71)
                    f_a = float(ep_ant.get("force_kn") or 0.0)
                    fc_a = float(
                        ep_ant.get("fc_mpa") or round((f_a * 10.0) / sec_a, 1)
                    )

                    export_data.append({
                        "repere_eprouvette": ep_ant.get(
                            "repere_eprouvette", "N/A"
                        ),
                        "forme": ep_ant.get("forme", "Cylindrique 150x300"),
                        "section": sec_a,
                        "force_kn": f_a,
                        "fc_mpa": fc_a,
                        "date_essai": ep_ant.get("date_ecrasement", "N/A"),
                        "age": str(ep_ant.get("echeance", "7"))
                        .replace(" jours", "")
                        .replace("j", ""),
                    })

            # 2. Éprouvettes de la saisie actuelle
            for _, row in df_actuel.iterrows():
                export_data.append({
                    "repere_eprouvette": row["Repère"],
                    "forme": row["Forme d'éprouvette"],
                    "section": row["_section"],
                    "force_kn": row["Force (kN)"],
                    "fc_mpa": row["Résistance Fc (MPa)"],
                    "date_essai": sample.get("date_ecrasement", "N/A"),
                    "age": str(sample.get("echeance", "28"))
                    .replace(" jours", "")
                    .replace("j", ""),
                })

            # Extraction dynamique des variables depuis suivi_betonnage et fallback
            num_bl_valeur = info_betonnage.get("num_bl") or sample.get("num_bl")
            affaissement_saisi = (
                info_betonnage.get("affaissement")
                or info_betonnage.get("slump")
                or sample.get("affaissement")
            )
            temp_saisie = (
                info_betonnage.get("temperature")
                or info_betonnage.get("temp_beton")
                or sample.get("temperature")
            )
            ouvrage_saisi = (
                info_betonnage.get("ouvrage")
                or sample.get("ouvrage")
            )
            date_coulee_saisie = (
                info_betonnage.get("date_coulee")
                or sample.get("date_coulee")
            )
            centrale_saisie = (
                info_betonnage.get("centrale")
                or info_betonnage.get("centrale_beton")
                or sample.get("centrale")
            )

            infos_header = {
                "re_num": "25/260/LGV/ B/01",
                "dossier": "2025-260-05985-2025-0247",
                "client": "TGCC",
                "num_bl": num_bl_valeur,
                "ouvrage": ouvrage_saisi,
                "lieu_prelevement": ouvrage_saisi,
                "classe_beton": sample.get("classe_beton", "C30/37"),
                "date_coulee": date_coulee_saisie,
                "affaissement": affaissement_saisi,
                "temperature": temp_saisie,
                "forme": sample.get("forme", "Cylindrique 150x300"),
                "centrale": centrale_saisie,
                "observations": obs_globale,
            }

            excel_file = generer_pv_excel(export_data, infos_header)
            filename = f"PV_Ecrasement_LPEE_{num_bl_valeur or 'BL'}.xlsx"

            st.markdown("---")
            col_b1, col_b2 = st.columns(2)

            with col_b1:
                btn_enregistrer = st.button(
                    "💾 Valider et Enregistrer Le Lot",
                    type="primary",
                    use_container_width=True,
                )

            with col_b2:
                st.download_button(
                    label="📄 Télécharger le PV (Excel Modèle LPEE)",
                    data=excel_file,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            if btn_enregistrer:
                if (df_actuel["Force (kN)"] == 0).any():
                    st.error(
                        "❌ Les forces de rupture doivent toutes être saisies"
                        " (> 0 kN)."
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
                            f"✅ Lot de {succes_lot} éprouvettes validé dans"
                            " Supabase ! Vous pouvez télécharger le PV"
                            " ci-dessus."
                        )

    # ---------------------------------------------------------
    # HISTORIQUE COMPLET & ÉDITION DE PV
    # ---------------------------------------------------------
    with tab_hist:
        st.subheader("📋 Historique Général & Consultation des PVs")
        try:
            res_all = (
                supabase.table("suivi_controle_beton")
                .select("*")
                .order("id", desc=True)
                .execute()
            )
            if res_all.data:
                df_all = pd.DataFrame(res_all.data)

                # Extraction des enregistrements validés
                df_valides = df_all[
                    (df_all["force_kn"].notnull()) & (df_all["force_kn"] > 0)
                ].copy()

                if not df_valides.empty:
                    st.markdown("##### 📥 Re-télécharger un Procès-Verbal")

                    groupes_valides = {}
                    for _, row in df_valides.iterrows():
                        b_id_ep = row.get("betonnage_id")
                        ech_ep = row.get("echeance", "28 jours")
                        ouv_ep = row.get("ouvrage", "N/A")
                        dt_ecras = row.get("date_ecrasement", "N/A")

                        cle_pv = (
                            f"Ouvrage: {ouv_ep} | Échéance: {ech_ep} (Date:"
                            f" {dt_ecras}) | Lot ID #{b_id_ep}"
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
                    b_id_h = sample_h.get("betonnage_id")

                    info_beton_h = obtenir_infos_betonnage_parent(
                        supabase, b_id_h
                    )
                    tous_essais_hist = obtenir_historique_betonnage(
                        supabase, b_id_h
                    )

                    export_data_h = []
                    items_a_exporter = (
                        tous_essais_hist if tous_essais_hist else lot_hist
                    )

                    for item in items_a_exporter:
                        sec = float(item.get("section") or 176.71)
                        f_kn = float(item.get("force_kn") or 0.0)
                        fc = float(
                            item.get("fc_mpa")
                            or round((f_kn * 10.0) / sec, 1)
                        )

                        export_data_h.append({
                            "repere_eprouvette": item.get(
                                "repere_eprouvette", f"/{item['id']}"
                            ),
                            "forme": item.get("forme", "Cylindrique 150x300"),
                            "section": sec,
                            "force_kn": f_kn,
                            "fc_mpa": fc,
                            "date_essai": item.get("date_ecrasement", "N/A"),
                            "age": str(item.get("echeance", "28"))
                            .replace(" jours", "")
                            .replace("j", ""),
                        })

                    num_bl_h = info_beton_h.get("num_bl") or sample_h.get("num_bl")
                    aff_h = (
                        info_beton_h.get("affaissement")
                        or info_beton_h.get("slump")
                        or sample_h.get("affaissement")
                    )
                    temp_h = (
                        info_beton_h.get("temperature")
                        or info_beton_h.get("temp_beton")
                        or sample_h.get("temperature")
                    )
                    ouv_h = (
                        info_beton_h.get("ouvrage")
                        or sample_h.get("ouvrage")
                    )
                    date_coulee_h = (
                        info_beton_h.get("date_coulee")
                        or sample_h.get("date_coulee")
                    )
                    centrale_h = (
                        info_beton_h.get("centrale")
                        or info_beton_h.get("centrale_beton")
                        or sample_h.get("centrale")
                    )

                    infos_header_h = {
                        "re_num": "25/260/LGV/ B/01",
                        "dossier": "2025-260-05985-2025-0247",
                        "client": "TGCC",
                        "num_bl": num_bl_h,
                        "ouvrage": ouv_h,
                        "lieu_prelevement": ouv_h,
                        "classe_beton": sample_h.get("classe_beton", "C30/37"),
                        "date_coulee": date_coulee_h,
                        "affaissement": aff_h,
                        "temperature": temp_h,
                        "forme": sample_h.get("forme", "Cylindrique 150x300"),
                        "centrale": centrale_h,
                        "observations": sample_h.get(
                            "observations",
                            "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES",
                        ),
                    }

                    excel_pv_hist = generer_pv_excel(
                        export_data_h, infos_header_h
                    )
                    file_name_h = f"PV_Ecrasement_RE-EXPORT_{num_bl_h or 'BL'}.xlsx"

                    st.download_button(
                        label="📄 Télécharger le PV (Excel Format LPEE)",
                        data=excel_pv_hist,
                        file_name=file_name_h,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_download_hist",
                    )

                st.markdown("---")
                st.markdown("##### 📊 Base de données globale")
                st.dataframe(df_all, use_container_width=True, hide_index=True)
            else:
                st.info("Aucun enregistrement d'écrasement dans la base.")
        except Exception as e:
            st.error(f"Erreur lors du chargement de l'historique : {e}")
