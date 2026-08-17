import io
import re
from datetime import date, datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st


# =========================================================
# FONCTION UTILITAIRE : EXTRACTION SÉCURISÉE DU N° BL
# =========================================================
def extraire_num_bl(*sources):
    """Inspecte récursivement les sources pour extraire le N° de Bon de Livraison (BL)."""
    clefs_possibles = [
        "num_bl",
        "bl",
        "num_bon_livraison",
        "n_bl",
        "bon_livraison",
        "num_bl_p",
        "n_bon",
        "bon_de_livraison",
        "code_bl",
    ]

    for source in sources:
        if isinstance(source, dict):
            for key in clefs_possibles:
                val = source.get(key)
                if val is not None:
                    val_str = str(val).strip()
                    if val_str and val_str.upper() not in [
                        "N/A",
                        "NONE",
                        "NAN",
                        "-",
                        "",
                    ]:
                        return val_str

            for key, val in source.items():
                if "bl" in key.lower() or "bon" in key.lower():
                    if val is not None:
                        val_str = str(val).strip()
                        if val_str and val_str.upper() not in [
                            "N/A",
                            "NONE",
                            "NAN",
                            "-",
                            "",
                        ]:
                            return val_str

        elif isinstance(source, str):
            match = re.search(r"BL\s*:\s*([^\|]+)", source, re.IGNORECASE)
            if match:
                val_str = match.group(1).strip()
                if val_str and val_str.upper() not in [
                    "N/A",
                    "NONE",
                    "NAN",
                    "-",
                    "",
                ]:
                    return val_str

    return "-"


# =========================================================
# 1. GÉNÉRATION DU PROCÈS-VERBAL EXCEL (FORMAT EXACT LPEE)
# =========================================================
def generer_pv_excel(export_data, infos_header):
    """Génère un Procès-Verbal (PV) d'écrasement de béton répliquant le modèle LPEE."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PV Écrasement LPEE"

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.page_margins = PageMargins(
        left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.2, footer=0.2
    )

    font_bold = Font(name="Calibri", size=9, bold=True)
    font_bold_white = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_title_white = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=8.5)
    font_small = Font(name="Calibri", size=8)

    fill_header_dark = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    fill_header_table = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    fill_section_label = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)

    thin_side = Side(border_style="thin", color="000000")
    border_cell = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )

    default_bl = extraire_num_bl(infos_header)

    def remplacer_na(valeur, fallback=None):
        val_str = str(valeur).strip() if valeur is not None else ""
        if val_str.upper() in ["N/A", "NONE", "NAN", "", "-"]:
            return fallback if fallback is not None else default_bl
        return valeur

    # ENTÊTE
    ws.merge_cells("A1:D1")
    ws["A1"] = "LPEE / CTR CSB"
    ws["A1"].font = font_bold_white
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:D3")
    ws["A2"] = "Laboratoire de Contrôle Externe"
    ws["A2"].font = font_bold_white
    ws["A2"].alignment = align_center

    for r in range(1, 4):
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = fill_header_dark

    ws["E1"] = "RE N° :"
    ws["E1"].font = font_bold
    ws.merge_cells("F1:H1")
    ws["F1"] = remplacer_na(infos_header.get("re_num"), "25/260/LGV/ B/01")
    ws["F1"].font = font_regular

    ws["E2"] = "DOSSIER :"
    ws["E2"].font = font_bold
    ws.merge_cells("F2:H2")
    ws["F2"] = remplacer_na(infos_header.get("dossier"), "2025-260-05985-2025-0247")
    ws["F2"].font = font_regular

    ws["E3"] = "CLIENT :"
    ws["E3"].font = font_bold
    ws.merge_cells("F3:H3")
    ws["F3"] = remplacer_na(infos_header.get("client"), "TGCC")
    ws["F3"].font = font_bold

    for r in range(1, 4):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # TITRE
    ws.merge_cells("A4:H4")
    ws["A4"] = "ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
    ws["A4"].font = font_title_white
    ws["A4"].alignment = align_center
    for c in range(1, 9):
        ws.cell(row=4, column=c).fill = fill_header_dark
        ws.cell(row=4, column=c).border = border_cell

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

    # FICHE TECHNIQUE
    ws["A7"] = "Date de\nprélèvement"
    ws["A7"].font = font_bold
    ws["A7"].alignment = align_center
    ws["B7"] = str(remplacer_na(infos_header.get("date_coulee"), "-"))
    ws["B7"].font = font_bold
    ws["B7"].alignment = align_center

    ws.merge_cells("C7:D7")
    ws["C7"] = "Lieu de\nprélèvement"
    ws["C7"].font = font_bold
    ws["C7"].alignment = align_center

    ws.merge_cells("E7:H7")
    ws["E7"] = remplacer_na(
        infos_header.get("lieu_prelevement", infos_header.get("ouvrage")), "-"
    )
    ws["E7"].font = font_regular
    ws["E7"].alignment = align_center

    ws["A8"] = "Chantier"
    ws["A8"].font = font_bold
    ws["A8"].alignment = align_center

    ws.merge_cells("B8:D8")
    ws["B8"] = remplacer_na(
        infos_header.get("chantier"),
        "Augmentation de la capacité ferroviaire entre Kenitra et Marrakech et au niveau du hub de Casablanca\nTravaux d'exécution de terrassement, ouvrages d'art et rétablissement de communication entre PK 5+450 et PK 10+000-GARE CASA SUD",
    )
    ws["B8"].font = font_small
    ws["B8"].alignment = align_center

    ws.merge_cells("E8:F8")
    ws["E8"] = "Type de béton"
    ws["E8"].font = font_bold
    ws["E8"].alignment = align_center

    ws.merge_cells("G8:H8")
    classe_beton_val = str(remplacer_na(infos_header.get("classe_beton"), "C35/45")).upper()
    ws["G8"] = classe_beton_val
    ws["G8"].font = font_bold
    ws["G8"].alignment = align_center

    centrale_saisie = remplacer_na(infos_header.get("centrale"), "Centrale à Béton")
    ws.merge_cells("A9:B9")
    ws["A9"] = centrale_saisie
    ws["A9"].font = font_bold
    ws["A9"].alignment = align_center

    ws["C9"] = "- Dimensions"
    ws["C9"].font = font_regular

    ws.merge_cells("D9:H9")
    ws["D9"] = remplacer_na(infos_header.get("forme"), "Cylindrique 150x300")
    ws["D9"].font = font_bold
    ws["D9"].alignment = align_center

    ws.merge_cells("A10:B10")
    ws["A10"] = "Affaissement au cône d'abrams NF EN 12350-2"
    ws["A10"].font = font_small
    ws["A10"].alignment = align_center

    ws["C10"] = str(remplacer_na(infos_header.get("affaissement"), "-"))
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

    ws["C11"] = str(remplacer_na(infos_header.get("temperature"), "-"))
    ws["C11"].font = font_bold
    ws["C11"].alignment = align_center

    ws["D11"] = "- Mode conservation"
    ws["D11"].font = font_regular
    ws["D11"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("E11:H11")
    ws["E11"] = "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C ± 2°C"
    ws["E11"].font = font_bold
    ws["E11"].alignment = align_center

    tech_prelevement = remplacer_na(
        infos_header.get("technicien_prelevement") 
        or infos_header.get("preleve_par") 
        or infos_header.get("technicien"), 
        "Technicien LPEE"
    )
    ws.merge_cells("A12:C12")
    ws["A12"] = f"prélèvement effectué par {tech_prelevement}"
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

    labels_coords = ["A7", "C7", "A8", "E8", "A9", "A10", "A11", "A12", "D12"]
    for r in range(7, 13):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell
            if ws.cell(row=r, column=c).coordinate in labels_coords:
                ws.cell(row=r, column=c).fill = fill_section_label

    # TABLEAU DES RÉSULTATS
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

    ws.merge_cells("G14:G14")
    ws["G14"] = "Traction"
    ws["G14"].font = font_regular
    ws["G14"].alignment = align_center

    ws.merge_cells("H14:H14")
    ws["H14"] = "Moyenne"
    ws["H14"].font = font_regular
    ws["H14"].alignment = align_center

    for r in range(13, 15):
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = fill_header_table
            ws.cell(row=r, column=c).border = border_cell

    row_start = 15
    nb_total = len(export_data)

    groupes_lots = {}
    a_des_28j_ecrases = False
    cellule_moyenne_28j = None

    for idx, item in enumerate(export_data):
        curr_row = row_start + idx

        ref_complete = str(item.get("repere_eprouvette", "B/01"))
        ws.cell(row=curr_row, column=1, value=ref_complete)

        ws.cell(
            row=curr_row,
            column=2,
            value=str(remplacer_na(infos_header.get("date_coulee"), "-")),
        )

        dt_essai = item.get("date_essai")

        try:
            f_kn_val = float(item.get("force_kn", 0.0))
        except (ValueError, TypeError):
            f_kn_val = 0.0

        is_en_cours = str(item.get("statut", "")).lower() == "en cours" or f_kn_val == 0.0

        if is_en_cours:
            ws.cell(row=curr_row, column=3, value=str(dt_essai) if dt_essai and dt_essai != "-" else "En cours")
        else:
            ws.cell(row=curr_row, column=3, value=str(remplacer_na(dt_essai, "-")))

        try:
            age_val = int(str(item.get("age", 7)).replace("j", "").replace("jours", "").strip())
        except (ValueError, TypeError):
            age_val = item.get("age", 7)

        ws.cell(row=curr_row, column=4, value=age_val)

        if is_en_cours:
            ws.cell(row=curr_row, column=5, value="En cours")
            ws.cell(row=curr_row, column=6, value="En cours")
        else:
            ws.cell(row=curr_row, column=5, value=f_kn_val)
            ws.cell(row=curr_row, column=5).number_format = "0.0"

            fc_mpa = float(item.get("fc_mpa", 0.0))
            ws.cell(row=curr_row, column=6, value=fc_mpa)
            ws.cell(row=curr_row, column=6).number_format = "0.0"

        ws.cell(row=curr_row, column=7, value="-")

        for c in range(1, 9):
            ws.cell(row=curr_row, column=c).font = font_regular
            ws.cell(row=curr_row, column=c).border = border_cell
            ws.cell(row=curr_row, column=c).alignment = align_center

        cle_lot = f"{item.get('age')}_{item.get('date_essai')}"
        if cle_lot not in groupes_lots:
            groupes_lots[cle_lot] = {"lignes": [], "en_cours": is_en_cours, "age": age_val}
        else:
            if is_en_cours:
                groupes_lots[cle_lot]["en_cours"] = True

        groupes_lots[cle_lot]["lignes"].append(curr_row)

    for cle_lot, data_lot in groupes_lots.items():
        lignes = data_lot["lignes"]
        start_r = min(lignes)
        end_r = max(lignes)
        age_lot = data_lot["age"]

        if data_lot["en_cours"]:
            if start_r != end_r:
                ws.merge_cells(f"H{start_r}:H{end_r}")
            ws[f"H{start_r}"] = "En cours"
            ws[f"H{start_r}"].alignment = align_center
            ws[f"H{start_r}"].font = font_bold
        else:
            if start_r == end_r:
                ws[f"H{start_r}"] = f"=ROUND(F{start_r}, 1)"
            else:
                ws.merge_cells(f"H{start_r}:H{end_r}")
                ws[f"H{start_r}"] = f"=ROUND(AVERAGE(F{start_r}:F{end_r}), 1)"

            ws[f"H{start_r}"].number_format = "0.0"
            ws[f"H{start_r}"].alignment = align_center
            ws[f"H{start_r}"].font = font_bold

            try:
                if int(str(age_lot).replace("j", "").replace("jours", "").strip()) >= 28:
                    a_des_28j_ecrases = True
                    cellule_moyenne_28j = f"H{start_r}"
            except (ValueError, TypeError):
                pass

    next_row = row_start + nb_total

    ws.cell(row=next_row, column=1, value="Commentaire :").font = font_bold
    ws.cell(row=next_row, column=1).alignment = align_left
    ws.cell(row=next_row, column=1).fill = fill_section_label

    ws.merge_cells(f"B{next_row}:H{next_row}")

    obs_defaut = infos_header.get("observations") or "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"

    seuil_min = 35.0
    if "C25/30" in classe_beton_val:
        seuil_min = 25.0
    elif "C30/37" in classe_beton_val:
        seuil_min = 30.0
    elif "C35/45" in classe_beton_val:
        seuil_min = 35.0
    elif "C40/50" in classe_beton_val:
        seuil_min = 40.0

    if not a_des_28j_ecrases or not cellule_moyenne_28j:
        formule_commentaires = "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT."
    else:
        moyenne_cell = cellule_moyenne_28j
        formule_commentaires = (
            f'=IF(OR(ISBLANK({moyenne_cell}), {moyenne_cell}="En cours"), '
            f'"PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT.", '
            f'IF({moyenne_cell}>={seuil_min}, "{obs_defaut}", "PERFORMANCES MECANIQUES NON CONFORMES"))'
        )

    ws.cell(row=next_row, column=2, value=formule_commentaires).font = font_bold
    ws.cell(row=next_row, column=2).alignment = align_left

    for c in range(1, 9):
        ws.cell(row=next_row, column=c).border = border_cell

    r_sig_titre = next_row + 2
    r_sig_debut = r_sig_titre + 1
    r_sig_fin = r_sig_debut + 3

    ws.merge_cells(start_row=r_sig_titre, start_column=2, end_row=r_sig_titre, end_column=4)
    ws.cell(row=r_sig_titre, column=2, value="Visa Responsable d'essai").font = font_bold
    ws.cell(row=r_sig_titre, column=2).alignment = align_center

    ws.merge_cells(start_row=r_sig_debut, start_column=2, end_row=r_sig_fin, end_column=4)
    ws.cell(row=r_sig_debut, column=2, value="O.IKKEN").font = font_bold
    ws.cell(row=r_sig_debut, column=2).alignment = align_top_center

    ws.merge_cells(start_row=r_sig_titre, start_column=6, end_row=r_sig_titre, end_column=8)
    ws.cell(row=r_sig_titre, column=6, value="Visa Chef du laboratoire").font = font_bold
    ws.cell(row=r_sig_titre, column=6).alignment = align_center

    ws.merge_cells(start_row=r_sig_debut, start_column=6, end_row=r_sig_fin, end_column=8)
    ws.cell(row=r_sig_debut, column=6, value="H.BAALLAL").font = font_bold
    ws.cell(row=r_sig_debut, column=6).alignment = align_top_center

    for r in range(1, r_sig_fin + 1):
        if r == 7:
            ws.row_dimensions[r].height = 32
        elif r == 8:
            ws.row_dimensions[r].height = 48
        elif r in [10, 11]:
            ws.row_dimensions[r].height = 23
        elif 15 <= r < (15 + nb_total):
            ws.row_dimensions[r].height = 28
        elif r in [9, 12, 13, 14]:
            ws.row_dimensions[r].height = 15
        elif r < 15:
            ws.row_dimensions[r].height = 16
        else:
            ws.row_dimensions[r].height = 28

    col_widths = {
        "A": 16,
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
# FONCTION UTILITAIRE : EXPORT EXCEL DU PLANNING DE LA DATE
# =========================================================
def exporter_dataframe_excel(df, date_chaine):
    """Génère un fichier Excel à partir du DataFrame de la liste du planning."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=f"Planning_{date_chaine}")

    buffer.seek(0)
    return buffer


# =========================================================
# FONCTIONS AUXILIAIRES DE SUPABASE
# =========================================================
def obtenir_historique_betonnage(supabase, betonnage_id):
    """Récupère l'intégralité des éprouvettes pour un même béton (betonnage_id)."""
    if not betonnage_id:
        return []
    try:
        res = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .eq("betonnage_id", betonnage_id)
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


def determiner_ref_controle(supabase, betonnage_id, info_betonnage, sample_ep):
    """Détermine la référence de contrôle avec priorité."""
    session_key = f"ref_controle_beton_{betonnage_id}"
    if session_key in st.session_state and st.session_state[session_key]:
        return st.session_state[session_key]

    ref_parent = info_betonnage.get("ref_controle") if info_betonnage else None
    if ref_parent and str(ref_parent).strip():
        st.session_state[session_key] = str(ref_parent).strip()
        return str(ref_parent).strip()

    ref_ep = sample_ep.get("ref_controle") if sample_ep else None
    if ref_ep and str(ref_ep).strip():
        st.session_state[session_key] = str(ref_ep).strip()
        return str(ref_ep).strip()

    defaut = f"REF-{betonnage_id}-{info_betonnage.get('ouvrage', 'N/A') if info_betonnage else 'N/A'}"
    st.session_state[session_key] = defaut
    return defaut


# =========================================================
# 2. APPLICATION PRINCIPALE STREAMLIT
# =========================================================
def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    # RESTRICTION D'ACCÈS
    role_utilisateur = str(
        st.session_state.get("user_role")
        or st.session_state.get("role")
        or ""
    ).lower()

    roles_autorises = ["laboratoire", "labo", "admin", "responsable_labo", "qualite"]

    if role_utilisateur not in roles_autorises and not st.session_state.get("is_admin", False):
        st.error("⛔ **Accès Restreint**")
        st.warning(
            "Ce module est réservé exclusivement au personnel du **Laboratoire de Contrôle**."
        )
        return

    est_compte_admin = (
        role_utilisateur == "admin"
        or st.session_state.get("is_admin") is True
    )

    mode_admin = False

    if est_compte_admin:
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔒 Mode Administration")
        mode_admin = st.sidebar.checkbox("Activer le Mode Admin / Edition", value=False)
        if mode_admin:
            st.sidebar.warning("⚠️ Mode Administrateur Actif.")

    tab_prog, tab_saisie, tab_hist = st.tabs([
        "📅 Phase 1 : Programmation",
        "💥 Phase 2 : Planning Daily & Saisie (Par Lot)",
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

    # PHASE 1 : PROGRAMMATION
    with tab_prog:
        st.subheader("📅 1. Programmer les Échéances d'Écrasement")

        prog_counts = {}
        try:
            res_deja_all = (
                supabase.table("suivi_controle_beton")
                .select("betonnage_id")
                .execute()
            )
            if res_deja_all.data:
                for row in res_deja_all.data:
                    b_id_val = row.get("betonnage_id")
                    if b_id_val:
                        prog_counts[b_id_val] = prog_counts.get(b_id_val, 0) + 1
        except Exception as e:
            st.warning(f"Note lors du contrôle des quotas : {e}")

        betonnages_non_programmes = []
        for b in betonnages_preleves:
            b_id = b.get("id")
            raw_nb_ep = b.get("nb_eprouvettes") or b.get("nombre_eprouvettes")
            try:
                total_prevu = int(raw_nb_ep) if raw_nb_ep is not None else 12
            except (ValueError, TypeError):
                total_prevu = 12

            deja_prog = prog_counts.get(b_id, 0)
            if (total_prevu - deja_prog) > 0 or mode_admin:
                betonnages_non_programmes.append(b)

        if not betonnages_non_programmes:
            st.info("ℹ️ Aucun bétonnage en attente de programmation.")
        else:
            options_beton = {
                (
                    f"Classe: {b.get('classe_beton', b.get('classe', 'N/A'))} | "
                    f"Date: {b.get('date_coulee', b.get('date_livraison', 'N/A'))} | "
                    f"Ouvrage: {b.get('ouvrage', 'N/A')} | "
                    f"BL: {extraire_num_bl(b)} | ID #{b['id']}"
                ): b
                for b in betonnages_non_programmes
            }

            choix_label_p = st.selectbox(
                "Sélectionner la fiche de bétonnage :",
                list(options_beton.keys()),
                key="prog_beton_select",
            )
            beton_p = options_beton[choix_label_p]

            b_id = beton_p.get("id")
            num_bl_p = extraire_num_bl(beton_p, choix_label_p)

            ouvrage_p = str(beton_p.get("ouvrage") or "-")
            classe_beton_p = str(
                beton_p.get("classe_beton") or beton_p.get("classe") or "-"
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

            eprouvettes_deja_prog = prog_counts.get(b_id, 0)
            solde_disponible = max(
                0, total_eprouvettes_prevues - eprouvettes_deja_prog
            )

            affaissement_raw = str(
                beton_p.get("affaissement") or beton_p.get("slump") or "-"
            )
            temp_beton_p = str(
                beton_p.get("temperature") or beton_p.get("temp_beton") or "-"
            )
            affaissement_p = (
                f"{affaissement_raw} mm" if affaissement_raw != "-" else "-"
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

            ref_controle_init = determiner_ref_controle(
                supabase, b_id, beton_p, {}
            )

            st.markdown("---")
            st.info(
                f"📊 **Quota Éprouvettes :** Total prévu :"
                f" **{total_eprouvettes_prevues}** | Déjà programmée(s) :"
                f" **{eprouvettes_deja_prog}** | Reste disponible :"
                f" **{solde_disponible}**"
            )

            ref_controle_p = st.text_input(
                "🏷️ Référence de Contrôle (Préfixe du repère)",
                value=ref_controle_init,
                key=f"p_ref_ctrl_{b_id}",
            )
            st.session_state[f"ref_controle_beton_{b_id}"] = ref_controle_p

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
                        f"{temp_beton_p} °C" if temp_beton_p != "-" else "-"
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

            max_allowed = solde_disponible if not mode_admin else 50
            min_val = 1 if max_allowed > 0 else 0
            val_defaut = min(2, max_allowed) if max_allowed > 0 else 0

            with col_e4:
                if max_allowed == 0 and not mode_admin:
                    st.warning("⚠️ Quota atteint.")
                    nb_eprouvettes_p = 0
                else:
                    nb_eprouvettes_p = st.number_input(
                        "Nombre d'éprouvettes à programmer",
                        min_value=min_val,
                        max_value=max_allowed,
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

            sect_def = 176.71 if "150x300" in forme_p else (201.06 if "160x320" in forme_p else 78.54)

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
                    try:
                        supabase.table("suivi_betonnage").update(
                            {"ref_controle": ref_controle_p}
                        ).eq("id", b_id).execute()
                    except Exception:
                        pass

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
                            "ref_controle": ref_controle_p,
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

    # PHASE 2 : PLANNING & SAISIE DES ÉCRASEMENTS
    with tab_saisie:
        st.subheader("💥 2. Planning des Échéances & Saisie des Écrasements")

        today_date = date.today()
        today_str = str(today_date)

        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            date_filtre = st.date_input(
                "📅 Choisir une date à consulter",
                value=today_date,
                key="filtre_date_planning"
            )
            date_filtre_str = str(date_filtre)

        try:
            res_retards = (
                supabase.table("suivi_controle_beton")
                .select("*")
                .lte("date_ecrasement", today_str)
                .or_("force_kn.is.null,force_kn.eq.0")
                .order("date_ecrasement", desc=False)
                .execute()
            )
            retards_list = res_retards.data if res_retards.data else []
        except Exception as err_retard:
            retards_list = []
            st.warning(f"Note lors de la recherche des échéances dépassées : {err_retard}")

        if retards_list:
            nb_retards = len(retards_list)
            st.error(
                f"🚨 **ATTENTION : {nb_retards} éprouvette(s) non écrasée(s) ont atteint ou dépassé leur date d'échéance !**"
            )
            
            rows_retard = []
            for ep in retards_list:
                dt_coul_str = ep.get("date_coulee")
                dt_ecras_str = ep.get("date_ecrasement")
                age_actuel = "-"

                if dt_coul_str:
                    try:
                        d_coul = datetime.strptime(str(dt_coul_str)[:10], "%Y-%m-%d").date()
                        age_actuel = f"{(today_date - d_coul).days} j (Aujourd'hui)"
                    except Exception:
                        age_actuel = str(ep.get("echeance", "-"))

                ref_p = str(ep.get("ref_controle") or "").strip()
                rep_s = str(ep.get("repere_eprouvette", "")).strip()
                rep_complet = f"{ref_p}{rep_s}" if ref_p else rep_s

                dt_ecras_obj = datetime.strptime(str(dt_ecras_str)[:10], "%Y-%m-%d").date() if dt_ecras_str else today_date
                if dt_ecras_obj < today_date:
                    statut_urgence = f"⚠️ En Retard ({(today_date - dt_ecras_obj).days} jour(s))"
                else:
                    statut_urgence = "🔥 Prévu Aujourd'hui"

                rows_retard.append({
                    "Priorité": statut_urgence,
                    "Date Écrasement Prévue": dt_ecras_str,
                    "Référence / Repère": rep_complet,
                    "N° BL": extraire_num_bl(ep),
                    "Ouvrage": ep.get("ouvrage", "-"),
                    "Classe Béton": ep.get("classe_beton", "-"),
                    "Date Coulée": dt_coul_str,
                    "Échéance Visée": ep.get("echeance", "-"),
                    "Âge Actuel Réel": age_actuel,
                })

            df_retard = pd.DataFrame(rows_retard)
            st.dataframe(df_retard, use_container_width=True, hide_index=True)
            st.markdown("---")

        try:
            res_date_sel = (
                supabase.table("suivi_controle_beton")
                .select("*")
                .eq("date_ecrasement", date_filtre_str)
                .order("id", desc=False)
                .execute()
            )
            eprouvettes_date_sel = res_date_sel.data if res_date_sel.data else []
        except Exception as err_sel:
            eprouvettes_date_sel = []
            st.warning(f"Note lors du chargement de la date {date_filtre_str} : {err_sel}")

        with st.expander(f"📆 Éprouvettes programmées spécifiquement pour le : {date_filtre_str} ({len(eprouvettes_date_sel)} éprouvette(s))", expanded=True):
            if eprouvettes_date_sel:
                rows_sel = []
                for ep in eprouvettes_date_sel:
                    dt_coul_str = ep.get("date_coulee")
                    dt_ecras_str = ep.get("date_ecrasement")
                    age_calc = "-"

                    if dt_coul_str and dt_ecras_str:
                        try:
                            d_coul = datetime.strptime(str(dt_coul_str)[:10], "%Y-%m-%d").date()
                            d_ecras = datetime.strptime(str(dt_ecras_str)[:10], "%Y-%m-%d").date()
                            age_calc = f"{(d_ecras - d_coul).days} jours"
                        except Exception:
                            age_calc = str(ep.get("echeance", "-"))

                    ref_p = str(ep.get("ref_controle") or "").strip()
                    rep_s = str(ep.get("repere_eprouvette", "")).strip()
                    rep_complet = f"{ref_p}{rep_s}" if ref_p else rep_s

                    f_kn_val = float(ep.get("force_kn") or 0.0)
                    statut_val = "✅ Écrasée" if f_kn_val > 0 else "⏳ En attente"

                    rows_sel.append({
                        "ID": ep.get("id"),
                        "Référence / Repère": rep_complet,
                        "N° BL": extraire_num_bl(ep),
                        "Ouvrage": ep.get("ouvrage", "-"),
                        "Classe Béton": ep.get("classe_beton", "-"),
                        "Date Coulée": dt_coul_str,
                        "Échéance Visée": ep.get("echeance", "-"),
                        "Âge Théorique": age_calc,
                        "Statut": statut_val
                    })

                df_sel = pd.DataFrame(rows_sel)
                st.dataframe(df_sel, use_container_width=True, hide_index=True)

                # BOUTON TÉLÉCHARGEMENT EXCEL DU PLANNING
                excel_planning_date = exporter_dataframe_excel(df_sel, date_filtre_str)
                st.download_button(
                    label=f"📊 Télécharger cette liste en Excel ({date_filtre_str})",
                    data=excel_planning_date,
                    file_name=f"Planning_Ecrasement_{date_filtre_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_download_planning_excel",
                )
            else:
                st.info(f"ℹ️ Aucune éprouvette programmée spécifiquement pour la date du {date_filtre_str}.")

        st.markdown("---")

        eprouvettes_en_attente = []
        try:
            res_att = (
                supabase.table("suivi_controle_beton")
                .select("*")
                .order("id", desc=False)
                .execute()
            )
            if res_att.data:
                if mode_admin:
                    eprouvettes_en_attente = res_att.data
                else:
                    eprouvettes_en_attente = [
                        e
                        for e in res_att.data
                        if e.get("force_kn") is None
                        or float(e.get("force_kn") or 0) == 0
                    ]
        except Exception as e:
            st.error(f"Erreur de chargement des essais en attente : {e}")

        if not eprouvettes_en_attente:
            st.info("👍 Aucune éprouvette en attente de saisie.")
        else:
            groupes_lots = {}
            for ep in eprouvettes_en_attente:
                b_id_ep = ep.get("betonnage_id")
                ech_ep = ep.get("echeance", "28 jours")
                ouv_ep = ep.get("ouvrage", "-")
                dt_ecras = ep.get("date_ecrasement", "-")
                classe_ep = ep.get("classe_beton", "-")

                info_b_temp = obtenir_infos_betonnage_parent(supabase, b_id_ep)
                ref_ctrl = determiner_ref_controle(supabase, b_id_ep, info_b_temp, ep)
                if not classe_ep or classe_ep == "-":
                    classe_ep = (info_b_temp.get("classe_beton") or info_b_temp.get("classe") or "-") if info_b_temp else "-"

                cle_groupe = (
                    f"Référence : {ref_ctrl} | Classe : {classe_ep} | Ouvrage : {ouv_ep}"
                    f" | Échéance : {ech_ep} (Date Prévue : {dt_ecras}) | Lot ID #{b_id_ep}"
                )

                if cle_groupe not in groupes_lots:
                    groupes_lots[cle_groupe] = []
                groupes_lots[cle_groupe].append(ep)

            choix_lot = st.selectbox(
                "📦 Sélectionner le lot d'éprouvettes à écraser / modifier :",
                list(groupes_lots.keys()),
                key="select_lot_saisie",
            )
            lot_selected = groupes_lots[choix_lot]

            sample = lot_selected[0]
            betonnage_id = sample.get("betonnage_id")

            info_betonnage = obtenir_infos_betonnage_parent(
                supabase, betonnage_id
            )

            historique_complet = obtenir_historique_betonnage(
                supabase, betonnage_id
            )

            exact_bl_phase1 = extraire_num_bl(sample, info_betonnage or {}, choix_lot)

            col_l1, col_l2, col_l3, col_l4 = st.columns(4)
            col_l1.metric("Client", "TGCC")
            col_l2.metric("N° Bon Livraison", exact_bl_phase1)
            col_l3.metric("Ouvrage", str((info_betonnage.get("ouvrage") if info_betonnage else None) or sample.get("ouvrage") or "-"))
            col_l4.metric("Échéance Visée", str(sample.get("echeance", "-")))

            st.markdown("---")

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                tech_global = st.text_input(
                    "Technicien / Opérateur",
                    value=sample.get("technicien", (info_betonnage.get("technicien_prelevement") if info_betonnage else None) or "Technicien LPEE"),
                    key="tech_global",
                )
            with col_g2:
                obs_globale = st.text_input(
                    "Commentaire / Observation",
                    value=sample.get("observations", "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"),
                    key="obs_global",
                )

            st.markdown("##### 📝 Saisie / Modification des forces d'écrasement")

            ref_controle_courante = determiner_ref_controle(
                supabase, betonnage_id, info_betonnage, sample
            )

            lot_key = f"df_lot_{choix_lot}"

            if lot_key not in st.session_state or mode_admin:
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
                        "🏷️ Référence de Contrôle": ref_controle_courante,
                        "Repère": ep.get("repere_eprouvette", f"/{ep['id']}"),
                        "Forme d'éprouvette": str(
                            ep.get("forme") or "Cylindrique 150x300"
                        ),
                        "_section": sec,
                        "Force (kN)": f_kn,
                        "Résistance Fc (MPa)": fc,
                        "Moyenne Resistance Fc (MPa)": 0.0,
                    })
                df_init = pd.DataFrame(rows_list)

                valides_init = df_init[df_init["Résistance Fc (MPa)"] > 0]
                moy_init = round(valides_init["Résistance Fc (MPa)"].mean(), 1) if not valides_init.empty else 0.0
                df_init["Moyenne Resistance Fc (MPa)"] = moy_init

                st.session_state[lot_key] = df_init

            def update_fc():
                editor_state = st.session_state.get("data_editor_ecrasement", {})
                changes = editor_state.get("edited_rows", {})

                for row_idx, updated_cols in changes.items():
                    if "Force (kN)" in updated_cols:
                        raw_force = updated_cols["Force (kN)"]
                        try:
                            new_force = float(raw_force) if raw_force is not None else 0.0
                        except (ValueError, TypeError):
                            new_force = 0.0

                        sec = float(st.session_state[lot_key].at[row_idx, "_section"])
                        st.session_state[lot_key].at[row_idx, "Force (kN)"] = new_force

                        if sec > 0 and new_force > 0:
                            st.session_state[lot_key].at[row_idx, "Résistance Fc (MPa)"] = round((new_force * 10.0) / sec, 1)
                        else:
                            st.session_state[lot_key].at[row_idx, "Résistance Fc (MPa)"] = 0.0

                    if "🏷️ Référence de Contrôle" in updated_cols:
                        nouvelle_ref = str(updated_cols["🏷️ Référence de Contrôle"] or "").strip()
                        st.session_state[lot_key].at[row_idx, "🏷️ Référence de Contrôle"] = nouvelle_ref
                        st.session_state[f"ref_controle_beton_{betonnage_id}"] = nouvelle_ref

                df_cur = st.session_state[lot_key]
                forces_valides = df_cur[df_cur["Résistance Fc (MPa)"].astype(float) > 0]
                fc_moy = (
                    round(forces_valides["Résistance Fc (MPa)"].astype(float).mean(), 1)
                    if not forces_valides.empty
                    else 0.0
                )
                st.session_state[lot_key]["Moyenne Resistance Fc (MPa)"] = fc_moy

            st.data_editor(
                st.session_state[lot_key],
                column_config={
                    "ID": st.column_config.NumberColumn("ID", disabled=True),
                    "🏷️ Référence de Contrôle": st.column_config.TextColumn(
                        "🏷️ Référence de Contrôle (Préfixe)",
                        help="Préfixe conservé pour tous les lots du même prélèvement",
                    ),
                    "Repère": st.column_config.TextColumn(
                        "Repère", disabled=not mode_admin
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
                    "Moyenne Resistance Fc (MPa)": st.column_config.NumberColumn(
                        "📊 Moyenne Resistance Fc (MPa)", disabled=True, format="%.1f"
                    ),
                },
                use_container_width=True,
                hide_index=True,
                key="data_editor_ecrasement",
                on_change=update_fc,
            )

            df_actuel = st.session_state[lot_key]

            export_data = []
            dict_actuel = {int(row["ID"]): row for _, row in df_actuel.iterrows()}

            items_source = historique_complet if historique_complet else lot_selected

            for ep_h in items_source:
                ep_id = ep_h["id"]
                sec_h = float(ep_h.get("section") or 176.71)

                if ep_id in dict_actuel:
                    row_saisie = dict_actuel[ep_id]
                    f_kn = float(row_saisie["Force (kN)"])
                    fc_mpa = float(row_saisie["Résistance Fc (MPa)"])
                    ref_p = str(row_saisie["🏷️ Référence de Contrôle"]).strip()
                    rep_s = str(row_saisie["Repère"]).strip()
                    statut = "En cours" if f_kn == 0 else "Réalisé"
                else:
                    f_kn = float(ep_h.get("force_kn") or 0.0)
                    fc_mpa = float(ep_h.get("fc_mpa") or (round((f_kn * 10.0) / sec_h, 1) if f_kn > 0 else 0.0))
                    ref_p = str(ep_h.get("ref_controle") or ref_controle_courante).strip()
                    rep_s = str(ep_h.get("repere_eprouvette", f"/{ep_id}")).strip()
                    statut = "En cours" if f_kn == 0 else "Réalisé"

                rep_c = f"{ref_p}{rep_s}" if ref_p else rep_s

                export_data.append({
                    "repere_eprouvette": rep_c,
                    "forme": ep_h.get("forme", "Cylindrique 150x300"),
                    "section": sec_h,
                    "force_kn": f_kn,
                    "fc_mpa": fc_mpa,
                    "date_essai": ep_h.get("date_ecrasement", "-"),
                    "age": str(ep_h.get("echeance", "28"))
                    .replace(" jours", "")
                    .replace("j", ""),
                    "statut": statut,
                })

            num_bl_valeur = exact_bl_phase1
            affaissement_saisi = (
                (info_betonnage.get("affaissement") or info_betonnage.get("slump"))
                if info_betonnage
                else None
            )
            temp_saisie = (
                (info_betonnage.get("temperature") or info_betonnage.get("temp_beton"))
                if info_betonnage
                else None
            )
            ouvrage_saisi = (
                (info_betonnage.get("ouvrage") if info_betonnage else None)
                or sample.get("ouvrage")
            )
            date_coulee_saisie = (
                (info_betonnage.get("date_coulee") if info_betonnage else None)
                or sample.get("date_coulee")
            )
            centrale_saisie = (
                (
                    info_betonnage.get("centrale")
                    or info_betonnage.get("centrale_beton")
                )
                if info_betonnage
                else sample.get("centrale")
            )
            tech_prelev = (
                (
                    info_betonnage.get("technicien_prelevement")
                    or info_betonnage.get("preleve_par")
                    or info_betonnage.get("technicien")
                )
                if info_betonnage
                else tech_global
            )

            infos_header = {
                "re_num": "25/260/LGV/ B/01",
                "dossier": "2025-260-05985-2025-0247",
                "client": "TGCC",
                "num_bl": num_bl_valeur,
                "ouvrage": ouvrage_saisi,
                "lieu_prelevement": ouvrage_saisi,
                "classe_beton": sample.get("classe_beton", "C35/45"),
                "date_coulee": date_coulee_saisie,
                "affaissement": affaissement_saisi,
                "temperature": temp_saisie,
                "forme": sample.get("forme", "Cylindrique 150x300"),
                "centrale": centrale_saisie,
                "observations": obs_globale,
                "technicien_prelevement": tech_prelev,
            }

            excel_file = generer_pv_excel(export_data, infos_header)
            filename = f"PV_Ecrasement_LPEE_{num_bl_valeur if num_bl_valeur != '-' else 'BL'}.xlsx"

            st.markdown("---")
            col_b1, col_b2 = st.columns(2)

            with col_b1:
                label_btn = "💾 Valider et Mettre à Jour Le Lot" if mode_admin else "💾 Valider et Enregistrer Le Lot"
                btn_enregistrer = st.button(
                    label_btn,
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
                if (df_actuel["Force (kN)"].astype(float) == 0).any() and not mode_admin:
                    st.error(
                        "❌ Les forces de rupture doivent toutes être saisies"
                        " (> 0 kN)."
                    )
                else:
                    succes_lot = 0
                    ref_finale = df_actuel.iloc[0].get("🏷️ Référence de Contrôle")

                    try:
                        supabase.table("suivi_betonnage").update(
                            {"ref_controle": ref_finale}
                        ).eq("id", betonnage_id).execute()
                    except Exception:
                        pass

                    for _, row in df_actuel.iterrows():
                        update_payload = {
                            "ref_controle": row.get("🏷️ Référence de Contrôle"),
                            "repere_eprouvette": row.get("Repère"),
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
                            f"✅ Lot de {succes_lot} éprouvettes mis à jour / validé !"
                        )

    # HISTORIQUE COMPLET & ÉDITION DE PV
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

                df_valides = df_all[
                    (df_all["force_kn"].notnull()) & (df_all["force_kn"] > 0)
                ].copy()

                if not df_valides.empty:
                    st.markdown("##### 📥 Re-télécharger un Procès-Verbal")

                    groupes_valides = {}
                    for _, row in df_valides.iterrows():
                        b_id_ep = row.get("betonnage_id")
                        ech_ep = row.get("echeance", "28 jours")
                        ouv_ep = row.get("ouvrage", "-")
                        dt_ecras = row.get("date_ecrasement", "-")
                        classe_ep = row.get("classe_beton", "-")

                        info_b_temp = obtenir_infos_betonnage_parent(supabase, b_id_ep)
                        ref_ctrl = determiner_ref_controle(supabase, b_id_ep, info_b_temp, row.to_dict())
                        if not classe_ep or classe_ep == "-":
                            classe_ep = (info_b_temp.get("classe_beton") or info_b_temp.get("classe") or "-") if info_b_temp else "-"

                        cle_pv = (
                            f"Référence : {ref_ctrl} | Classe : {classe_ep} | Ouvrage : {ouv_ep}"
                            f" | Échéance : {ech_ep} (Date : {dt_ecras}) | Lot ID #{b_id_ep}"
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
                            or (round((f_kn * 10.0) / sec, 1) if f_kn > 0 else 0.0)
                        )

                        ref_p = str(item.get("ref_controle") or "").strip()
                        rep_s = str(
                            item.get("repere_eprouvette", f"/{item['id']}")
                        ).strip()
                        rep_c = f"{ref_p}{rep_s}" if ref_p else rep_s
                        statut = "En cours" if f_kn == 0 else "Réalisé"

                        export_data_h.append({
                            "repere_eprouvette": rep_c,
                            "forme": item.get("forme", "Cylindrique 150x300"),
                            "section": sec,
                            "force_kn": f_kn,
                            "fc_mpa": fc,
                            "date_essai": item.get("date_ecrasement", "-"),
                            "age": str(item.get("echeance", "28"))
                            .replace(" jours", "")
                            .replace("j", ""),
                            "statut": statut,
                        })

                    num_bl_h = extraire_num_bl(sample_h, info_beton_h or {}, choix_pv_hist)
                    aff_h = (
                        (info_beton_h.get("affaissement") or info_beton_h.get("slump"))
                        if info_beton_h
                        else None
                    )
                    temp_h = (
                        (info_beton_h.get("temperature") or info_beton_h.get("temp_beton"))
                        if info_beton_h
                        else None
                    )
                    ouv_h = (
                        (info_beton_h.get("ouvrage") if info_beton_h else None)
                        or sample_h.get("ouvrage")
                    )
                    date_coulee_h = (
                        (info_beton_h.get("date_coulee") if info_beton_h else None)
                        or sample_h.get("date_coulee")
                    )
                    centrale_h = (
                        (
                            info_beton_h.get("centrale")
                            or info_beton_h.get("centrale_beton")
                        )
                        if info_beton_h
                        else sample_h.get("centrale")
                    )
                    tech_prelev_h = (
                        (
                            info_beton_h.get("technicien_prelevement")
                            or info_beton_h.get("preleve_par")
                            or info_beton_h.get("technicien")
                        )
                        if info_beton_h
                        else sample_h.get("technicien")
                    )

                    infos_header_h = {
                        "re_num": "25/260/LGV/ B/01",
                        "dossier": "2025-260-05985-2025-0247",
                        "client": "TGCC",
                        "num_bl": num_bl_h,
                        "ouvrage": ouv_h,
                        "lieu_prelevement": ouv_h,
                        "classe_beton": sample_h.get("classe_beton", "C35/45"),
                        "date_coulee": date_coulee_h,
                        "affaissement": aff_h,
                        "temperature": temp_h,
                        "forme": sample_h.get("forme", "Cylindrique 150x300"),
                        "centrale": centrale_h,
                        "observations": sample_h.get(
                            "observations",
                            "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES",
                        ),
                        "technicien_prelevement": tech_prelev_h,
                    }

                    excel_pv_hist = generer_pv_excel(
                        export_data_h, infos_header_h
                    )
                    file_name_h = f"PV_Ecrasement_RE-EXPORT_{num_bl_h if num_bl_h != '-' else 'BL'}.xlsx"

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

                if "affaissement" not in df_all.columns:
                    df_all["affaissement"] = None
                if "temperature" not in df_all.columns:
                    df_all["temperature"] = None

                for idx_row, row_data in df_all.iterrows():
                    b_id_row = row_data.get("betonnage_id")
                    if b_id_row:
                        parent_info = obtenir_infos_betonnage_parent(supabase, b_id_row)
                        if parent_info:
                            df_all.at[idx_row, "affaissement"] = parent_info.get("affaissement") or parent_info.get("slump") or "-"
                            df_all.at[idx_row, "temperature"] = parent_info.get("temperature") or parent_info.get("temp_beton") or "-"
                        else:
                            df_all.at[idx_row, "affaissement"] = "-"
                            df_all.at[idx_row, "temperature"] = "-"

                renommage_colonnes = {
                    "affaissement": "affaissement_mm",
                    "temperature": "temp_beton_C",
                }

                colonnes_ordre = [
                    "id",
                    "ref_controle",
                    "repere_eprouvette",
                    "date_coulee",
                    "classe_beton",
                    "ouvrage",
                    "date_ecrasement",
                    "echeance",
                    "force_kn",
                    "fc_mpa",
                    "technicien",
                ]

                cols_disponibles = [c for c in colonnes_ordre if c in df_all.columns]
                cols_restantes = [c for c in df_all.columns if c not in cols_disponibles]
                df_ordered = df_all[cols_disponibles + cols_restantes].rename(columns=renommage_colonnes)

                st.dataframe(df_ordered, use_container_width=True, hide_index=True)
            else:
                st.info("Aucun enregistrement d'écrasement dans la base.")
        except Exception as e:
            st.error(f"Erreur lors du chargement de l'historique : {e}")
