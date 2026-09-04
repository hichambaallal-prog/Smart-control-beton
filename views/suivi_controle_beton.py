import base64
import io
import math
import re
import unicodedata
from datetime import date, datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import qrcode
from audit_log import enregistrer_modification, afficher_historique_modifications
import projets_config

try:
    from supabase import create_client, Client
except ImportError:
    Client = None

OPTIONS_ONGLETS = [
    "📋 Phase 0 : Réception & Validation",
    "📅 Phase 1 : Programmation",
    "💥 Phase 2 : Planning Daily & Saisie (Par Lot)",
    "🛡️ Phase 3 : Validation Admin (PVs)",
]


# ==============================================================================
# FONCTIONS UTILITAIRES : CALCUL DE RÉSISTANCE ET EXTRACTION DATES
# ==============================================================================
def calculer_resistance_mpa(force_kn, type_essai="Compression", forme="Cylindrique 150x300"):
    """
    Calcule la résistance en MPa selon le type d'essai :
    - Compression (NF EN 12390-3) : Fc = Force (N) / Section (mm²)
    - Traction par fendage (NF EN 12390-6) : Fct = (2 * Force) / (pi * L * d)
    """
    try:
        f_kn = float(force_kn or 0.0)
    except (ValueError, TypeError):
        return 0.0

    if f_kn <= 0:
        return 0.0

    force_n = f_kn * 1000.0

    if "160x320" in str(forme):
        d, L, sec = 160.0, 320.0, 20106.19
    elif "100x200" in str(forme):
        d, L, sec = 100.0, 200.0, 7853.98
    else:  # Cylindrique 150x300 par défaut
        d, L, sec = 150.0, 300.0, 17671.46

    if type_essai == "Traction par fendage":
        f_ct = (2.0 * force_n) / (math.pi * L * d)
        return round(f_ct, 2)
    else:
        f_c = force_n / sec
        return round(f_c, 1)


def calculer_date_ecrasement(df):
    df_result = df.copy()
    col_ech = next((c for c in ['Échéance Visée', 'echeance', 'Échéance'] if c in df_result.columns), None)
    col_coul = next((c for c in ['Date Coulée', 'date_coulee'] if c in df_result.columns), None)

    if not col_coul or not col_ech:
        return df_result

    df_result['Date Coulée'] = pd.to_datetime(df_result[col_coul], errors='coerce')
    nb_jours = (
        df_result[col_ech]
        .astype(str)
        .str.extract(r'(\d+)')
        .fillna(28)[0]
        .astype(int)
    )

    df_result['Date Écrasement Prévue'] = df_result['Date Coulée'] + pd.to_timedelta(nb_jours, unit='D')
    df_result['Date Coulée'] = df_result['Date Coulée'].dt.strftime('%Y-%m-%d')
    df_result['Date Écrasement Prévue'] = df_result['Date Écrasement Prévue'].dt.strftime('%Y-%m-%d')

    return df_result


def extraire_nb_jours(echeance_str, default=28):
    if pd.isna(echeance_str) or not echeance_str:
        return default
    match = re.search(r'\d+', str(echeance_str))
    return int(match.group()) if match else default


# ==============================================================================
# 1. GESTION DES UTILISATEURS ET SUPABASE
# ==============================================================================
def connecter_utilisateur(supabase, nom_utilisateur, mot_de_passe):
    try:
        res = (
            supabase.table("users")
            .select("*")
            .eq("username", nom_utilisateur)
            .eq("password", mot_de_passe)
            .execute()
        )
        if res.data:
            u = res.data[0]
            st.session_state.update({
                "user_logged": True,
                "user": u,
                "username": u.get("username"),
                "role": u.get("role"),
                "user_role": u.get("role"),
                "can_edit": bool(u.get("can_edit", False)),
            })
            return True
        st.error("Nom d'utilisateur ou mot de passe incorrect.")
    except Exception as e:
        st.error(f"Erreur lors de la connexion : {e}")
    return False


def generer_qr_code(data_url):
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(data_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return buffer.getvalue()


def verifier_doublon_num_reception(supabase, num_reception, current_beton_id=None, projet_id=None):
    if not num_reception or str(num_reception).strip() in ["", "-", "None", "NaN", "N/A"]:
        return False
    num_clean = str(num_reception).strip()
    try:
        query = (
            supabase.table("suivi_betonnage")
            .select("id, num_reception")
            .eq("num_reception", num_clean)
        )
        if projet_id:
            query = query.eq("projet_id", projet_id)
        res = query.execute()
        for m in (res.data or []):
            if current_beton_id is None or int(m.get("id")) != int(current_beton_id):
                return True
    except Exception as e:
        st.warning(f"Note lors de la vérification des doublons : {e}")
    return False


def extraire_num_bl(*sources):
    keys = ["num_bl", "bl", "num_bon_livraison", "n_bl", "bon_livraison", "num_bl_p", "n_bon", "bon_de_livraison", "code_bl"]
    invalid = {"N/A", "NONE", "NAN", "-", ""}
    for src in sources:
        if isinstance(src, dict):
            for k in keys:
                val = str(src.get(k) or "").strip()
                if val and val.upper() not in invalid:
                    return val
            for k, v in src.items():
                if "bl" in k.lower() or "bon" in k.lower():
                    val = str(v or "").strip()
                    if val and val.upper() not in invalid:
                        return val
        elif isinstance(src, str):
            match = re.search(r"BL\s*:\s*([^\|]+)", src, re.IGNORECASE)
            if match:
                val = match.group(1).strip()
                if val and val.upper() not in invalid:
                    return val
    return "-"


def extraire_date_coulee(item):
    if not item or not isinstance(item, dict):
        return str(date.today())
    for k in ["date_coulee", "date_livraison", "date_prelevement"]:
        val = str(item.get(k) or "").strip()
        if val and val.upper() not in ["N/A", "NONE", "NAN", "-", ""]:
            return val[:10]
    return str(date.today())


# =========================================================
# 2. GÉNÉRATION DU PROCÈS-VERBAL EXCEL
# =========================================================
def generer_pv_excel(export_data, infos_header):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PV Écrasement LPEE"

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.2, footer=0.2)

    font_bold = Font(name="Calibri", size=9, bold=True)
    font_bold_white = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_title_white = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=8.5)
    font_small = Font(name="Calibri", size=8)

    fill_dark = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_table = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_label = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)

    border_cell = Border(left=Side(style="thin", color="000000"), right=Side(style="thin", color="000000"),
                         top=Side(style="thin", color="000000"), bottom=Side(style="thin", color="000000"))

    default_bl = extraire_num_bl(infos_header)

    def remplacer_na(valeur, fallback=None):
        val_str = str(valeur).strip() if valeur is not None else ""
        return fallback if fallback is not None else default_bl if val_str.upper() in ["N/A", "NONE", "NAN", "", "-"] else valeur

    def format_cell(cell, font=font_regular, align=align_center, fill=None, border=border_cell):
        if font: cell.font = font
        if align: cell.alignment = align
        if fill: cell.fill = fill
        if border: cell.border = border

    # ENTÊTE
    ws.merge_cells("A1:D1")
    ws["A1"] = "LPEE / CTR CSB"
    format_cell(ws["A1"], font_bold_white, align_center)

    ws.merge_cells("A2:D3")
    ws["A2"] = "Laboratoire de Contrôle Externe"
    format_cell(ws["A2"], font_bold_white, align_center)

    for r in range(1, 4):
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = fill_dark

    ws["E1"] = "RE N° :"
    format_cell(ws["E1"], font_bold, align_left)
    ws.merge_cells("F1:G1")
    ws["F1"] = remplacer_na(infos_header.get("re_num"), "25/260/LGV/ B/")
    format_cell(ws["F1"], font_regular, align_left)
    
    ws["H1"] = "BETON"
    format_cell(ws["H1"], font_bold, align_center)

    ws["E2"] = "DOSSIER :"
    format_cell(ws["E2"], font_bold, align_left)
    ws.merge_cells("F2:H2")
    ws["F2"] = remplacer_na(infos_header.get("dossier"), "2025-260-05985-2025-0247")
    format_cell(ws["F2"], font_regular, align_left)

    ws["E3"] = "CLIENT :"
    format_cell(ws["E3"], font_bold, align_left)
    ws.merge_cells("F3:H3")
    ws["F3"] = remplacer_na(infos_header.get("client"), "TGCC")
    format_cell(ws["F3"], font_bold, align_left)

    for r in range(1, 4):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # TITRE
    ws.merge_cells("A4:H4")
    ws["A4"] = "ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
    format_cell(ws["A4"], font_title_white, align_center, fill_dark)
    for c in range(1, 9):
        ws.cell(row=4, column=c).fill = fill_dark
        ws.cell(row=4, column=c).border = border_cell

    # Détection du type d'essai prédominant
    a_fendage = any(str(item.get("type_essai", "")).strip() == "Traction par fendage" for item in export_data)
    ws.merge_cells("A5:D5")
    ws["A5"] = f"[{' ' if a_fendage else 'X'}] COMPRESSION NF EN 12390-3 (2019)"
    format_cell(ws["A5"], font_bold, align_center)

    ws.merge_cells("E5:H5")
    ws["E5"] = f"[{'X' if a_fendage else ' '}] TRACTION PAR FENDAGE NF EN 12390-6 (2019)"
    format_cell(ws["E5"], font_bold, align_center)

    for c in range(1, 9): ws.cell(row=5, column=c).border = border_cell

    ws.merge_cells("A6:F6")
    ws["A6"] = "Presse : Marque: Controls"
    format_cell(ws["A6"], font_bold, align_right)

    ws.merge_cells("G6:H6")
    ws["G6"] = "Classe : A"
    format_cell(ws["G6"], font_bold, align_center)

    for c in range(1, 9): ws.cell(row=6, column=c).border = border_cell

    # FICHE TECHNIQUE
    ws["A7"] = "Date de\nprélèvement"
    format_cell(ws["A7"], font_bold, align_center)
    ws["B7"] = str(remplacer_na(infos_header.get("date_coulee"), "-"))
    format_cell(ws["B7"], font_bold, align_center)

    ws.merge_cells("C7:D7")
    ws["C7"] = "Lieu de\nprélèvement"
    format_cell(ws["C7"], font_bold, align_center)

    ws.merge_cells("E7:H7")
    ws["E7"] = remplacer_na(infos_header.get("lieu_prelevement", infos_header.get("ouvrage")), "-")
    format_cell(ws["E7"], font_regular, align_center)

    ws["A8"] = "Chantier"
    format_cell(ws["A8"], font_bold, align_center)

    ws.merge_cells("B8:D8")
    ws["B8"] = remplacer_na(infos_header.get("chantier"), "Augmentation de la capacité ferroviaire entre Kenitra et Marrakech")
    format_cell(ws["B8"], font_small, align_center)

    ws.merge_cells("E8:F8")
    ws["E8"] = "Type de béton"
    format_cell(ws["E8"], font_bold, align_center)

    ws.merge_cells("G8:H8")
    classe_beton_val = str(remplacer_na(infos_header.get("classe_beton"), "C35/45")).upper()
    ws["G8"] = classe_beton_val
    format_cell(ws["G8"], font_bold, align_center)

    ws.merge_cells("A9:B9")
    ws["A9"] = remplacer_na(infos_header.get("centrale"), "Centrale à Béton")
    format_cell(ws["A9"], font_bold, align_center)

    ws["C9"] = "- Dimensions"
    format_cell(ws["C9"], font_regular, align_left)

    ws.merge_cells("D9:H9")
    ws["D9"] = remplacer_na(infos_header.get("forme"), "Cylindrique 150x300")
    format_cell(ws["D9"], font_bold, align_center)

    ws.merge_cells("A10:B10")
    ws["A10"] = "Affaissement au cône d'abrams NF EN 12350-2"
    format_cell(ws["A10"], font_small, align_center)

    ws["C10"] = str(remplacer_na(infos_header.get("affaissement"), "-"))
    format_cell(ws["C10"], font_bold, align_center)

    ws["D10"] = "- Mode confection"
    format_cell(ws["D10"], font_regular, align_left)

    ws.merge_cells("E10:H10")
    ws["E10"] = "Par vibration NF EN 12390-2 (2019)"
    format_cell(ws["E10"], font_bold, align_center)

    ws.merge_cells("A11:B11")
    ws["A11"] = "Température °C"
    format_cell(ws["A11"], font_regular, align_center)

    ws["C11"] = str(remplacer_na(infos_header.get("temperature"), "-"))
    format_cell(ws["C11"], font_bold, align_center)

    ws["D11"] = "- Mode conservation"
    format_cell(ws["D11"], font_regular, align_left)

    ws.merge_cells("E11:H11")
    ws["E11"] = "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C ± 2°C"
    format_cell(ws["E11"], font_bold, align_center)

    tech_prelevement = remplacer_na(infos_header.get("technicien_prelevement") or infos_header.get("preleve_par") or infos_header.get("technicien"), "Technicien LPEE")
    ws.merge_cells("A12:C12")
    ws["A12"] = f"prélèvement effectué par {tech_prelevement}"
    format_cell(ws["A12"], font_small, align_center)

    ws.merge_cells("D12:E12")
    ws["D12"] = "N° de bon de livraison"
    format_cell(ws["D12"], font_regular, align_center)

    ws.merge_cells("F12:H12")
    ws["F12"] = default_bl
    format_cell(ws["F12"], font_bold, align_center)

    labels_coords = {"A7", "C7", "A8", "E8", "A9", "A10", "A11", "A12", "D12"}
    for r in range(7, 13):
        for c in range(1, 9):
            cell = ws.cell(row=r, column=c)
            cell.border = border_cell
            if cell.coordinate in labels_coords:
                cell.fill = fill_label

    # TABLEAU DES RÉSULTATS
    ws.merge_cells("A13:A14")
    ws["A13"] = "Réf,"
    ws.merge_cells("B13:C13")
    ws["B13"] = "Date"
    ws["B14"], ws["C14"] = "Fabri", "Essai"
    ws.merge_cells("D13:D14")
    ws["D13"] = "Age (jours)"
    ws.merge_cells("E13:E14")
    ws["E13"] = "Charge rupture(KN)"
    ws.merge_cells("F13:H13")
    ws["F13"] = "Résistance (MPa)"
    ws["F14"], ws["G14"], ws["H14"] = "Compression", "Traction", "Moyenne"

    for r in range(13, 15):
        for c in range(1, 9):
            format_cell(ws.cell(row=r, column=c), font=font_bold if r == 13 or c == 1 else font_regular, align=align_center, fill=fill_table)

    row_start = 15
    nb_total = len(export_data)
    groupes_lots = {}
    a_des_28j_ecrases, cellule_moyenne_28j = False, None

    for idx, item in enumerate(export_data):
        curr_row = row_start + idx
        try:
            f_kn_val = float(item.get("force_kn", 0.0))
        except (ValueError, TypeError):
            f_kn_val = 0.0
            
        is_en_cours = str(item.get("statut", "")).lower() == "en cours" or f_kn_val == 0.0
        dt_essai = item.get("date_essai") or item.get("date_ecrasement")

        ws.cell(row=curr_row, column=1, value=str(item.get("repere_eprouvette", "B/01")))
        ws.cell(row=curr_row, column=2, value=str(remplacer_na(infos_header.get("date_coulee"), "-")))
        ws.cell(row=curr_row, column=3, value=str(dt_essai) if (is_en_cours and dt_essai and dt_essai != "-") else ("En cours" if is_en_cours else str(remplacer_na(dt_essai, "-"))))

        age_val = extraire_nb_jours(item.get("age"), default=7)
        ws.cell(row=curr_row, column=4, value=age_val)

        t_essai = str(item.get("type_essai", "Compression")).strip()

        if is_en_cours:
            ws.cell(row=curr_row, column=5, value="En cours")
            ws.cell(row=curr_row, column=6, value="En cours")
            ws.cell(row=curr_row, column=7, value="-")
        else:
            try:
                fc_val = float(item.get("fc_mpa", 0.0))
            except (ValueError, TypeError):
                fc_val = 0.0
            ws.cell(row=curr_row, column=5, value=f_kn_val).number_format = "0.0"
            
            if t_essai == "Traction par fendage":
                ws.cell(row=curr_row, column=6, value="-")
                ws.cell(row=curr_row, column=7, value=fc_val).number_format = "0.00"
            else:
                ws.cell(row=curr_row, column=6, value=fc_val).number_format = "0.0"
                ws.cell(row=curr_row, column=7, value="-")

        for c in range(1, 9):
            format_cell(ws.cell(row=curr_row, column=c), font=font_regular, align=align_center)

        cle_lot = f"{item.get('age')}_{dt_essai}_{t_essai}"
        if cle_lot not in groupes_lots:
            groupes_lots[cle_lot] = {"lignes": [], "en_cours": is_en_cours, "age": age_val, "type_essai": t_essai}
        elif is_en_cours:
            groupes_lots[cle_lot]["en_cours"] = True
        groupes_lots[cle_lot]["lignes"].append(curr_row)

    for cle_lot, data_lot in groupes_lots.items():
        lignes = data_lot["lignes"]
        start_r, end_r = min(lignes), max(lignes)
        cell_h = ws[f"H{start_r}"]
        col_res = "G" if data_lot["type_essai"] == "Traction par fendage" else "F"

        if data_lot["en_cours"]:
            if start_r != end_r: ws.merge_cells(f"H{start_r}:H{end_r}")
            cell_h.value = "En cours"
        else:
            if start_r == end_r:
                cell_h.value = f"=ROUND({col_res}{start_r}, 1)"
            else:
                ws.merge_cells(f"H{start_r}:H{end_r}")
                cell_h.value = f"=ROUND(AVERAGE({col_res}{start_r}:{col_res}{end_r}), 1)"
            cell_h.number_format = "0.00" if data_lot["type_essai"] == "Traction par fendage" else "0.0"
            if extraire_nb_jours(data_lot["age"]) >= 28:
                a_des_28j_ecrases, cellule_moyenne_28j = True, f"H{start_r}"

        format_cell(cell_h, font=font_bold, align=align_center)

    next_row = row_start + nb_total
    ws.cell(row=next_row, column=1, value="Commentaire :")
    format_cell(ws.cell(row=next_row, column=1), font=font_bold, align=align_left, fill=fill_label)

    ws.merge_cells(f"B{next_row}:H{next_row}")
    obs_defaut = infos_header.get("observations") or "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"

    seuil_min = 35.0
    for classe, seuil in [("C25/30", 25.0), ("C30/37", 30.0), ("C35/45", 35.0), ("C40/50", 40.0)]:
        if classe in classe_beton_val:
            seuil_min = seuil
            break

    if not a_des_28j_ecrases or not cellule_moyenne_28j:
        formule_commentaires = "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT."
    else:
        m_cell = cellule_moyenne_28j
        formule_commentaires = f'=IF(OR(ISBLANK({m_cell}), {m_cell}="En cours"), "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT.", IF({m_cell}>={seuil_min}, "{obs_defaut}", "PERFORMANCES MECANIQUES NON CONFORMES"))'

    format_cell(ws.cell(row=next_row, column=2, value=formule_commentaires), font=font_bold, align=align_left)
    for c in range(1, 9): ws.cell(row=next_row, column=c).border = border_cell

    r_sig_titre = next_row + 2
    r_sig_debut, r_sig_fin = r_sig_titre + 1, r_sig_titre + 4

    ws.merge_cells(start_row=r_sig_titre, start_column=2, end_row=r_sig_titre, end_column=4)
    format_cell(ws.cell(row=r_sig_titre, column=2, value="Visa Responsable d'essai"), font=font_bold, align=align_center)

    ws.merge_cells(start_row=r_sig_debut, start_column=2, end_row=r_sig_fin, end_column=4)
    format_cell(ws.cell(row=r_sig_debut, column=2, value="O.IKKEN"), font=font_bold, align=align_top_center)

    ws.merge_cells(start_row=r_sig_titre, start_column=6, end_row=r_sig_titre, end_column=8)
    format_cell(ws.cell(row=r_sig_titre, column=6, value="Visa Chef du laboratoire"), font=font_bold, align=align_center)

    ws.merge_cells(start_row=r_sig_debut, start_column=6, end_row=r_sig_fin, end_column=8)
    format_cell(ws.cell(row=r_sig_debut, column=6, value="H.BAALLAL"), font=font_bold, align=align_top_center)

    row_heights = {7: 32, 8: 48, 10: 23, 11: 23, 9: 15, 12: 15, 13: 15, 14: 15}
    for r in range(1, r_sig_fin + 1):
        if r in row_heights: ws.row_dimensions[r].height = row_heights[r]
        elif 15 <= r < (15 + nb_total) or r > 14: ws.row_dimensions[r].height = 28
        else: ws.row_dimensions[r].height = 16

    col_widths = {"A": 16, "B": 12, "C": 12, "D": 10, "E": 18, "F": 14, "G": 12, "H": 12}
    for col, width in col_widths.items(): ws.column_dimensions[col].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================================================
# FONCTIONS AUXILIAIRES DE SUPABASE & PARSING
# =========================================================
def exporter_dataframe_excel(df, date_chaine):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=f"Planning_{date_chaine}"[:31])
    buffer.seek(0)
    return buffer


def obtenir_historique_betonnage(supabase, betonnage_id):
    if not betonnage_id: return []
    try:
        res = supabase.table("suivi_controle_beton").select("*").eq("betonnage_id", betonnage_id).order("id", desc=False).execute()
        return res.data or []
    except Exception as e:
        st.warning(f"Note : Historique du bétonnage #{betonnage_id} non disponible : {e}")
        return []


def obtenir_infos_betonnage_parent(supabase, betonnage_id):
    if not betonnage_id: return {}
    try:
        res = supabase.table("suivi_betonnage").select("*").eq("id", betonnage_id).execute()
        if res.data: return res.data[0]
    except Exception as e:
        st.warning(f"Note : Impossible de charger la fiche parent #{betonnage_id} : {e}")
        return {}


def obtenir_infos_betonnage_parents_bulk(supabase, betonnage_ids):
    ids_valides = sorted({int(b) for b in betonnage_ids if b is not None})
    if not ids_valides:
        return {}
    try:
        res = supabase.table("suivi_betonnage").select("*").in_("id", ids_valides).execute()
        return {p["id"]: p for p in (res.data or [])}
    except Exception as e:
        st.warning(f"Note : Impossible de charger les fiches parentes en lot : {e}")
        return {}


def determiner_ref_controle(supabase, betonnage_id, info_betonnage, sample_ep):
    session_key = f"ref_controle_beton_{betonnage_id}"
    if st.session_state.get(session_key): return st.session_state[session_key]

    num_rec = (info_betonnage or {}).get("num_reception")
    if num_rec and str(num_rec).strip() not in ["", "-", "None", "NaN", "N/A"]:
        st.session_state[session_key] = str(num_rec).strip()
        return str(num_rec).strip()

    for candidate in [(info_betonnage or {}).get("ref_controle"), (sample_ep or {}).get("ref_controle")]:
        if candidate and str(candidate).strip():
            st.session_state[session_key] = str(candidate).strip()
            return str(candidate).strip()

    defaut = f"REF-{betonnage_id}-{(info_betonnage or {}).get('ouvrage', 'N/A')}"
    st.session_state[session_key] = defaut
    return defaut


def _format_ep_row(ep, date_ref=None):
    dt_coul, dt_ecras = ep.get("date_coulee"), ep.get("date_ecrasement")
    age_calc = "-"
    if dt_coul and dt_ecras:
        try:
            d_c, d_e = datetime.strptime(str(dt_coul)[:10], "%Y-%m-%d").date(), datetime.strptime(str(dt_ecras)[:10], "%Y-%m-%d").date()
            age_calc = f"{(d_e - d_c).days} jours" if not date_ref else f"{(date_ref - d_c).days} j (Aujourd'hui)"
        except Exception:
            age_calc = str(ep.get("echeance", "-"))

    ref_p = str(ep.get("ref_controle") or "").strip()
    rep_s = str(ep.get("repere_eprouvette", "")).strip()
    rep_complet = f"{ref_p}{rep_s}" if ref_p else rep_s
    try:
        f_kn = float(ep.get("force_kn") or 0.0)
    except (ValueError, TypeError):
        f_kn = 0.0

    return {
        "ID": ep.get("id"),
        "Référence / Repère": rep_complet,
        "Type Essai": ep.get("type_essai", "Compression"),
        "N° BL": extraire_num_bl(ep),
        "Ouvrage": ep.get("ouvrage", "-"),
        "Classe Béton": ep.get("classe_beton", "-"),
        "Date Coulée": dt_coul,
        "Date Écrasement Prévue": dt_ecras,
        "Échéance Visée": ep.get("echeance", "-"),
        "Âge Théorique": age_calc,
        "Statut": "✅ Écrasée" if f_kn > 0 else "⏳ En attente"
    }


def executer_update_eprouvette(supabase, ep_id, update_payload):
    """
    Exécute une mise à jour d'éprouvette en retirant les clés obsolètes si elles
    ne figurent pas dans la structure de la table Supabase (ex: PGRST204 date_essai).
    """
    try:
        return supabase.table("suivi_controle_beton").update(update_payload).eq("id", ep_id).execute()
    except Exception as e:
        err_str = str(e)
        if "PGRST204" in err_str and "date_essai" in err_str:
            payload_clean = update_payload.copy()
            payload_clean.pop("date_essai", None)
            return supabase.table("suivi_controle_beton").update(payload_clean).eq("id", ep_id).execute()
        raise e


# =========================================================
# MODULE : PHASE 3 - VALIDATION ADMIN (PVs)
# =========================================================
def afficher_module_validation_admin(supabase, est_admin=False):
    st.subheader("🛡️ 3. Validation & Consultation des PVs")

    projet_id_actif = projets_config.projet_actif(st.session_state.get("user") or {})
    if not projet_id_actif:
        st.error("⚠️ Aucun projet ne vous est autorisé. Contactez un administrateur.")
        return

    if est_admin:
        st.info("💡 **Espace Administrateur BAALLAL** : vérifiez la conformité des écrasements et validez/signalez officiellement les PVs.")
    else:
        st.info("👁️ **Mode consultation** : cette phase est ouverte aux utilisateurs connectés. La validation officielle, le rejet, la signature et la modification des résultats sont réservés à l'administrateur BAALLAL.")

    try:
        res = supabase.table("suivi_controle_beton").select("*").eq("projet_id", projet_id_actif).not_.is_("force_kn", "null").gt("force_kn", 0).order("id", desc=True).execute()
        essais_realises = res.data or []
    except Exception as e:
        st.error(f"❌ Erreur de chargement des essais réalisés : {e}")
        return

    if not essais_realises:
        st.warning("ℹ️ Aucun essai écrasé n'est actuellement en attente de validation.")
        return

    lots_dict = {}
    for ep in essais_realises:
        b_id = ep.get("betonnage_id")
        if b_id not in lots_dict:
            lots_dict[b_id] = []
        lots_dict[b_id].append(ep)

    def _est_pv_deja_valide(statut):
        if not statut:
            return False
        s = (
            unicodedata.normalize("NFKD", str(statut).strip().lower())
            .encode("ascii", "ignore")
            .decode("ascii")
        ).strip()
        if any(mot in s for mot in ["rejet", "invalide", "non valide"]):
            return False
        return "valide" in s

    options_valid = []
    parents_dict_admin = obtenir_infos_betonnage_parents_bulk(supabase, list(lots_dict.keys()))
    for b_id, list_ep in lots_dict.items():
        info_b = parents_dict_admin.get(b_id, {})
        statut_lot = info_b.get("statut_pv", "⏳ En attente de validation")
        if _est_pv_deja_valide(statut_lot):
            continue
        ref_ctrl = determiner_ref_controle(supabase, b_id, info_b, list_ep[0])
        bl_num = extraire_num_bl(list_ep[0], info_b or {})
        label = f"Réf: {ref_ctrl} | BL: {bl_num} | Ouvrage: {list_ep[0].get('ouvrage', '-')} | Statut: {statut_lot}"
        options_valid.append((label, b_id, list_ep, info_b))

    if not options_valid:
        st.success("✅ Tous les PV en attente ont été traités — aucun PV ne reste à valider pour le moment.")
        return

    choix_label, b_id_sel, ep_sel_list, info_b_sel = st.selectbox(
        "📦 Choisir le PV / Lot à réviser :",
        options=options_valid,
        format_func=lambda x: x[0],
        key="select_pv_admin"
    )

    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    c1.metric("Nombre d'éprouvettes écrasées", len(ep_sel_list))
    c2.metric("Date de coulée", extraire_date_coulee(info_b_sel))
    c3.metric("Statut Actuel du PV", info_b_sel.get("statut_pv", "⏳ En attente"))

    df_key = f"admin_edit_pv_{b_id_sel}"
    if df_key not in st.session_state or st.session_state.get(f"{df_key}_len") != len(ep_sel_list):
        rows_val = []
        for ep in ep_sel_list:
            f_kn = float(ep.get("force_kn") or 0.0)
            t_essai = ep.get("type_essai", "Compression")
            forme_v = ep.get("forme", "Cylindrique 150x300")
            fc = float(ep.get("fc_mpa") or calculer_resistance_mpa(f_kn, t_essai, forme_v))
            rows_val.append({
                "ID": ep.get("id"),
                "Repère": ep.get("repere_eprouvette", "-"),
                "Type Essai": t_essai,
                "Échéance": ep.get("echeance", "-"),
                "Date Écrasement": ep.get("date_ecrasement", "-"),
                "Force (kN)": f_kn,
                "Résistance (MPa)": fc,
                "Opérateur": ep.get("technicien", "-"),
                "_forme": forme_v,
                "_force_orig": f_kn,
                "_fc_orig": fc,
            })
        st.session_state[df_key] = pd.DataFrame(rows_val)
        st.session_state[f"{df_key}_len"] = len(ep_sel_list)

    editor_key = f"editor_{df_key}"

    def _maj_resistance_admin():
        editor_state = st.session_state.get(editor_key, {})
        for row_idx, updated_cols in editor_state.get("edited_rows", {}).items():
            if "Force (kN)" in updated_cols or "Type Essai" in updated_cols:
                try:
                    new_force = float(updated_cols.get("Force (kN)", st.session_state[df_key].at[row_idx, "Force (kN)"]))
                except (ValueError, TypeError):
                    new_force = 0.0
                t_essai = updated_cols.get("Type Essai", st.session_state[df_key].at[row_idx, "Type Essai"])
                forme_v = st.session_state[df_key].at[row_idx, "_forme"]
                
                st.session_state[df_key].at[row_idx, "Force (kN)"] = new_force
                st.session_state[df_key].at[row_idx, "Type Essai"] = t_essai
                st.session_state[df_key].at[row_idx, "Résistance (MPa)"] = calculer_resistance_mpa(new_force, t_essai, forme_v)

    if est_admin:
        st.caption(
            "✏️ Mode administrateur : la **Force (kN)** et le **Type d'essai** sont modifiables"
            " ci-dessous — la Résistance (MPa) se recalcule automatiquement."
        )

    st.data_editor(
        st.session_state[df_key],
        column_config={
            "ID": None,
            "_forme": None,
            "_force_orig": None,
            "_fc_orig": None,
            "Repère": st.column_config.TextColumn("Repère", disabled=True),
            "Type Essai": st.column_config.SelectboxColumn("Type d'essai", options=["Compression", "Traction par fendage"], disabled=not est_admin),
            "Échéance": st.column_config.TextColumn("Échéance", disabled=True),
            "Date Écrasement": st.column_config.TextColumn("Date Écrasement", disabled=True),
            "Force (kN)": st.column_config.NumberColumn(
                "⚡ Force (kN)", disabled=not est_admin,
                min_value=0.0, max_value=3000.0, step=0.1, format="%.1f",
            ),
            "Résistance (MPa)": st.column_config.NumberColumn("Résistance (MPa)", disabled=True, format="%.2f"),
            "Opérateur": st.column_config.TextColumn("Opérateur", disabled=True),
        },
        use_container_width=True, hide_index=True, key=editor_key, on_change=_maj_resistance_admin,
    )

    st.markdown("---")
    if not est_admin:
        st.warning("🔐 **Validation officielle désactivée pour votre compte.** Seul l'administrateur **BAALLAL** peut enregistrer une décision, modifier les forces ou signer le PV.")
        st.markdown("### 📄 Statut du PV")
        st.write(info_b_sel.get("statut_pv", "⏳ En attente de validation"))
        if info_b_sel.get("visa_resp"):
            st.write(f"**Visa Responsable d'essai :** {info_b_sel.get('visa_resp')}")
        if info_b_sel.get("visa_chef"):
            st.write(f"**Visa Chef du laboratoire :** {info_b_sel.get('visa_chef')}")
    else:
        with st.form("form_valider_pv"):
            st.markdown("##### ✍️ Décision & Signatures Officielles")
            col_sig1, col_sig2 = st.columns(2)
            resp_essai = col_sig1.text_input("Visa Responsable d'essai", value=info_b_sel.get("visa_resp", "O.IKKEN"))
            chef_labo = col_sig2.text_input("Visa Chef du laboratoire", value=info_b_sel.get("visa_chef", "H.BAALLAL"))

            statut_decision = st.radio(
                "Décision d'approbation :",
                ["✅ Valider et Signer le PV", "⚠️ Remettre en Révision / Rejeter"],
                horizontal=True
            )
            comm_admin = st.text_area("Observations / Instructions complémentaires", value=info_b_sel.get("observations_admin", "Conforme aux spécifications NF EN 12390."))

            submit_val = st.form_submit_button("💾 Enregistrer la décision de validation", type="primary", use_container_width=True)

            if submit_val:
                nouveau_statut = "✅ Validé & Signé" if "Valider" in statut_decision else "❌ Rejeté / En Révision"
                update_payload = {
                    "statut_pv": str(nouveau_statut),
                    "visa_resp": str(resp_essai),
                    "visa_chef": str(chef_labo),
                    "observations_admin": str(comm_admin),
                    "date_validation": str(date.today())
                }
                try:
                    supabase.table("suivi_betonnage").update(update_payload).eq("id", int(b_id_sel)).execute()

                    df_edit = st.session_state.get(df_key)
                    if df_edit is not None:
                        for _, r in df_edit.iterrows():
                            ep_id_maj = int(r["ID"])
                            nouvelles_force = {
                                "force_kn": float(r["Force (kN)"]),
                                "type_essai": str(r["Type Essai"]),
                                "fc_mpa": float(r["Résistance (MPa)"]),
                            }
                            executer_update_eprouvette(supabase, ep_id_maj, nouvelles_force)

                    st.session_state.pop(df_key, None)
                    st.session_state.pop(f"{df_key}_len", None)
                    st.success(f"✅ Le statut du PV a été mis à jour avec succès : **{nouveau_statut}**")
                    st.balloons()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur lors de la mise à jour du statut dans Supabase : {e}")


# =========================================================
# 3. APPLICATION PRINCIPALE STREAMLIT
# =========================================================
def show(supabase):
    st.title("🧪 Contrôle & Écrasement du Béton (NF EN 12390)")

    st.session_state.setdefault("phase_actuelle", OPTIONS_ONGLETS[0])
    st.session_state.setdefault("nav_widget_seed", 0)

    user_info = st.session_state.get("user", {})
    role_user = str(st.session_state.get("user_role") or st.session_state.get("role") or user_info.get("role", "")).lower()
    can_edit = st.session_state.get("can_edit", False) or bool(user_info.get("can_edit", False))

    current_username = str(st.session_state.get("username") or user_info.get("username") or "").strip().upper()
    is_baallal_admin = current_username == "BAALLAL" and (role_user == "admin" or st.session_state.get("is_admin", False))

    projet_id_actif = projets_config.projet_actif(user_info)
    if not projet_id_actif:
        st.error("⚠️ Aucun projet ne vous est autorisé. Contactez un administrateur.")
        return
    st.caption(f"📁 Projet actif : **{projets_config.nom_projet(projet_id_actif)}**")

    mode_admin = False
    if is_baallal_admin:
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔒 Mode Administration / Édition")
        mode_admin = st.sidebar.checkbox("Activer le Mode Admin / Édition", value=False)

    widget_key = f"nav_phase_widget_{st.session_state['nav_widget_seed']}"
    try:
        onglet_courant = st.segmented_control("Navigation entre phases :", OPTIONS_ONGLETS, default=st.session_state["phase_actuelle"], key=widget_key)
    except AttributeError:
        onglet_courant = st.radio("Navigation entre phases :", OPTIONS_ONGLETS, index=OPTIONS_ONGLETS.index(st.session_state["phase_actuelle"]), horizontal=True, key=widget_key)

    st.session_state["phase_actuelle"] = onglet_courant

    betonnages_preleves = []
    try:
        res_b = supabase.table("suivi_betonnage").select("*").eq("projet_id", projet_id_actif).order("id", desc=True).execute()
        if res_b.data:
            betonnages_preleves = [b for b in res_b.data if b.get("prelevement") and str(b.get("prelevement")).upper().startswith("OUI")]
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des bétonnages : {e}")

    map_betonnages = {b.get("id"): b for b in betonnages_preleves}

    # =========================================================
    # PHASE 0 : RÉCEPTION & SAISIE DU NUMÉRO DE RÉCEPTION
    # =========================================================
    if onglet_courant == OPTIONS_ONGLETS[0]:
        st.subheader("📋 0. Réception & Validation des Bétons")
        st.info("💡 Saisissez le **N° Réception** pour débloquer la Phase 1 et générer les étiquettes QR Code.")

        if betonnages_preleves:
            rows_reception = []
            for item in betonnages_preleves:
                rows_reception.append({
                    "_id_beton": item.get("id"),
                    "1-Numero de reception": str(item.get("num_reception") or ""),
                    "2-Date de livraison": extraire_date_coulee(item),
                    "3-Nb d'éprouvettes": int(item.get("nb_eprouvettes") or 12),
                    "4-Classe de béton": item.get("classe_beton") or "-",
                    "5-Ouvrage": item.get("ouvrage") or "-",
                })

            df_edited = st.data_editor(
                pd.DataFrame(rows_reception),
                column_config={"_id_beton": None},
                use_container_width=True, hide_index=True, key="editor_reception_phase0",
            )

            if st.button("💾 Enregistrer les N° de Réception", type="primary"):
                try:
                    for _, row in df_edited.iterrows():
                        b_id, n_rec = int(row["_id_beton"]), str(row["1-Numero de reception"]).strip()
                        if n_rec and n_rec not in ["-", ""]:
                            supabase.table("suivi_betonnage").update({"num_reception": n_rec}).eq("id", b_id).execute()
                    st.success("✅ N° de Réception mis à jour !")
                    st.rerun()
                except Exception as err:
                    st.error(f"❌ Erreur lors de la mise à jour des N° de Réception : {err}")

    # =========================================================
    # PHASE 1 : PROGRAMMATION DES ÉCHÉANCES & TYPE D'ESSAI
    # =========================================================
    elif onglet_courant == OPTIONS_ONGLETS[1]:
        st.subheader("📅 1. Programmer les Échéances & Types d'Essai")

        if can_edit:
            with st.expander("✏️ Modification / Ajustement d'une Programmation Existante", expanded=False):
                try:
                    res_p = supabase.table("suivi_controle_beton").select("*").eq("projet_id", projet_id_actif).order("id", desc=True).execute()
                    eprouvettes_enregistrees = res_p.data or []
                except Exception as e:
                    eprouvettes_enregistrees = []

                if eprouvettes_enregistrees:
                    df_display_prog = pd.DataFrame(eprouvettes_enregistrees)
                    
                    if "type_essai" not in df_display_prog.columns:
                        df_display_prog["type_essai"] = "Compression"

                    cols_ed = [c for c in ["id", "ref_controle", "repere_eprouvette", "type_essai", "echeance", "date_ecrasement", "date_coulee", "ouvrage"] if c in df_display_prog.columns]

                    df_prog_modifiee = st.data_editor(
                        df_display_prog[cols_ed],
                        column_config={
                            "id": st.column_config.NumberColumn("ID", disabled=True),
                            "type_essai": st.column_config.SelectboxColumn("Type d'essai", options=["Compression", "Traction par fendage"]),
                            "echeance": st.column_config.SelectboxColumn("Échéance Visée", options=["3 jours", "7 jours", "28 jours", "90 jours"]),
                        },
                        use_container_width=True, hide_index=True, key="editor_modification_phase1",
                    )

                    if st.button("💾 Enregistrer les Modifications de Programmation", type="primary", key="btn_save_mod_prog"):
                        try:
                            for _, r_m in df_prog_modifiee.iterrows():
                                pay = {
                                    "ref_controle": str(r_m.get("ref_controle", "")).strip(),
                                    "repere_eprouvette": str(r_m.get("repere_eprouvette", "")).strip(),
                                    "type_essai": str(r_m.get("type_essai", "Compression")),
                                    "echeance": str(r_m.get("echeance", "")).strip(),
                                    "date_coulee": str(r_m.get("date_coulee", "")).strip(),
                                    "date_ecrasement": str(r_m.get("date_ecrasement", "")).strip(),
                                }
                                executer_update_eprouvette(supabase, int(r_m["id"]), pay)
                            st.success("✅ Modifications enregistrées avec succès !")
                            st.rerun()
                        except Exception as err:
                            st.error(f"❌ Erreur de mise à jour dans la base de données : {err}")

        st.divider()
        st.subheader("➕ Ajouter une Nouvelle Programmation")

        if betonnages_preleves:
            options_beton = {f"N° Réception: {b.get('num_reception')} | Ouvrage: {b.get('ouvrage')}": b for b in betonnages_preleves if b.get('num_reception')}
            if options_beton:
                choix_label_p = st.selectbox("Sélectionner la fiche de bétonnage :", list(options_beton.keys()), key="prog_beton_select")
                beton_p = options_beton[choix_label_p]
                b_id = beton_p.get("id")

                col_type, col_ech = st.columns(2)
                type_essai_p = col_type.selectbox("🧪 Type d'Essai", ["Compression", "Traction par fendage"], key=f"p_type_essai_{b_id}")
                echeance_p = col_ech.selectbox("Âge / Échéance visée", ["3 jours", "7 jours", "28 jours", "90 jours"], key=f"p_echeance_{b_id}")

                date_coulee_p = datetime.strptime(extraire_date_coulee(beton_p), "%Y-%m-%d").date()
                nb_j = extraire_nb_jours(echeance_p)
                date_ecrasement_prevue = date_coulee_p + timedelta(days=nb_j)

                nb_eprouvettes_p = st.number_input("Nombre d'éprouvettes", min_value=1, max_value=12, value=3, key=f"p_nb_ep_{b_id}")
                forme_p = st.selectbox("Forme", ["Cylindrique 150x300", "Cylindrique 160x320", "Cylindrique 100x200"], key=f"p_forme_{b_id}")

                if st.button("📌 Enregistrer la Programmation", type="primary", use_container_width=True):
                    try:
                        for i in range(int(nb_eprouvettes_p)):
                            pay = {
                                "betonnage_id": int(b_id),
                                "type_essai": str(type_essai_p),
                                "num_bl": str(extraire_num_bl(beton_p)),
                                "ouvrage": str(beton_p.get("ouvrage", "")),
                                "classe_beton": str(beton_p.get("classe_beton", "")),
                                "date_coulee": str(date_coulee_p),
                                "echeance": str(echeance_p),
                                "date_ecrasement": str(date_ecrasement_prevue),
                                "ref_controle": str(beton_p.get("num_reception", "")),
                                "repere_eprouvette": f"/{i+1}",
                                "forme": str(forme_p),
                                "projet_id": int(projet_id_actif),
                            }
                            supabase.table("suivi_controle_beton").insert(pay).execute()
                        st.success("✅ Éprouvettes programmées avec succès !")
                        st.rerun()
                    except Exception as err:
                        st.error(f"❌ Erreur lors de l'insertion dans Supabase : {err}")

    # =========================================================
    # PHASE 2 : PLANNING & SAISIE DES ÉCRASEMENTS
    # =========================================================
    elif onglet_courant == OPTIONS_ONGLETS[2]:
        st.subheader("💥 2. Planning des Échéances & Saisie des Écrasements")

        try:
            res_att = supabase.table("suivi_controle_beton").select("*").eq("projet_id", projet_id_actif).or_("force_kn.is.null,force_kn.eq.0").order("id", desc=False).execute()
            eprouvettes_en_attente = res_att.data or []
        except Exception as e:
            eprouvettes_en_attente = []

        if not eprouvettes_en_attente:
            st.info("👍 Aucune éprouvette en attente de saisie.")
        else:
            groupes_lots = {}
            for ep in eprouvettes_en_attente:
                ref_ctrl = ep.get("ref_controle", "N/A")
                t_essai = ep.get("type_essai", "Compression")
                cle_groupe = f"Réf: {ref_ctrl} | Type: {t_essai} | Échéance: {ep.get('echeance')} | Ouvrage: {ep.get('ouvrage')} (Lot #{ep.get('betonnage_id')})"
                groupes_lots.setdefault(cle_groupe, []).append(ep)

            choix_lot = st.selectbox("📦 Sélectionner le lot à écraser :", list(groupes_lots.keys()))
            lot_selected = groupes_lots[choix_lot]

            st.markdown("##### 📝 Saisie des Forces de Rupture (kN)")

            type_essai_lot = st.radio("Type d'essai pour ce lot :", ["Compression", "Traction par fendage"], index=0 if lot_selected[0].get("type_essai") != "Traction par fendage" else 1, horizontal=True)

            df_saisie_key = f"df_saisie_p2_{lot_selected[0].get('betonnage_id')}_{lot_selected[0].get('echeance')}"

            if df_saisie_key not in st.session_state or len(st.session_state[df_saisie_key]) != len(lot_selected):
                rows_saisie = []
                for ep in lot_selected:
                    f_kn = float(ep.get("force_kn") or 0.0)
                    forme_ep = ep.get("forme", "Cylindrique 150x300")
                    fc_calc = calculer_resistance_mpa(f_kn, type_essai_lot, forme_ep)
                    rows_saisie.append({
                        "ID": ep["id"],
                        "Repère": ep.get("repere_eprouvette"),
                        "Type Essai": type_essai_lot,
                        "Forme": forme_ep,
                        "Force (kN)": f_kn,
                        "Résistance (MPa)": fc_calc,
                    })
                st.session_state[df_saisie_key] = pd.DataFrame(rows_saisie)

            editor_p2_key = f"editor_p2_{df_saisie_key}"

            def _maj_resistance_phase2():
                editor_state = st.session_state.get(editor_p2_key, {})
                for row_idx, updated_cols in editor_state.get("edited_rows", {}).items():
                    if "Force (kN)" in updated_cols or "Type Essai" in updated_cols:
                        try:
                            new_force = float(updated_cols.get("Force (kN)", st.session_state[df_saisie_key].at[row_idx, "Force (kN)"]))
                        except (ValueError, TypeError):
                            new_force = 0.0
                        t_essai = type_essai_lot
                        forme_v = st.session_state[df_saisie_key].at[row_idx, "Forme"]

                        st.session_state[df_saisie_key].at[row_idx, "Force (kN)"] = new_force
                        st.session_state[df_saisie_key].at[row_idx, "Résistance (MPa)"] = calculer_resistance_mpa(new_force, t_essai, forme_v)

            st.data_editor(
                st.session_state[df_saisie_key],
                column_config={
                    "ID": None,
                    "Forme": None,
                    "Repère": st.column_config.TextColumn("Repère", disabled=True),
                    "Type Essai": st.column_config.TextColumn("Type d'essai", disabled=True),
                    "Force (kN)": st.column_config.NumberColumn("⚡ Force (kN)", min_value=0.0, max_value=3000.0, step=0.1, format="%.1f"),
                    "Résistance (MPa)": st.column_config.NumberColumn("Résistance (MPa)", disabled=True, format="%.2f"),
                },
                use_container_width=True,
                hide_index=True,
                key=editor_p2_key,
                on_change=_maj_resistance_phase2,
            )

            if st.button("💾 Enregistrer la Saisie du Lot", type="primary", use_container_width=True):
                try:
                    df_to_save = st.session_state[df_saisie_key]
                    for _, r in df_to_save.iterrows():
                        f_kn = float(r["Force (kN)"])
                        if f_kn > 0:
                            fc_val = float(calculer_resistance_mpa(f_kn, type_essai_lot, r["Forme"]))
                            pay_update = {
                                "force_kn": f_kn,
                                "fc_mpa": fc_val,
                                "type_essai": str(type_essai_lot),
                                "date_ecrasement": str(date.today()),
                                "technicien": str(st.session_state.get("username", "Technicien")),
                            }
                            executer_update_eprouvette(supabase, int(r["ID"]), pay_update)

                    st.session_state.pop(df_saisie_key, None)
                    st.success("✅ Résultats enregistrés avec succès !")
                    st.rerun()
                except Exception as err:
                    st.error(f"❌ Échec de l'enregistrement PostgREST / Supabase : {err}")

    # =========================================================
    # PHASE 3 : VALIDATION ADMIN (PVs)
    # =========================================================
    elif onglet_courant == OPTIONS_ONGLETS[3]:
        afficher_module_validation_admin(supabase, est_admin=is_baallal_admin)
