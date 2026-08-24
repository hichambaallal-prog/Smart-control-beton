import io
import re
from datetime import date, datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st


# ==============================================================================
# 1. GESTION DES UTILISATEURS ET CONNEXION SUPABASE
# ==============================================================================
def connecter_utilisateur(supabase, nom_utilisateur, mot_de_passe):
    """
    Vérifie le nom d'utilisateur, le mot de passe et récupère le champ 'can_edit'
    depuis la table 'users' de Supabase.
    """
    try:
        res = (
            supabase.table("users")
            .select("*")
            .eq("username", nom_utilisateur)
            .eq("password", mot_de_passe)
            .execute()
        )

        if res.data:
            user_info = res.data[0]
            st.session_state["user_logged"] = True
            st.session_state["user"] = user_info
            st.session_state["username"] = user_info.get("username")
            st.session_state["role"] = user_info.get("role")
            st.session_state["user_role"] = user_info.get("role")
            st.session_state["can_edit"] = bool(user_info.get("can_edit", False))
            return True
        else:
            st.error("Nom d'utilisateur ou mot de passe incorrect.")
            return False
    except Exception as e:
        st.error(f"Erreur lors de la connexion : {e}")
        return False


# =========================================================
# FONCTION UTILITAIRE : VÉRIFICATION DES DOUBLONS DU N° DE RÉCEPTION
# =========================================================
def verifier_doublon_num_reception(supabase, num_reception, current_beton_id=None):
    """
    Vérifie dans la table 'suivi_betonnage' si le num_reception existe déjà.
    Retourne True si le numéro est en doublon, sinon False.
    """
    if not num_reception or str(num_reception).strip() in ["", "-", "None", "NaN", "N/A"]:
        return False
        
    num_clean = str(num_reception).strip()
    try:
        # Recherche par num_reception
        res1 = (
            supabase.table("suivi_betonnage")
            .select("id, num_reception")
            .eq("num_reception", num_clean)
            .execute()
        )
        # Recherche par num_reception si applicable
        res2 = (
            supabase.table("suivi_betonnage")
            .select("id, num_reception")
            .eq("num_reception", num_clean)
            .execute()
        )
        
        matches = (res1.data or []) + (res2.data or [])
        
        for m in matches:
            # S'il existe un enregistrement ayant le même numéro mais un ID différent
            if current_beton_id is None or int(m.get("id")) != int(current_beton_id):
                return True
    except Exception as e:
        st.warning(f"Note lors de la vérification des doublons : {e}")
        
    return False


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
# 2. GÉNÉRATION DU PROCÈS-VERBAL EXCEL (FORMAT EXACT LPEE)
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
    
    # ----------- MODIFICATIONS COLONNES F, G et H -----------
    ws.merge_cells("F1:G1")
    ws["F1"] = remplacer_na(infos_header.get("re_num"), "25/260/LGV/ B/")
    ws["F1"].font = font_regular
    
    ws["H1"] = "BETON"
    ws["H1"].font = font_bold
    ws["H1"].alignment = align_center
    # --------------------------------------------------------

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
# FONCTION UTILITAIRE : EXPORT EXCEL
# =========================================================
def exporter_dataframe_excel(df, date_chaine):
    """Génère un fichier Excel à partir du DataFrame."""
    buffer = io.BytesIO()
    nom_feuille = f"Planning_{date_chaine}"[:31]  # Limitation max 31 caractères pour Excel
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nom_feuille)

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
    """Détermine la référence de contrôle avec priorité donnée au N° de Réception."""
    session_key = f"ref_controle_beton_{betonnage_id}"
    if session_key in st.session_state and st.session_state[session_key]:
        return st.session_state[session_key]

    # --- Priorité au Numéro de Réception ---
    num_reception = None
    if info_betonnage:
        num_reception = info_betonnage.get("num_reception") or info_betonnage.get("num_reception")
    
    if num_reception and str(num_reception).strip() not in ["", "-", "None", "NaN", "N/A"]:
        val_defaut = str(num_reception).strip()
        st.session_state[session_key] = val_defaut
        return val_defaut

    ref_parent = info_betonnage.get("ref_controle") if info_betonnage else None
    if ref_parent and str(ref_parent).strip():
        st.session_state[session_key] = str(ref_parent).strip()
        return str(ref_parent).strip()

    ref_ep = sample_ep.get("ref_controle") if sample_ep else None
    if ref_ep and str(ref_ep).strip():
        st.session_state[session_key] = str(ref_ep).strip()
        return str(ref_ep).strip()

    ouvrage = info_betonnage.get('ouvrage', 'N/A') if info_betonnage else 'N/A'
    defaut = f"REF-{betonnage_id}-{ouvrage}"
    st.session_state[session_key] = defaut
    return defaut


# =========================================================
# 3. APPLICATION PRINCIPALE STREAMLIT
# =========================================================
def show(supabase):
    st.title("📋 Historique & Procès-Verbaux d'Écrasement (NF EN 12390)")

    user_info = st.session_state.get("user", {})
    role_utilisateur = str(
        st.session_state.get("user_role")
        or st.session_state.get("role")
        or user_info.get("role", "")
    ).lower()

    can_edit = st.session_state.get("can_edit", False) or bool(user_info.get("can_edit", False))

    roles_autorises = ["laboratoire", "labo", "admin", "responsable_labo", "qualite"]

    if role_utilisateur not in roles_autorises and not st.session_state.get("is_admin", False) and not can_edit:
        st.error("⛔ **Accès Restreint**")
        st.warning(
            "Ce module est réservé exclusivement au personnel du **Laboratoire de Contrôle**."
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
        if res_all.data:
            df_all = pd.DataFrame(res_all.data)

            df_valides = df_all[(df_all["force_kn"].notnull()) & (df_all["force_kn"] > 0)].copy()

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

                info_beton_h = obtenir_infos_betonnage_parent(supabase, b_id_h)
                tous_essais_hist = obtenir_historique_betonnage(supabase, b_id_h)

                export_data_h = []
                items_a_exporter = tous_essais_hist if tous_essais_hist else lot_hist

                for item in items_a_exporter:
                    sec = float(item.get("section") or 176.71)
                    f_kn = float(item.get("force_kn") or 0.0)
                    fc = float(item.get("fc_mpa") or (round((f_kn * 10.0) / sec, 1) if f_kn > 0 else 0.0))

                    ref_p = str(item.get("ref_controle") or "").strip()
                    rep_s = str(item.get("repere_eprouvette", f"/{item['id']}")).strip()
                    rep_c = f"{ref_p}{rep_s}" if ref_p else rep_s
                    statut = "En cours" if f_kn == 0 else "Réalisé"

                    export_data_h.append({
                        "repere_eprouvette": rep_c,
                        "forme": item.get("forme", "Cylindrique 150x300"),
                        "section": sec,
                        "force_kn": f_kn,
                        "fc_mpa": fc,
                        "date_essai": item.get("date_ecrasement", "-"),
                        "age": str(item.get("echeance", "28")).replace(" jours", "").replace("j", ""),
                        "statut": statut,
                    })

                num_bl_h = extraire_num_bl(sample_h, info_beton_h or {}, choix_pv_hist)
                aff_h = (info_beton_h.get("affaissement") or info_beton_h.get("slump")) if info_beton_h else None
                temp_h = (info_beton_h.get("temperature") or info_beton_h.get("temp_beton")) if info_beton_h else None
                ouv_h = (info_beton_h.get("ouvrage") if info_beton_h else None) or sample_h.get("ouvrage")
                date_coulee_h = (info_beton_h.get("date_coulee") if info_beton_h else None) or sample_h.get("date_coulee")
                centrale_h = (info_beton_h.get("centrale") or info_beton_h.get("centrale_beton")) if info_beton_h else sample_h.get("centrale")
                tech_prelev_h = (
                    (info_beton_h.get("technicien_prelevement") or info_beton_h.get("preleve_par") or info_beton_h.get("technicien"))
                    if info_beton_h
                    else sample_h.get("technicien")
                )

                infos_header_h = {
                    "re_num": "25/260/LGV/ B/",
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
                    "observations": sample_h.get("observations", "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"),
                    "technicien_prelevement": tech_prelev_h,
                }

                excel_pv_hist = generer_pv_excel(export_data_h, infos_header_h)
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

            df_all.rename(columns=renommage_colonnes, inplace=True)

            colonnes_ordre = [
                "id",
                "betonnage_id",
                "ref_controle",
                "repere_eprouvette",
                "num_bl",
                "ouvrage",
                "classe_beton",
                "date_coulee",
                "affaissement_mm",
                "temp_beton_C",
                "echeance",
                "date_ecrasement",
                "forme",
                "section",
                "force_kn",
                "fc_mpa",
                "technicien",
                "observations"
            ]

            cols_existantes = [col for col in colonnes_ordre if col in df_all.columns]
            cols_restantes = [col for col in df_all.columns if col not in cols_existantes]
            df_final = df_all[cols_existantes + cols_restantes]

            # ==========================================
            # AJOUT DES FILTRES DE RECHERCHE
            # ==========================================
            col_search1, col_search2 = st.columns(2)
            
            with col_search1:
                search_ref = st.text_input("🔍 Recherche par Réf. Contrôle", placeholder="Ex: REF-123-GARE CASA SUD")
                
            with col_search2:
                search_date = st.text_input("📅 Recherche par Date de coulée", placeholder="Ex: 2026-08-24")

            # Application des filtres si les champs ne sont pas vides
            if search_ref:
                df_final = df_final[df_final["ref_controle"].astype(str).str.contains(search_ref, case=False, na=False)]
                
            if search_date:
                df_final = df_final[df_final["date_coulee"].astype(str).str.contains(search_date, case=False, na=False)]
            # ==========================================

            st.dataframe(df_final, use_container_width=True, hide_index=True)

            excel_historique = exporter_dataframe_excel(df_final, "Historique_Global")
            st.download_button(
                label="📊 Télécharger la base de données globale (Excel)",
                data=excel_historique,
                file_name=f"Historique_Global_Beton_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_download_hist_global",
            )
        else:
            st.info("ℹ️ Aucun historique disponible dans la base de données.")
            
    except Exception as e:
        st.error(f"Erreur lors du chargement de l'historique global : {e}")
