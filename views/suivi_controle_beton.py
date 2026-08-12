from datetime import date, datetime, timedelta
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st


def generer_pv_excel(export_data, infos_header):
    """Génère un Procès-Verbal (PV) d'écrasement de béton répliquant le modèle LPEE / CTR CSB."""
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
    font_title = Font(name="Calibri", size=13, bold=True)
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

    # =========================================================
    # 1. ENTÊTE DU LABORATOIRE ET REFERENCES
    # =========================================================
    # Logo / Nom Laboratoire (Colonnes A-D)
    ws.merge_cells("A1:D1")
    ws["A1"] = "LPEE / CTR CSB"
    ws["A1"].font = font_bold
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:D3")
    ws["A2"] = "Laboratoire de Contrôle Externe"
    ws["A2"].font = font_bold
    ws["A2"].alignment = align_center

    # Informations dossier (Colonnes E-H)
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

    # Applicateur de bordure pour l'en-tête
    for r in range(1, 4):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # =========================================================
    # 2. TITRE DE L'ESSAI ET TYPE D'ESSAI
    # =========================================================
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
    ws["E5"] = "[  ] TRACTION PAR FENDAGE NF EN 12390-6 (2019)"
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

    # =========================================================
    # 3. FICHE TECHNIQUE DE PRELEVEMENT ET CHANTIER
    # =========================================================
    # Date de prélèvement
    ws["A7"] = "Date de\nprélèvement"
    ws["A7"].font = font_bold
    ws["A7"].alignment = align_center
    ws["B7"] = str(infos_header.get("date_coulee", "02/06/2025"))
    ws["B7"].font = font_bold
    ws["B7"].alignment = align_center

    # Lieu de prélèvement
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

    # Chantier
    ws.merge_cells("A8:A9")
    ws["A8"] = "Chantier"
    ws["A8"].font = font_bold
    ws["A8"].alignment = align_center

    ws.merge_cells("B8:D9")
    ws["B8"] = infos_header.get(
        "chantier",
        "Augmentation de la capacité ferroviaire entre Kenitra et Marrakech et au niveau du hub de Casablanca\n--- Travaux d'exécution de terrassement, ouvrages d'art et rétablissement de communication",
    )
    ws["B8"].font = font_small
    ws["B8"].alignment = align_center

    # Type de Béton
    ws.merge_cells("E8:F8")
    ws["E8"] = "Type de béton"
    ws["E8"].font = font_bold
    ws["E8"].alignment = align_center

    ws.merge_cells("G8:H8")
    ws["G8"] = infos_header.get("classe_beton", "C30/37")
    ws["G8"].font = font_bold
    ws["G8"].alignment = align_center

    # Section Éprouvettes Header
    ws.merge_cells("E9:H9")
    ws["E9"] = "EPROUVETTES"
    ws["E9"].font = font_bold
    ws["E9"].alignment = align_center

    # Ligne Fournisseur / Caractéristiques Physiques
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

    # Affaissement & Confection
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

    # Température & Conservation
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
        "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à"
        " 20°C ± 2°C"
    )
    ws["E12"].font = font_bold
    ws["E12"].alignment = align_center

    # Densité & Bon de Livraison
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

    # Application des bordures des fiches techniques
    for r in range(7, 14):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # =========================================================
    # 4. TABLEAU DES RÉSULTATS D'ÉCRASEMENT
    # =========================================================
    # En-tête Tableau
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

    # Remplissage des données
    row_start = 16
    ws.merge_cells(f"A{row_start}:A{row_start + len(export_data) - 1}")
    ws[f"A{row_start}"] = "B/01"
    ws[f"A{row_start}"].font = font_bold
    ws[f"A{row_start}"].alignment = align_center

    ws.merge_cells(f"B{row_start}:B{row_start + len(export_data) - 1}")
    ws[f"B{row_start}"] = str(infos_header.get("date_coulee", "02/06/2025"))
    ws[f"B{row_start}"].font = font_bold
    ws[f"B{row_start}"].alignment = align_center

    for idx, item in enumerate(export_data):
        curr_row = row_start + idx

        # Repère Éprouvette (ex: /1, /2, ...)
        repere = item.get("repere_eprouvette", f"/{idx + 1}")
        if not str(repere).startswith("/"):
            repere = f"/{idx + 1}"

        # Écriture des lignes
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

    # Exemple de fusion dynamique pour la moyenne (si 3 éprouvettes par âge)
    nb_total = len(export_data)
    if nb_total >= 3:
        # Fusion pour le bloc 7 jours
        ws.merge_cells(f"C16:C18")
        ws.merge_cells(f"D16:D18")
        ws.merge_cells(f"H16:H18")
        ws["H16"] = f"=ROUND(AVERAGE(F16:F18),1)"
        ws["H16"].alignment = align_center
        ws["H16"].font = font_bold

    if nb_total >= 12:
        # Fusion pour le bloc 28 jours (éprouvettes 4 à 12)
        ws.merge_cells(f"C19:C27")
        ws.merge_cells(f"D19:D27")

        ws.merge_cells(f"H19:H21")
        ws["H19"] = f"=ROUND(AVERAGE(F19:F21),1)"
        ws["H19"].alignment = align_center
        ws["H19"].font = font_bold

        ws.merge_cells(f"H22:H27")
        ws["H22"] = f"=ROUND(AVERAGE(F22:F27),1)"
        ws["H22"].alignment = align_center
        ws["H22"].font = font_bold

    # =========================================================
    # 5. PIED DE PAGE ET COMMENTAIRES
    # =========================================================
    last_row = row_start + len(export_data)

    ws.cell(row=last_row, column=1, value="COMMENTAIRE :").font = font_bold
    ws.cell(row=last_row, column=1).alignment = align_center

    ws.merge_cells(
        start_row=last_row, start_column=2, end_row=last_row, end_column=8
    )
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

    # LARGEURS COLONNES HARMONISÉES FORMAT A4
    col_widths = {
        "A": 8,  # Réf
        "B": 12,  # Date Fabri
        "C": 12,  # Date Essai
        "D": 10,  # Age
        "E": 18,  # Charge KN
        "F": 14,  # Fc Compression
        "G": 12,  # Traction
        "H": 12,  # Moyenne
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
