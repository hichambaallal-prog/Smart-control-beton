import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import re

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# UTILITAIRE D'EXTRACTION NUMÉRIQUE
# =========================================================

def extract_numeric(val):
    if pd.isna(val) or val is None:
        return None
    val_str = str(val).replace(',', '.')
    match = re.search(r"[-+]?\d*\.\d+|\d+", val_str)
    if match:
        num = float(match.group())
        return int(num) if num.is_integer() else num
    return None


# =========================================================
# 1. GENERATION EXCEL (CONFORME LPEE - LGV CASA SUD)
# =========================================================

def generate_excel_synthesis_betonnage(df_data, titre_periode):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Synthèse Bétonnage"

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

    color_primary = "1F4E79"
    color_header = "2D572C"
    color_card_bg = "F7F9FA"
    color_kpi_bg = "EDF2F8"

    font_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    font_section = Font(name="Calibri", size=13, bold=True, color=color_primary)
    font_bold = Font(name="Calibri", size=12, bold=True)
    font_normal = Font(name="Calibri", size=12)
    font_th = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    font_kpi_val = Font(name="Calibri", size=14, bold=True, color=color_primary)

    fill_title = PatternFill(start_color=color_primary, end_color=color_primary, fill_type="solid")
    fill_th = PatternFill(start_color=color_header, end_color=color_header, fill_type="solid")
    fill_card = PatternFill(start_color=color_card_bg, end_color=color_card_bg, fill_type="solid")
    fill_kpi = PatternFill(start_color=color_kpi_bg, end_color=color_kpi_bg, fill_type="solid")

    thin_border_side = Side(style='thin', color='B0C4DE')
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    nb_cols = max(len(df_data.columns), 6)
    last_col_letter = get_column_letter(nb_cols)
    mid_col_idx = nb_cols // 2
    mid_col_letter = get_column_letter(mid_col_idx)
    next_mid_letter = get_column_letter(mid_col_idx + 1)

    ws.merge_cells(f"A1:{last_col_letter}2")
    cell_title = ws["A1"]
    cell_title.value = "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - CTR-CSB\nRAPPORT DE SYNTHÈSE DU BÉTONNAGE"
    cell_title.font = font_title
    cell_title.fill = fill_title
    cell_title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 28

    ws.merge_cells(f"A4:{mid_col_letter}4")
    ws["A4"].value = "   CLIENT :   TGCC"
    ws["A4"].font = font_bold
    ws["A4"].fill = fill_card
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"{next_mid_letter}4:{last_col_letter}4")
    ws[f"{next_mid_letter}4"].value = "   PROJET :   LGV CASA SUD"
    ws[f"{next_mid_letter}4"].font = font_bold
    ws[f"{next_mid_letter}4"].fill = fill_card
    ws[f"{next_mid_letter}4"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"A5:{mid_col_letter}5")
    ws["A5"].value = f"   PÉRIODE :   {titre_periode}"
    ws["A5"].font = font_bold
    ws["A5"].fill = fill_card
    ws["A5"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"{next_mid_letter}5:{last_col_letter}5")
    ws[f"{next_mid_letter}5"].value = f"   DATE ÉDITION :   {datetime.now().strftime('%d/%m/%Y')}"
    ws[f"{next_mid_letter}5"].font = font_bold
    ws[f"{next_mid_letter}5"].fill = fill_card
    ws[f"{next_mid_letter}5"].alignment = Alignment(horizontal="left", vertical="center")

    for r in range(4, 6):
        ws.row_dimensions[r].height = 32
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    row_idx = 7
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📊 RÉSUMÉ GLOBAL"
    ws[f"A{row_idx}"].font = font_section
    ws.row_dimensions[row_idx].height = 30

    row_idx += 1
    vol_tot = df_data["Quantité (m³)"].sum() if "Quantité (m³)" in df_data.columns else 0

    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"].value = "Volume Total Béton"
    ws[f"A{row_idx}"].font = font_bold
    ws[f"A{row_idx}"].fill = fill_kpi
    ws[f"A{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"A{row_idx+1}:{last_col_letter}{row_idx+1}")
    ws[f"A{row_idx+1}"].value = f"{vol_tot:.1f} m³"
    ws[f"A{row_idx+1}"].font = font_kpi_val
    ws[f"A{row_idx+1}"].fill = fill_kpi
    ws[f"A{row_idx+1}"].alignment = Alignment(horizontal="center", vertical="center")

    for r in range(row_idx, row_idx + 2):
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    ws.row_dimensions[row_idx].height = 28
    ws.row_dimensions[row_idx+1].height = 36
    row_idx += 3

    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📋 DÉTAIL DES CONTRÔLES"
    ws[f"A{row_idx}"].font = font_section
    ws.row_dimensions[row_idx].height = 30
    row_idx += 1

    headers = list(df_data.columns)
    for col_num, h_name in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_num)
        cell.value = str(h_name)
        cell.font = font_th
        cell.fill = fill_th
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[row_idx].height = 42
    row_idx += 1

    start_data_row = row_idx
    for row_data in df_data.itertuples(index=False):
        for col_num, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_num)
            cell.value = val if pd.notna(val) else ""
            cell.font = font_normal
            cell.border = thin_border
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        ws.row_dimensions[row_idx].height = 36
        row_idx += 1

    end_data_row = row_idx - 1
    ws.row_dimensions[row_idx].height = 38
    total_cell = ws.cell(row=row_idx, column=1)
    total_cell.value = "TOTAL"
    total_cell.font = font_bold
    total_cell.border = total_border

    for col_num in range(1, len(headers) + 1):
        c = ws.cell(row=row_idx, column=col_num)
        c.border = total_border
        c.font = font_bold
        col_name = headers[col_num - 1]
        col_ltr = get_column_letter(col_num)
        if col_name == "Quantité (m³)":
            c.value = f"=SUM({col_ltr}{start_data_row}:{col_ltr}{end_data_row})"
            c.number_format = '0.0 "m³"'
            c.alignment = Alignment(horizontal="right", vertical="center")

    row_idx += 3

    ws.merge_cells(f"A{row_idx}:{mid_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "Responsables d'essai"
    ws[f"A{row_idx}"].font = font_bold
    ws[f"A{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"{next_mid_letter}{row_idx}:{last_col_letter}{row_idx}")
    ws[f"{next_mid_letter}{row_idx}"] = "Chef du Laboratoire"
    ws[f"{next_mid_letter}{row_idx}"].font = font_bold
    ws[f"{next_mid_letter}{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_idx].height = 30

    row_idx += 1
    ws.merge_cells(f"A{row_idx}:{mid_col_letter}{row_idx+3}")
    ws[f"A{row_idx}"] = "Visa & Signature :"
    ws[f"A{row_idx}"].font = font_normal
    ws[f"A{row_idx}"].alignment = Alignment(horizontal="left", vertical="top")

    ws.merge_cells(f"{next_mid_letter}{row_idx}:{last_col_letter}{row_idx+3}")
    ws[f"{next_mid_letter}{row_idx}"] = "Visa & Signature :"
    ws[f"{next_mid_letter}{row_idx}"].font = font_normal
    ws[f"{next_mid_letter}{row_idx}"].alignment = Alignment(horizontal="left", vertical="top")

    for r in range(row_idx, row_idx + 4):
        ws.row_dimensions[r].height = 22

    for r in range(row_idx - 1, row_idx + 4):
        for c in range(1, mid_col_idx + 1):
            ws.cell(row=r, column=c).border = thin_border
        for c in range(mid_col_idx + 1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    col_width_map = {
        "Date Livraison": 16, "Heure d'arrivée": 15, "N° BL": 16, "Ouvrage": 22,
        "Quantité (m³)": 16, "Classe": 14, "Durée de transport": 18, "Temp. Béton": 15,
        "Temp. Ambiante": 16, "Affaissement": 15, "Prélèvement": 18, "Météo": 15
    }

    for col_idx, col_name in enumerate(headers, 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width_map.get(col_name, 16)

    wb.save(output)
    output.seek(0)
    return output.getvalue()


def generate_excel_synthesis_controle(df_data, titre_periode):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Synthèse Contrôle Béton"

    # Configuration A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

    color_primary = "1F4E79"
    color_header = "2D572C"
    color_card_bg = "F7F9FA"
    color_kpi_bg = "EDF2F8"
    color_stat_hdr = "1F4E79"

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_section = Font(name="Calibri", size=12, bold=True, color=color_primary)
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_normal = Font(name="Calibri", size=11)
    font_th = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_kpi_val = Font(name="Calibri", size=13, bold=True, color=color_primary)

    fill_title = PatternFill(start_color=color_primary, end_color=color_primary, fill_type="solid")
    fill_th = PatternFill(start_color=color_header, end_color=color_header, fill_type="solid")
    fill_card = PatternFill(start_color=color_card_bg, end_color=color_card_bg, fill_type="solid")
    fill_kpi = PatternFill(start_color=color_kpi_bg, end_color=color_kpi_bg, fill_type="solid")
    fill_stat_hdr = PatternFill(start_color=color_stat_hdr, end_color=color_stat_hdr, fill_type="solid")

    thin_border_side = Side(style='thin', color='B0C4DE')
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    nb_cols = max(len(df_data.columns), 9)
    last_col_letter = get_column_letter(nb_cols)
    mid_col_idx = nb_cols // 2
    mid_col_letter = get_column_letter(mid_col_idx)
    next_mid_letter = get_column_letter(mid_col_idx + 1)

    # Entête - Titre
    ws.merge_cells(f"A1:{last_col_letter}2")
    ws["A1"].value = "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - CTR-CSB\nRAPPORT DE SYNTHÈSE DU CONTRÔLE BÉTON"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 25

    ws.merge_cells(f"A4:{mid_col_letter}4")
    ws["A4"].value = "   CLIENT :   TGCC"
    ws["A4"].font = font_bold
    ws["A4"].fill = fill_card
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"{next_mid_letter}4:{last_col_letter}4")
    ws[f"{next_mid_letter}4"].value = "   PROJET :   LGV CASA SUD"
    ws[f"{next_mid_letter}4"].font = font_bold
    ws[f"{next_mid_letter}4"].fill = fill_card
    ws[f"{next_mid_letter}4"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"A5:{mid_col_letter}5")
    ws["A5"].value = f"   PÉRIODE :   {titre_periode}"
    ws["A5"].font = font_bold
    ws["A5"].fill = fill_card
    ws["A5"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"{next_mid_letter}5:{last_col_letter}5")
    ws[f"{next_mid_letter}5"].value = f"   DATE ÉDITION :   {datetime.now().strftime('%d/%m/%Y')}"
    ws[f"{next_mid_letter}5"].font = font_bold
    ws[f"{next_mid_letter}5"].fill = fill_card
    ws[f"{next_mid_letter}5"].alignment = Alignment(horizontal="left", vertical="center")

    for r in range(4, 6):
        ws.row_dimensions[r].height = 28
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    # Résumé
    row_idx = 7
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📊 RÉSUMÉ GLOBAL"
    ws[f"A{row_idx}"].font = font_section

    row_idx += 1
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"].value = "Nombre Total de Prélèvements Contrôlés"
    ws[f"A{row_idx}"].font = font_bold
    ws[f"A{row_idx}"].fill = fill_kpi
    ws[f"A{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"A{row_idx+1}:{last_col_letter}{row_idx+1}")
    ws[f"A{row_idx+1}"].value = f"{len(df_data)} prélèvement(s)"
    ws[f"A{row_idx+1}"].font = font_kpi_val
    ws[f"A{row_idx+1}"].fill = fill_kpi
    ws[f"A{row_idx+1}"].alignment = Alignment(horizontal="center", vertical="center")

    for r in range(row_idx, row_idx + 2):
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    # Tableau Données
    row_idx += 3
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📋 MOYENNE DES ÉCRASEMENTS PAR ÉCHÉANCE"
    ws[f"A{row_idx}"].font = font_section
    row_idx += 1

    headers = list(df_data.columns)
    for col_num, h_name in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_num)
        cell.value = str(h_name)
        cell.font = font_th
        cell.fill = fill_th
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[row_idx].height = 35
    row_idx += 1

    start_data_row = row_idx
    for row_data in df_data.itertuples(index=False):
        for col_num, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_num)
            cell.value = val if pd.notna(val) else ""
            cell.font = font_normal
            cell.border = thin_border
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row_idx].height = 28
        row_idx += 1
    end_data_row = row_idx - 1

    # =========================================================
    # TABLEAU DE SYNTHÈSE STATISTIQUE DANS EXCEL
    # =========================================================
    row_idx += 2
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📈 SYNTHÈSE STATISTIQUE DES PARAMÈTRES"
    ws[f"A{row_idx}"].font = font_section
    row_idx += 1

    stat_headers = ["Indicateur", "Affaissement (mm)", "Temp. Béton (°C)", "Fc (MPa) [7 Jours]", "Fc (MPa) [28 Jours]"]
    
    for idx, h in enumerate(stat_headers):
        col_start = 1 if idx == 0 else 2 + (idx - 1) * 2
        col_end = 1 if idx == 0 else col_start + 1
        
        if col_start == col_end:
            c = ws.cell(row=row_idx, column=col_start)
            c.value = h
        else:
            ws.merge_cells(start_row=row_idx, start_column=col_start, end_row=row_idx, end_column=col_end)
            c = ws.cell(row=row_idx, column=col_start)
            c.value = h
            
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = fill_stat_hdr
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for c_i in range(1, nb_cols + 1):
        ws.cell(row=row_idx, column=c_i).border = thin_border

    ws.row_dimensions[row_idx].height = 25
    row_idx += 1

    col_map_excel = {
        "aff": "E",
        "temp": "F",
        "fc7": "H",
        "fc28": "I"
    }

    # Calcul dynamique des lignes dans la table statistique
    row_moy = row_idx + 1
    row_std = row_idx + 3

    if start_data_row <= end_data_row:
        f_min_aff = f"=MIN({col_map_excel['aff']}{start_data_row}:{col_map_excel['aff']}{end_data_row})"
        f_moy_aff = f"=AVERAGE({col_map_excel['aff']}{start_data_row}:{col_map_excel['aff']}{end_data_row})"
        f_max_aff = f"=MAX({col_map_excel['aff']}{start_data_row}:{col_map_excel['aff']}{end_data_row})"

        f_min_temp = f"=MIN({col_map_excel['temp']}{start_data_row}:{col_map_excel['temp']}{end_data_row})"
        f_moy_temp = f"=AVERAGE({col_map_excel['temp']}{start_data_row}:{col_map_excel['temp']}{end_data_row})"
        f_max_temp = f"=MAX({col_map_excel['temp']}{start_data_row}:{col_map_excel['temp']}{end_data_row})"

        f_min_fc7 = f"=MIN({col_map_excel['fc7']}{start_data_row}:{col_map_excel['fc7']}{end_data_row})"
        f_moy_fc7 = f"=AVERAGE({col_map_excel['fc7']}{start_data_row}:{col_map_excel['fc7']}{end_data_row})"
        f_max_fc7 = f"=MAX({col_map_excel['fc7']}{start_data_row}:{col_map_excel['fc7']}{end_data_row})"

        f_min_fc28 = f"=MIN({col_map_excel['fc28']}{start_data_row}:{col_map_excel['fc28']}{end_data_row})"
        f_moy_fc28 = f"=AVERAGE({col_map_excel['fc28']}{start_data_row}:{col_map_excel['fc28']}{end_data_row})"
        f_max_fc28 = f"=MAX({col_map_excel['fc28']}{start_data_row}:{col_map_excel['fc28']}{end_data_row})"

        # CORRECTION : Formule ECARTYPE.STANDARD et liaison directe CV% sur les cellules MOY et σ
        f_std_fc28 = f"=ECARTYPE.STANDARD({col_map_excel['fc28']}{start_data_row}:{col_map_excel['fc28']}{end_data_row})"
        f_cv_fc28 = f"=SIERREUR(({col_map_excel['fc28']}{row_std}/{col_map_excel['fc28']}{row_moy})*100, 0)"
    else:
        f_min_aff = f_moy_aff = f_max_aff = "-"
        f_min_temp = f_moy_temp = f_max_temp = "-"
        f_min_fc7 = f_moy_fc7 = f_max_fc7 = "-"
        f_min_fc28 = f_moy_fc28 = f_max_fc28 = f_std_fc28 = f_cv_fc28 = "-"

    stat_rows = [
        ("MIN", f_min_aff, f_min_temp, f_min_fc7, f_min_fc28),
        ("MOY", f_moy_aff, f_moy_temp, f_moy_fc7, f_moy_fc28),
        ("MAX", f_max_aff, f_max_temp, f_max_fc7, f_max_fc28),
        ("σ", "-", "-", "-", f_std_fc28),
        ("CV %", "-", "-", "-", f_cv_fc28)
    ]

    for label, v_aff, v_temp, v_fc7, v_fc28 in stat_rows:
        ws.cell(row=row_idx, column=1, value=label).font = font_bold
        ws.cell(row=row_idx, column=1).alignment = Alignment(horizontal="left", vertical="center")

        vals = [v_aff, v_temp, v_fc7, v_fc28]
        for idx, val in enumerate(vals):
            c_start = 2 + idx * 2
            c_end = c_start + 1
            ws.merge_cells(start_row=row_idx, start_column=c_start, end_row=row_idx, end_column=c_end)
            c = ws.cell(row=row_idx, column=c_start)
            c.value = val
            c.font = font_normal
            c.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(val, str) and val.startswith("="):
                c.number_format = '0.00' if label in ["σ", "CV %"] else '0.0'

        for c_i in range(1, nb_cols + 1):
            ws.cell(row=row_idx, column=c_i).border = thin_border

        ws.row_dimensions[row_idx].height = 22
        row_idx += 1

    # Largeurs de colonnes
    specific_widths = {
        'A': 8, 'B': 11, 'C': 10, 'D': 28, 'E': 10,
        'F': 10, 'G': 14, 'H': 14, 'I': 14
    }
    for col_letter, width in specific_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# 2. CHARGEMENT & TRAITEMENT SUPABASE AVEC RECHERCHE CROISÉE
# =========================================================

def load_and_process_controle_data(supabase):
    res_ecrasement = supabase.table("suivi_controle_beton").select("*").order("id", desc=True).execute()
    df_raw = pd.DataFrame(res_ecrasement.data) if res_ecrasement and res_ecrasement.data else pd.DataFrame()

    if df_raw.empty:
        return pd.DataFrame()

    res_betonnage = supabase.table("suivi_betonnage").select("*").execute()
    df_betonnage = pd.DataFrame(res_betonnage.data) if res_betonnage and res_betonnage.data else pd.DataFrame()

    col_mapping_input = {
        'affaissement': 'affaissement_mm',
        'affaissement_mm': 'affaissement_mm',
        'temperature': 'temp_beton_C',
        'temp_beton': 'temp_beton_C',
        'temp_beton_C': 'temp_beton_C'
    }
    df_raw = df_raw.rename(columns={k: v for k, v in col_mapping_input.items() if k in df_raw.columns})

    if "affaissement_mm" not in df_raw.columns:
        df_raw["affaissement_mm"] = None
    if "temp_beton_C" not in df_raw.columns:
        df_raw["temp_beton_C"] = None

    if not df_betonnage.empty:
        df_betonnage = df_betonnage.rename(columns={
            'affaissement': 'affaissement_mm_b',
            'temperature': 'temp_beton_C_b',
            'prelevement': 'ref_controle_b',
            'date_livraison': 'date_coulee_b',
            'ouvrage': 'ouvrage_b'
        })

    def clean_echeance(val):
        val_str = str(val).lower().strip()
        if "3" in val_str:
            return "3 jours"
        elif "7" in val_str:
            return "7 jours"
        elif "28" in val_str:
            return "28 jours"
        return val_str

    if "echeance" in df_raw.columns:
        df_raw["echeance_clean"] = df_raw["echeance"].apply(clean_echeance)
    else:
        df_raw["echeance_clean"] = ""

    if "fc_mpa" in df_raw.columns:
        df_raw["fc_mpa"] = pd.to_numeric(df_raw["fc_mpa"], errors="coerce")

    base_group_cols = ["ref_controle", "date_coulee", "classe_beton", "ouvrage"]
    existing_group_cols = [c for c in base_group_cols if c in df_raw.columns]

    if not existing_group_cols:
        return pd.DataFrame()

    if "date_coulee" in df_raw.columns:
        df_raw["date_dt"] = pd.to_datetime(df_raw["date_coulee"], errors="coerce")
    elif "date_ecrasement" in df_raw.columns:
        df_raw["date_dt"] = pd.to_datetime(df_raw["date_ecrasement"], errors="coerce")
    else:
        df_raw["date_dt"] = pd.NaT

    pivot_rows = []
    grouped = df_raw.groupby(existing_group_cols, dropna=False)

    for group_key, group_df in grouped:
        first_row = group_df.iloc[0]
        row_dict = {col: first_row[col] for col in existing_group_cols}
        
        aff_idx = group_df["affaissement_mm"].dropna().first_valid_index()
        temp_idx = group_df["temp_beton_C"].dropna().first_valid_index()

        aff_val = group_df.loc[aff_idx, "affaissement_mm"] if aff_idx is not None else None
        temp_val = group_df.loc[temp_idx, "temp_beton_C"] if temp_idx is not None else None

        if (pd.isna(aff_val) or pd.isna(temp_val)) and not df_betonnage.empty:
            ref = str(row_dict.get("ref_controle", "")).strip()
            dt_str = str(row_dict.get("date_coulee", "")).strip()
            ovr = str(row_dict.get("ouvrage", "")).strip()

            matched_b = pd.DataFrame()
            if ref and "ref_controle_b" in df_betonnage.columns:
                matched_b = df_betonnage[df_betonnage["ref_controle_b"].astype(str).str.strip() == ref]

            if matched_b.empty and dt_str and ovr and "date_coulee_b" in df_betonnage.columns and "ouvrage_b" in df_betonnage.columns:
                matched_b = df_betonnage[
                    (df_betonnage["date_coulee_b"].astype(str).str.strip() == dt_str) &
                    (df_betonnage["ouvrage_b"].astype(str).str.strip() == ovr)
                ]

            if not matched_b.empty:
                b_row = matched_b.iloc[0]
                if pd.isna(aff_val) and "affaissement_mm_b" in b_row:
                    aff_val = b_row["affaissement_mm_b"]
                if pd.isna(temp_val) and "temp_beton_C_b" in b_row:
                    temp_val = b_row["temp_beton_C_b"]

        row_dict["affaissement_mm"] = extract_numeric(aff_val)
        row_dict["temp_beton_C"] = extract_numeric(temp_val)

        valid_dates = group_df["date_dt"].dropna()
        row_dict["date_dt"] = valid_dates.min() if not valid_dates.empty else pd.NaT

        for ech in ["3 jours", "7 jours", "28 jours"]:
            vals = group_df[group_df["echeance_clean"] == ech]["fc_mpa"].dropna()
            if not vals.empty:
                row_dict[f"fc_mpa_{ech}"] = round(vals.mean(), 1)
            else:
                row_dict[f"fc_mpa_{ech}"] = None

        pivot_rows.append(row_dict)

    df_pivoted = pd.DataFrame(pivot_rows)

    desired_order = [
        "ref_controle", 
        "date_coulee", 
        "classe_beton", 
        "ouvrage",
        "affaissement_mm",
        "temp_beton_C",
        "fc_mpa_3 jours", 
        "fc_mpa_7 jours", 
        "fc_mpa_28 jours",
        "date_dt"
    ]
    
    existing_cols = [c for c in desired_order if c in df_pivoted.columns]
    df_pivoted = df_pivoted[existing_cols]

    rename_map = {
        "ref_controle": "Réf. Contrôle",
        "date_coulee": "Date Coulée",
        "classe_beton": "Classe Béton",
        "ouvrage": "Ouvrage",
        "affaissement_mm": "Affaissement (mm)",
        "temp_beton_C": "Temp. Béton (°C)",
        "fc_mpa_3 jours": "Moy. Fc (MPa) [3 Jours]",
        "fc_mpa_7 jours": "Moy. Fc (MPa) [7 Jours]",
        "fc_mpa_28 jours": "Moy. Fc (MPa) [28 Jours]"
    }

    df_pivoted = df_pivoted.rename(columns=rename_map)
    return df_pivoted


def format_controle_dataframe(df_filtered):
    df_display = df_filtered.copy()
    if "date_dt" in df_display.columns:
        df_display = df_display.drop(columns=["date_dt"])
    return df_display


# =========================================================
# CALCULS STATISTIQUES PANDAS
# =========================================================

def compute_statistics_df(df_display):
    cols_target = [
        "Affaissement (mm)", 
        "Temp. Béton (°C)", 
        "Moy. Fc (MPa) [7 Jours]", 
        "Moy. Fc (MPa) [28 Jours]"
    ]
    
    stats_data = {
        "Indicateur": ["MIN", "MOY", "MAX", "σ", "CV %"]
    }
    
    for col in cols_target:
        if col in df_display.columns:
            s = pd.to_numeric(df_display[col], errors='coerce').dropna()
            if not s.empty:
                v_min = round(s.min(), 1)
                v_moy = round(s.mean(), 1)
                v_max = round(s.max(), 1)
                
                if col == "Moy. Fc (MPa) [28 Jours]":
                    v_std = round(s.std(ddof=1), 2) if len(s) > 1 else 0.0
                    v_cv = round((v_std / v_moy) * 100, 1) if v_moy > 0 else 0.0
                    stats_data[col] = [v_min, v_moy, v_max, v_std, f"{v_cv} %"]
                else:
                    stats_data[col] = [v_min, v_moy, v_max, "-", "-"]
            else:
                stats_data[col] = ["-", "-", "-", "-", "-"]
        else:
            stats_data[col] = ["-", "-", "-", "-", "-"]

    return pd.DataFrame(stats_data)


# =========================================================
# 3. STREAMLIT APP VUES
# =========================================================

def show(supabase):
    st.title("📊 Module de Synthèses du Béton")

    main_tab_betonnage, main_tab_controle = st.tabs([
        "🏗️ Synthèse de Suivi de Bétonnage", 
        "🧪 Synthèse de Contrôle Béton"
    ])

    with main_tab_betonnage:
        st.subheader("Bilan du Suivi de Bétonnage")
        tab_j_b, tab_m_b = st.tabs(["📅 Bilan Journalier", "📅 Bilan Mensuel"])

        with tab_j_b:
            st.markdown("### Filtrage par jour et par classe de béton")
            col1, col2 = st.columns(2)
            with col1:
                selected_date = st.date_input("Sélectionnez une date :", value=date.today(), key="b_date_j")
            with col2:
                selected_class = st.selectbox(
                    "Filtrer par classe de béton :", 
                    ["Toutes", "C25/30", "C30/37", "C35/45", "C40/50", "C45/55"],
                    key="b_class_j"
                )

            try:
                res = supabase.table("suivi_betonnage").select("*").eq("date_livraison", str(selected_date)).execute()
                data = res.data if res else []

                if data:
                    df = pd.DataFrame(data)
                    if selected_class != "Toutes":
                        df = df[df["classe_beton"] == selected_class]

                    if df.empty:
                        st.info("Aucun coulage enregistré pour les critères sélectionnés.")
                    else:
                        if "heure_fin_coulage" in df.columns and "heure_arrivee" in df.columns:
                            def calc_duree(row):
                                try:
                                    h_fin = datetime.strptime(str(row["heure_fin_coulage"]), "%H:%M")
                                    h_arr = datetime.strptime(str(row["heure_arrivee"]), "%H:%M")
                                    return f"{int((h_arr - h_fin).total_seconds() / 60)} min"
                                except:
                                    return "-"
                            df["Durée de transport"] = df.apply(calc_duree, axis=1)

                        cols_drop = [c for c in ["id", "created_at", "created", "heure_fin_coulage", "client", "centrale_beton", "technicien", "observations", "nb_eprouvettes"] if c in df.columns]
                        df = df.drop(columns=cols_drop)

                        cols = list(df.columns)
                        if "date_livraison" in cols and "heure_arrivee" in cols:
                            cols.remove("heure_arrivee")
                            cols.insert(cols.index("date_livraison") + 1, "heure_arrivee")
                        if "meteo" in cols:
                            cols.remove("meteo")
                            cols.append("meteo")
                        df = df[cols]

                        renames = {
                            "date_livraison": "Date Livraison", "heure_arrivee": "Heure d'arrivée",
                            "bl_num": "N° BL", "ouvrage": "Ouvrage", "quantite_m3": "Quantité (m³)",
                            "classe_beton": "Classe", "temperature": "Temp. Béton",
                            "temperature_ambiante": "Temp. Ambiante", "affaissement": "Affaissement",
                            "prelevement": "Prélèvement", "meteo": "Météo"
                        }
                        df_display = df.rename(columns=renames)

                        st.markdown("---")
                        k1, k2 = st.columns(2)
                        k1.metric("Volume Total", f"{df_display['Quantité (m³)'].sum():.1f} m³")
                        if "Affaissement" in df_display.columns:
                            k2.metric("Affaissement Moyen", f"{pd.to_numeric(df_display['Affaissement'], errors='coerce').mean():.0f} mm")
                        st.markdown("---")

                        excel_file = generate_excel_synthesis_betonnage(df_display, f"Journée du {selected_date.strftime('%d/%m/%Y')}")
                        st.download_button(
                            label="📥 Télécharger la Synthèse Excel Bétonnage (A4 Portrait)",
                            data=excel_file,
                            file_name=f"Synthese_Betonnage_{selected_date}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.dataframe(df_display, use_container_width=True)
                else:
                    st.info("Aucun coulage enregistré pour les critères sélectionnés.")
            except Exception as e:
                st.error(f"Erreur de chargement : {e}")

        with tab_m_b:
            st.markdown("### Bilan mensuel global")
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                annee = date.today().year
                mois_liste = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
                mois_selected = st.selectbox("Sélectionnez le mois :", mois_liste, index=date.today().month - 1, key="b_mois_m")
                mois_num = mois_liste.index(mois_selected) + 1
            with col_m2:
                selected_class_m = st.selectbox(
                    "Filtrer par classe de béton (Mensuel) :", 
                    ["Toutes", "C25/30", "C30/37", "C35/45", "C40/50", "C45/55"],
                    key="b_class_m"
                )

            try:
                date_debut = f"{annee}-{mois_num:02d}-01"
                dernier_jour = 31 if mois_num in [1,3,5,7,8,10,12] else (30 if mois_num in [4,6,9,11] else 28)
                date_fin = f"{annee}-{mois_num:02d}-{dernier_jour}"

                res_m = supabase.table("suivi_betonnage").select("*").gte("date_livraison", date_debut).lte("date_livraison", date_fin).execute()
                data_m = res_m.data if res_m else []

                if data_m:
                    df_m = pd.DataFrame(data_m)
                    if selected_class_m != "Toutes":
                        df_m = df_m[df_m["classe_beton"] == selected_class_m]

                    if df_m.empty:
                        st.info("Aucun coulage enregistré pour ce mois.")
                    else:
                        if "heure_fin_coulage" in df_m.columns and "heure_arrivee" in df_m.columns:
                            def calc_duree(row):
                                try:
                                    h_fin = datetime.strptime(str(row["heure_fin_coulage"]), "%H:%M")
                                    h_arr = datetime.strptime(str(row["heure_arrivee"]), "%H:%M")
                                    return f"{int((h_arr - h_fin).total_seconds() / 60)} min"
                                except:
                                    return "-"
                            df_m["Durée de transport"] = df_m.apply(calc_duree, axis=1)

                        cols_drop = [c for c in ["id", "created_at", "created", "heure_fin_coulage", "client", "centrale_beton", "technicien", "observations", "nb_eprouvettes"] if c in df_m.columns]
                        df_m = df_m.drop(columns=cols_drop)

                        cols_m = list(df_m.columns)
                        if "date_livraison" in cols_m and "heure_arrivee" in cols_m:
                            cols_m.remove("heure_arrivee")
                            cols_m.insert(cols_m.index("date_livraison") + 1, "heure_arrivee")
                        if "meteo" in cols_m:
                            cols_m.remove("meteo")
                            cols_m.append("meteo")
                        df_m = df_m[cols_m]

                        renames = {
                            "date_livraison": "Date Livraison", "heure_arrivee": "Heure d'arrivée",
                            "bl_num": "N° BL", "ouvrage": "Ouvrage", "quantite_m3": "Quantité (m³)",
                            "classe_beton": "Classe", "temperature": "Temp. Béton",
                            "temperature_ambiante": "Temp. Ambiante", "affaissement": "Affaissement",
                            "prelevement": "Prélèvement", "meteo": "Météo"
                        }
                        df_m_display = df_m.rename(columns=renames)

                        st.markdown("---")
                        st.metric("Volume Cumulé du Mois", f"{df_m_display['Quantité (m³)'].sum():.1f} m³")
                        st.markdown("---")

                        excel_file_m = generate_excel_synthesis_betonnage(df_m_display, f"Mois de {mois_selected} {annee}")
                        st.download_button(
                            label="📥 Télécharger la Synthèse Mensuelle Excel Bétonnage (A4 Portrait)",
                            data=excel_file_m,
                            file_name=f"Synthese_Mensuelle_Betonnage_{mois_selected}_{annee}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.dataframe(df_m_display, use_container_width=True)
                else:
                    st.info("Aucun coulage enregistré pour ce mois.")
            except Exception as e:
                st.error(f"Erreur de chargement : {e}")

    with main_tab_controle:
        st.subheader("Bilan du Contrôle Béton (Moyennes par Prélèvement)")
        tab_j_c, tab_m_c = st.tabs(["📅 Bilan Journalier", "📅 Bilan Mensuel"])

        try:
            df_merged = load_and_process_controle_data(supabase)
        except Exception as e:
            st.error(f"Erreur lors de la préparation des données : {e}")
            df_merged = pd.DataFrame()

        classes_dispo = ["Toutes"]
        if not df_merged.empty and "Classe Béton" in df_merged.columns:
            classes_dispo += sorted(list(df_merged["Classe Béton"].dropna().unique()))

        with tab_j_c:
            st.markdown("### Filtrage journalier par classe de béton")
            col1, col2 = st.columns(2)
            with col1:
                selected_date_c = st.date_input("Sélectionnez une date :", value=date.today(), key="c_date_j")
            with col2:
                selected_class_cj = st.selectbox("Filtrer par classe de béton :", classes_dispo, key="c_class_j")

            if df_merged.empty:
                st.info("Aucune donnée disponible.")
            else:
                df_j_c = df_merged[df_merged["date_dt"].dt.date == selected_date_c]
                if selected_class_cj != "Toutes" and "Classe Béton" in df_j_c.columns:
                    df_j_c = df_j_c[df_j_c["Classe Béton"] == selected_class_cj]

                if df_j_c.empty:
                    st.info("Aucun contrôle enregistré pour les critères sélectionnés.")
                else:
                    df_display_cj = format_controle_dataframe(df_j_c)

                    st.markdown("---")
                    st.metric("Nombre de Prélèvements", f"{len(df_display_cj)}")
                    st.markdown("---")

                    excel_file_cj = generate_excel_synthesis_controle(df_display_cj, f"Journée du {selected_date_c.strftime('%d/%m/%Y')}")
                    st.download_button(
                        label="📥 Télécharger la Synthèse Contrôle Excel (A4 Portrait)",
                        data=excel_file_cj,
                        file_name=f"Synthese_Controle_Beton_{selected_date_c}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.dataframe(df_display_cj, use_container_width=True)

                    st.markdown("### 📈 Synthèse Statistique")
                    df_stats_j = compute_statistics_df(df_display_cj)
                    st.dataframe(df_stats_j, use_container_width=True)

        with tab_m_c:
            st.markdown("### Bilan mensuel par classe de béton")
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                annee_c = date.today().year
                mois_liste = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
                mois_selected_c = st.selectbox("Sélectionnez le mois :", mois_liste, index=date.today().month - 1, key="c_mois_m")
                mois_num_c = mois_liste.index(mois_selected_c) + 1
            with col_m2:
                selected_class_cm = st.selectbox("Filtrer par classe de béton :", classes_dispo, key="c_class_m")

            if df_merged.empty:
                st.info("Aucune donnée disponible.")
            else:
                df_m_c = df_merged[(df_merged["date_dt"].dt.year == annee_c) & (df_merged["date_dt"].dt.month == mois_num_c)]
                if selected_class_cm != "Toutes" and "Classe Béton" in df_m_c.columns:
                    df_m_c = df_m_c[df_m_c["Classe Béton"] == selected_class_cm]

                if df_m_c.empty:
                    st.info("Aucun contrôle enregistré pour ce mois.")
                else:
                    df_display_cm = format_controle_dataframe(df_m_c)

                    st.markdown("---")
                    st.metric("Total Prélèvements du Mois", f"{len(df_display_cm)}")
                    st.markdown("---")

                    excel_file_cm = generate_excel_synthesis_controle(df_display_cm, f"Mois de {mois_selected_c} {annee_c}")
                    st.download_button(
                        label="📥 Télécharger la Synthèse Mensuelle Contrôle Excel (A4 Portrait)",
                        data=excel_file_cm,
                        file_name=f"Synthese_Mensuelle_Controle_{mois_selected_c}_{annee_c}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.dataframe(df_display_cm, use_container_width=True)

                    st.markdown("### 📈 Synthèse Statistique Mensuelle")
                    df_stats_m = compute_statistics_df(df_display_cm)
                    st.dataframe(df_stats_m, use_container_width=True)
