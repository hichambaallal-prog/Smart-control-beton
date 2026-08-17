import streamlit as st
import pandas as pd
from datetime import datetime, date
import io

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# 1. GENERATION EXCEL
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
            cell.value = val
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

    thin_border_side = Side(style='thin', color='B0C4DE')
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    nb_cols = max(len(df_data.columns), 7)
    last_col_letter = get_column_letter(nb_cols)
    mid_col_idx = nb_cols // 2
    mid_col_letter = get_column_letter(mid_col_idx)
    next_mid_letter = get_column_letter(mid_col_idx + 1)

    ws.merge_cells(f"A1:{last_col_letter}2")
    ws["A1"].value = "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - CTR-CSB\nRAPPORT DE SYNTHÈSE DU CONTRÔLE BÉTON"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

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

    row_idx = 7
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📊 RÉSUMÉ GLOBAL"
    ws[f"A{row_idx}"].font = font_section

    row_idx += 1
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"].value = "Nombre Total de Contrôles"
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

    row_idx += 3
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📋 DÉTAIL DES ESSAIS ET ÉCRASEMENTS"
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

    for row_data in df_data.itertuples(index=False):
        for col_num, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_num)
            cell.value = val
            cell.font = font_normal
            cell.border = thin_border
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row_idx].height = 28
        row_idx += 1

    row_idx += 2
    ws.merge_cells(f"A{row_idx}:{mid_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "Responsables d'essai"
    ws[f"A{row_idx}"].font = font_bold
    ws[f"A{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"{next_mid_letter}{row_idx}:{last_col_letter}{row_idx}")
    ws[f"{next_mid_letter}{row_idx}"] = "Chef du Laboratoire"
    ws[f"{next_mid_letter}{row_idx}"].font = font_bold
    ws[f"{next_mid_letter}{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")

    row_idx += 1
    ws.merge_cells(f"A{row_idx}:{mid_col_letter}{row_idx+3}")
    ws[f"A{row_idx}"] = "Visa & Signature :"
    ws[f"A{row_idx}"].alignment = Alignment(horizontal="left", vertical="top")

    ws.merge_cells(f"{next_mid_letter}{row_idx}:{last_col_letter}{row_idx+3}")
    ws[f"{next_mid_letter}{row_idx}"] = "Visa & Signature :"
    ws[f"{next_mid_letter}{row_idx}"].alignment = Alignment(horizontal="left", vertical="top")

    for r in range(row_idx - 1, row_idx + 4):
        for c in range(1, mid_col_idx + 1):
            ws.cell(row=r, column=c).border = thin_border
        for c in range(mid_col_idx + 1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    col_widths = [20, 16, 16, 16, 18, 18, 18]
    for col_idx, width in enumerate(col_widths, 1):
        if col_idx <= nb_cols:
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# 2. CHARGEMENT & FUSION AVEC FALLBACK ET DEBOGAGE
# =========================================================

def load_and_process_controle_data(supabase):
    """Charge les données Supabase avec repli automatique si la clé de jointure echoue."""
    res_betonnage = supabase.table("suivi_betonnage").select("*").execute()
    res_ecrasement = supabase.table("suivi_controle_beton").select("*").execute()

    df_beton = pd.DataFrame(res_betonnage.data) if res_betonnage and res_betonnage.data else pd.DataFrame()
    df_ecrasement = pd.DataFrame(res_ecrasement.data) if res_ecrasement and res_ecrasement.data else pd.DataFrame()

    # Affichage du diagnostic dans un expander Streamlit
    with st.expander("🔍 Diagnostic Supabase (Vérification des colonnes et données)"):
        st.write("**Données `suivi_betonnage` :**", df_beton.head(3) if not df_beton.empty else "Vide")
        st.write("**Données `suivi_controle_beton` :**", df_ecrasement.head(3) if not df_ecrasement.empty else "Vide")

    if df_beton.empty:
        return pd.DataFrame()

    # Normalisation de la date bétonnage
    date_col_beton = None
    for col in ["date_livraison", "date_prelevement", "date", "created_at"]:
        if col in df_beton.columns:
            date_col_beton = col
            break

    if date_col_beton:
        df_beton["date_dt"] = pd.to_datetime(df_beton[date_col_beton], errors="coerce")
        df_beton["date_str"] = df_beton["date_dt"].dt.strftime("%Y-%m-%d")
    else:
        df_beton["date_dt"] = pd.NaT
        df_beton["date_str"] = ""

    if df_ecrasement.empty:
        df_beton["res_3j"] = None
        df_beton["res_7j"] = None
        df_beton["res_28j"] = None
        df_beton["_ref_col"] = "-"
        return df_beton

    # Normalisation des colonnes de suivi_controle_beton
    # Recherche de la colonne de référence
    ref_col = None
    for c in ["ref_controle", "prefixe_repere", "repere", "reference", "code_controle", "id"]:
        if c in df_ecrasement.columns:
            ref_col = c
            break

    # Normalisation de la date d'écrasement/prélèvement
    date_col_ecr = None
    for c in ["date_prelevement", "date_livraison", "date_essai", "date"]:
        if c in df_ecrasement.columns:
            date_col_ecr = c
            break

    if date_col_ecr:
        df_ecrasement["date_ecr_str"] = pd.to_datetime(df_ecrasement[date_col_ecr], errors="coerce").dt.strftime("%Y-%m-%d")
    else:
        df_ecrasement["date_ecr_str"] = ""

    # Nettoyage échéance et résistance
    df_ecrasement["resistance_num"] = pd.to_numeric(df_ecrasement.get("resistance"), errors="coerce")
    df_ecrasement["echeance_clean"] = (
        df_ecrasement.get("echeance", "")
        .astype(str)
        .str.lower()
        .str.extract(r'(\d+)')[0]
    )

    # Recherche clé de jointure directe
    join_key_b = None
    for k in ["bl_num", "bl", "id_betonnage", "id"]:
        if k in df_beton.columns:
            join_key_b = k
            break

    join_key_e = None
    for k in ["bl_num", "bl", "id_betonnage", "id_beton"]:
        if k in df_ecrasement.columns:
            join_key_e = k
            break

    # TENTATIVE 1: Jointure directe par clé (ex: N° BL)
    matched = False
    if join_key_b and join_key_e:
        df_beton["key_clean"] = df_beton[join_key_b].astype(str).str.strip().str.upper()
        df_ecrasement["key_clean"] = df_ecrasement[join_key_e].astype(str).str.strip().str.upper()

        common_keys = set(df_beton["key_clean"]).intersection(set(df_ecrasement["key_clean"]))
        if common_keys and common_keys != {""}, common_keys != {"NONE"}, common_keys != {"NAN"}:
            matched = True

    # TENTATIVE 2 (FALLBACK) : Jointure par Date si la clé directe échoue
    if not matched:
        df_beton["key_clean"] = df_beton["date_str"]
        df_ecrasement["key_clean"] = df_ecrasement["date_ecr_str"]

    # Agrégation des résistances par key_clean
    res_3j = df_ecrasement[df_ecrasement["echeance_clean"] == "3"].groupby("key_clean")["resistance_num"].mean().rename("res_3j")
    res_7j = df_ecrasement[df_ecrasement["echeance_clean"] == "7"].groupby("key_clean")["resistance_num"].mean().rename("res_7j")
    res_28j = df_ecrasement[df_ecrasement["echeance_clean"] == "28"].groupby("key_clean")["resistance_num"].mean().rename("res_28j")

    if ref_col:
        refs = (
            df_ecrasement[df_ecrasement[ref_col].notnull() & (df_ecrasement[ref_col].astype(str).str.strip() != "")]
            .groupby("key_clean")[ref_col]
            .first()
            .rename("ref_controle_found")
        )
    else:
        refs = pd.Series(dtype=str, name="ref_controle_found")

    df_merged = df_beton.merge(refs, left_on="key_clean", right_index=True, how="left")
    df_merged = df_merged.merge(res_3j, left_on="key_clean", right_index=True, how="left")
    df_merged = df_merged.merge(res_7j, left_on="key_clean", right_index=True, how="left")
    df_merged = df_merged.merge(res_28j, left_on="key_clean", right_index=True, how="left")

    df_merged["_ref_col"] = df_merged["ref_controle_found"].fillna("-")

    return df_merged


def format_controle_dataframe(df_filtered):
    df_display = pd.DataFrame()
    
    df_display["Référence de Contrôle"] = df_filtered["_ref_col"]
    df_display["Date de Prélèvement"] = df_filtered["date_dt"].dt.strftime("%d/%m/%Y").fillna("-")
    df_display["Affaissement (cm)"] = df_filtered.get("affaissement", "-").fillna("-")
    df_display["Température (°C)"] = df_filtered.get("temperature", "-").fillna("-")
    
    df_display["Résistance moyenne 3J (MPa)"] = df_filtered["res_3j"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "-")
    df_display["Résistance moyenne 7J (MPa)"] = df_filtered["res_7j"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "-")
    df_display["Résistance moyenne 28J (MPa)"] = df_filtered["res_28j"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "-")
    
    return df_display


# =========================================================
# 3. STREAMLIT APP
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
                        k2.metric("Affaissement Moyen", f"{df_display['Affaissement'].mean():.0f} mm")
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
        st.subheader("Bilan du Contrôle Béton (Résistances aux écrasements)")
        tab_j_c, tab_m_c = st.tabs(["📅 Bilan Journalier", "📅 Bilan Mensuel"])

        try:
            df_merged = load_and_process_controle_data(supabase)
        except Exception as e:
            st.error(f"Erreur lors de la préparation des données : {e}")
            df_merged = pd.DataFrame()

        classes_dispo = ["Toutes"]
        if not df_merged.empty and "classe_beton" in df_merged.columns:
            classes_dispo += sorted(list(df_merged["classe_beton"].dropna().unique()))

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
                if selected_class_cj != "Toutes" and "classe_beton" in df_j_c.columns:
                    df_j_c = df_j_c[df_j_c["classe_beton"] == selected_class_cj]

                if df_j_c.empty:
                    st.info("Aucun contrôle enregistré pour les critères sélectionnés.")
                else:
                    df_display_cj = format_controle_dataframe(df_j_c)

                    st.markdown("---")
                    st.metric("Nombre de Contrôles", f"{len(df_display_cj)}")
                    st.markdown("---")

                    excel_file_cj = generate_excel_synthesis_controle(df_display_cj, f"Journée du {selected_date_c.strftime('%d/%m/%Y')}")
                    st.download_button(
                        label="📥 Télécharger la Synthèse Contrôle Excel (A4 Portrait)",
                        data=excel_file_cj,
                        file_name=f"Synthese_Controle_Beton_{selected_date_c}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.dataframe(df_display_cj, use_container_width=True)

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
                if selected_class_cm != "Toutes" and "classe_beton" in df_m_c.columns:
                    df_m_c = df_m_c[df_m_c["classe_beton"] == selected_class_cm]

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
