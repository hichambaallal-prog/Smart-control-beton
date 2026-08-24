import io
import re
from datetime import date, datetime
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# =========================================================
# STYLES ET CONSTANTES EXCEL
# =========================================================
COLOR_PRIMARY, COLOR_HEADER, COLOR_SUB_HEADER = "1F4E79", "2D572C", "E2EFDA"
COLOR_CARD_BG, COLOR_KPI_BG, COLOR_STAT_BG = "F7F9FA", "EDF2F8", "F2F2F2"

FONT_TITLE = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
FONT_TITLE_CTRL = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
FONT_SECTION = Font(name="Calibri", size=13, bold=True, color=COLOR_PRIMARY)
FONT_SECTION_CTRL = Font(name="Calibri", size=12, bold=True, color=COLOR_PRIMARY)
FONT_BOLD = Font(name="Calibri", size=11, bold=True)
FONT_NORMAL = Font(name="Calibri", size=11)
FONT_TH = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_SUB_TH = Font(name="Calibri", size=10, bold=True, color="000000")
FONT_KPI_VAL = Font(name="Calibri", size=14, bold=True, color=COLOR_PRIMARY)
FONT_KPI_VAL_CTRL = Font(name="Calibri", size=13, bold=True, color=COLOR_PRIMARY)
FONT_NAME_BLUE = Font(name="Calibri", size=11, bold=True, color=COLOR_PRIMARY)

FILL_TITLE = PatternFill("solid", fgColor=COLOR_PRIMARY)
FILL_TH = PatternFill("solid", fgColor=COLOR_HEADER)
FILL_SUB_TH = PatternFill("solid", fgColor=COLOR_SUB_HEADER)
FILL_CARD = PatternFill("solid", fgColor=COLOR_CARD_BG)
FILL_KPI = PatternFill("solid", fgColor=COLOR_KPI_BG)
FILL_STAT = PatternFill("solid", fgColor=COLOR_STAT_BG)
FILL_STAT_HDR = PatternFill("solid", fgColor=COLOR_PRIMARY)

THIN_SIDE = Side(style="thin", color="B0C4DE")
BORDER_THIN = Border(
    left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE
)
DARK_SIDE = Side(style="thin", color="000000")
BORDER_DARK = Border(
    left=DARK_SIDE, right=DARK_SIDE, top=DARK_SIDE, bottom=DARK_SIDE
)
BORDER_TOTAL = Border(
    top=DARK_SIDE, bottom=Side(style="double", color="000000")
)

ALIGN_CENTER = Alignment(
    horizontal="center", vertical="center", wrap_text=True
)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")

MOIS_LISTE = [
    "Janvier",
    "Février",
    "Mars",
    "Avril",
    "Mai",
    "Juin",
    "Juillet",
    "Août",
    "Septembre",
    "Octobre",
    "Novembre",
    "Décembre",
]


# =========================================================
# HELPERS EXCEL & UTILITAIRES
# =========================================================
def extract_numeric(val):
  if pd.isna(val) or val is None:
    return None
  match = re.search(r"[-+]?\d*\.\d+|\d+", str(val).replace(",", "."))
  if match:
    num = float(match.group())
    return int(num) if num.is_integer() else num
  return None


def _init_excel_sheet(title_name, title_text, periode, font_title_obj, nb_cols):
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = title_name
  ws.page_setup.orientation, ws.page_setup.paperSize = (
      ws.ORIENTATION_PORTRAIT,
      ws.PAPERSIZE_A4,
  )
  ws.sheet_properties.pageSetUpPr.fitToPage = True
  ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
  for m, v in zip(["left", "right", "top", "bottom"], [0.3, 0.3, 0.4, 0.4]):
    setattr(ws.page_margins, m, v)

  last_col = get_column_letter(nb_cols)
  mid_idx = max(nb_cols // 2, 1)
  mid_col, next_mid = get_column_letter(mid_idx), get_column_letter(mid_idx + 1)

  ws.merge_cells(f"A1:{last_col}2")
  ws["A1"].value, ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = (
      title_text,
      font_title_obj,
      FILL_TITLE,
      ALIGN_CENTER,
  )
  ws.row_dimensions[1].height = ws.row_dimensions[2].height = 25

  cards = [
      (f"A4:{mid_col}4", "   CLIENT :   TGCC"),
      (f"{next_mid}4:{last_col}4", "   PROJET :   LGV CASA SUD"),
      (f"A5:{mid_col}5", f"   PÉRIODE :   {periode}"),
      (
          f"{next_mid}5:{last_col}5",
          f"   DATE ÉDITION :   {datetime.now().strftime('%d/%m/%Y')}",
      ),
  ]
  for rng, txt in cards:
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value, c.font, c.fill, c.alignment = txt, FONT_BOLD, FILL_CARD, ALIGN_LEFT

  for r in (4, 5):
    ws.row_dimensions[r].height = 28
    for c_i in range(1, nb_cols + 1):
      ws.cell(row=r, column=c_i).border = BORDER_THIN

  return wb, ws, last_col, mid_idx


def _add_signatures(ws, row_idx, mid_col_idx, nb_cols):
  last_col, mid_col, next_mid = (
      get_column_letter(nb_cols),
      get_column_letter(mid_col_idx),
      get_column_letter(mid_col_idx + 1),
  )
  sigs = [
      (f"A{row_idx}:{mid_col}{row_idx}", "Responsable d'essai :", FONT_BOLD),
      (
          f"{next_mid}{row_idx}:{last_col}{row_idx}",
          "Chef du laboratoire :",
          FONT_BOLD,
      ),
      (f"A{row_idx+1}:{mid_col}{row_idx+1}", "O.IKKEN", FONT_NAME_BLUE),
      (
          f"{next_mid}{row_idx+1}:{last_col}{row_idx+1}",
          "H.BAALLAL",
          FONT_NAME_BLUE,
      ),
  ]
  for rng, txt, font in sigs:
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value, c.font, c.alignment = txt, font, ALIGN_CENTER
  ws.row_dimensions[row_idx].height = ws.row_dimensions[row_idx + 1].height = 20


# =========================================================
# 1. GENERATION EXCEL
# =========================================================
def generate_excel_synthesis_betonnage(
    df_data, titre_periode, is_mensuel=False
):
  output = io.BytesIO()
  is_multi = isinstance(df_data.columns, pd.MultiIndex)
  nb_cols = min(max(len(df_data.columns), 10), 10)

  wb, ws, last_col_letter, mid_col_idx = _init_excel_sheet(
      "Synthèse Bétonnage",
      (
          "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) -"
          " CTR-CSB\nRAPPORT DE SYNTHÈSE DU BÉTONNAGE"
      ),
      titre_periode,
      FONT_TITLE,
      nb_cols,
  )

  # KPI Volume
  ws.merge_cells(f"A7:{last_col_letter}7")
  ws["A7"].value, ws["A7"].font = "📊 RÉSUMÉ GLOBAL", FONT_SECTION
  ws.row_dimensions[7].height = 28

  vol_tot = (
      df_data[[c for c in df_data.columns if c[0] == "Quantité (m³)"][0]].sum()
      if is_multi
      else df_data.get("Quantité (m³)", pd.Series([0])).sum()
  )

  ws.merge_cells(f"A8:{last_col_letter}8")
  ws["A8"].value, ws["A8"].font, ws["A8"].fill, ws["A8"].alignment = (
      "Volume Total Béton",
      FONT_BOLD,
      FILL_KPI,
      ALIGN_CENTER,
  )
  ws.merge_cells(f"A9:{last_col_letter}9")
  ws["A9"].value, ws["A9"].font, ws["A9"].fill, ws["A9"].alignment = (
      f"{vol_tot:.1f} m³",
      FONT_KPI_VAL,
      FILL_KPI,
      ALIGN_CENTER,
  )

  for r in (8, 9):
    for c in range(1, nb_cols + 1):
      ws.cell(row=r, column=c).border = BORDER_THIN
  ws.row_dimensions[8].height, ws.row_dimensions[9].height = 24, 32

  row_idx = 11
  ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
  ws[f"A{row_idx}"].value, ws[f"A{row_idx}"].font = (
      "📋 DÉTAIL DES CONTRÔLES",
      FONT_SECTION,
  )
  ws.row_dimensions[row_idx].height = 28
  row_idx += 1

  # Entêtes Tableau
  if is_multi:
    for col_i, (top_cat, sub_cat) in enumerate(df_data.columns[:nb_cols], 1):
      for r_off, val, font, fill in [
          (0, top_cat, FONT_TH, FILL_TH),
          (1, sub_cat, FONT_SUB_TH, FILL_SUB_TH),
      ]:
        c = ws.cell(row=row_idx + r_off, column=col_i, value=val)
        c.font, c.fill, c.alignment, c.border = (
            font,
            fill,
            ALIGN_CENTER,
            BORDER_THIN,
        )

    col_i, real_len = 1, min(len(df_data.columns), nb_cols)
    while col_i <= real_len:
      top_val = df_data.columns[col_i - 1][0]
      span = sum(
          1
          for c in df_data.columns[:nb_cols]
          if c[0] == top_val and top_val != ""
      )
      if span > 1:
        ws.merge_cells(
            start_row=row_idx,
            start_column=col_i,
            end_row=row_idx,
            end_column=col_i + span - 1,
        )
        col_i += span
      else:
        if df_data.columns[col_i - 1][1] == "":
          ws.merge_cells(
              start_row=row_idx,
              start_column=col_i,
              end_row=row_idx + 1,
              end_column=col_i,
          )
        col_i += 1
    ws.row_dimensions[row_idx].height, ws.row_dimensions[row_idx + 1].height = (
        25,
        22,
    )
    row_idx += 2
  else:
    for col_num, h_name in enumerate(df_data.columns[:nb_cols], 1):
      c = ws.cell(row=row_idx, column=col_num, value=str(h_name))
      c.font, c.fill, c.alignment, c.border = (
          FONT_TH,
          FILL_TH,
          ALIGN_CENTER,
          BORDER_THIN,
      )
    ws.row_dimensions[row_idx].height = 35
    row_idx += 1

  # Données Tableau
  start_data_row = row_idx
  for row_data in df_data.itertuples(index=False):
    for col_num, val in enumerate(row_data[:nb_cols], 1):
      c = ws.cell(
          row=row_idx, column=col_num, value=val if pd.notna(val) else ""
      )
      c.font, c.border, c.alignment = FONT_NORMAL, BORDER_THIN, ALIGN_CENTER
    ws.row_dimensions[row_idx].height = 26
    row_idx += 1
  end_data_row = row_idx - 1

  # Statistiques MIN / MAX
  headers_flat = [
      c[0] if is_multi else str(c) for c in df_data.columns[:nb_cols]
  ]
  target_kws = [
      "temp. béton",
      "temp. ambiante",
      "affaissement",
      "temperature",
      "aff",
  ]
  target_cols = [
      i + 1
      for i, h in enumerate(headers_flat)
      if any(kw in h.lower() for kw in target_kws)
  ]

  for stat_label in ["MIN", "MAX"]:
    ws.row_dimensions[row_idx].height = 26
    ws.cell(
        row=row_idx,
        column=1,
        value=stat_label,
        font=FONT_BOLD,
        fill=FILL_STAT,
        alignment=ALIGN_CENTER,
        border=BORDER_DARK,
    )
    for col_num in range(2, len(headers_flat) + 1):
      c = ws.cell(
          row=row_idx,
          column=col_num,
          font=FONT_BOLD,
          fill=FILL_STAT,
          alignment=ALIGN_CENTER,
          border=BORDER_DARK,
      )
      if col_num in target_cols:
        col_ltr = get_column_letter(col_num)
        c.value = (
            f"={stat_label}({col_ltr}{start_data_row}:{col_ltr}{end_data_row})"
            if start_data_row <= end_data_row
            else "-"
        )
        if start_data_row <= end_data_row:
          c.number_format = "0.0"
      else:
        c.value = ""
    row_idx += 1

  # Total Volume
  ws.row_dimensions[row_idx].height = 30
  for col_num in range(1, len(headers_flat) + 1):
    c = ws.cell(
        row=row_idx,
        column=col_num,
        font=FONT_BOLD,
        border=BORDER_TOTAL,
        alignment=ALIGN_CENTER,
    )
    if col_num == 1:
      c.value = "TOTAL"
    else:
      col_name = headers_flat[col_num - 1]
      if "quantité" in col_name.lower() or "volume" in col_name.lower():
        col_ltr = get_column_letter(col_num)
        c.value, c.number_format = (
            f"=SUM({col_ltr}{start_data_row}:{col_ltr}{end_data_row})",
            '0.0 "m³"',
        )

  _add_signatures(ws, row_idx + 3, mid_col_idx, nb_cols)

  widths = [14, 12, 11.22, 40.0, 11.22, 11.22, 11.22, 11.22, 12, 12]
  for idx, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(idx)].width = w

  wb.save(output)
  output.seek(0)
  return output.getvalue()


def generate_excel_synthesis_controle(df_data, titre_periode):
  output = io.BytesIO()
  nb_cols = min(max(len(df_data.columns), 9), 10)

  wb, ws, last_col_letter, mid_col_idx = _init_excel_sheet(
      "Synthèse Contrôle Béton",
      (
          "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) -"
          " CTR-CSB\nRAPPORT DE SYNTHÈSE DU CONTRÔLE BÉTON"
      ),
      titre_periode,
      FONT_TITLE_CTRL,
      nb_cols,
  )

  # KPI Nb Prélèvements
  ws.merge_cells(f"A7:{last_col_letter}7")
  ws["A7"].value, ws["A7"].font = "📊 RÉSUMÉ GLOBAL", FONT_SECTION_CTRL

  ws.merge_cells(f"A8:{last_col_letter}8")
  ws["A8"].value, ws["A8"].font, ws["A8"].fill, ws["A8"].alignment = (
      "Nombre Total de Prélèvements Contrôlés",
      FONT_BOLD,
      FILL_KPI,
      ALIGN_CENTER,
  )
  ws.merge_cells(f"A9:{last_col_letter}9")
  ws["A9"].value, ws["A9"].font, ws["A9"].fill, ws["A9"].alignment = (
      f"{len(df_data)} prélèvement(s)",
      FONT_KPI_VAL_CTRL,
      FILL_KPI,
      ALIGN_CENTER,
  )

  for r in (8, 9):
    for c in range(1, nb_cols + 1):
      ws.cell(row=r, column=c).border = BORDER_THIN

  row_idx = 11
  ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
  ws[f"A{row_idx}"].value, ws[f"A{row_idx}"].font = (
      "📋 MOYENNE DES ÉCRASEMENTS PAR ÉCHÉANCE",
      FONT_SECTION_CTRL,
  )
  row_idx += 1

  # Entêtes Tableau
  for col_num, h_name in enumerate(df_data.columns[:nb_cols], 1):
    c = ws.cell(row=row_idx, column=col_num, value=str(h_name))
    c.font, c.fill, c.alignment = FONT_TH, FILL_TH, ALIGN_CENTER
  ws.row_dimensions[row_idx].height = 35
  row_idx += 1

  # Données Tableau
  start_data_row = row_idx
  for row_data in df_data.itertuples(index=False):
    for col_num, val in enumerate(row_data[:nb_cols], 1):
      c = ws.cell(
          row=row_idx, column=col_num, value=val if pd.notna(val) else ""
      )
      c.font, c.border, c.alignment = FONT_NORMAL, BORDER_THIN, ALIGN_CENTER
    ws.row_dimensions[row_idx].height = 28
    row_idx += 1
  end_data_row = row_idx - 1

  # Synthèse Statistique
  row_idx += 1
  ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
  ws[f"A{row_idx}"].value, ws[f"A{row_idx}"].font = (
      "📈 SYNTHÈSE STATISTIQUE DES PARAMÈTRES",
      FONT_SECTION_CTRL,
  )
  row_idx += 1

  stat_headers = [
      "Indicateur",
      "Affaissement (mm)",
      "Temp. Béton (°C)",
      "Fc (MPa) [7 Jours]",
      "Fc (MPa) [28 Jours]",
  ]
  for idx, h in enumerate(stat_headers):
    col_start = 1 if idx == 0 else 2 + (idx - 1) * 2
    col_end = 1 if idx == 0 else col_start + 1
    if col_start <= nb_cols:
      if col_start < col_end and col_end <= nb_cols:
        ws.merge_cells(
            start_row=row_idx,
            start_column=col_start,
            end_row=row_idx,
            end_column=col_end,
        )
      c = ws.cell(row=row_idx, column=col_start, value=h)
      c.font, c.fill, c.alignment = (
          Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
          FILL_STAT_HDR,
          ALIGN_CENTER,
      )

  for c_i in range(1, nb_cols + 1):
    ws.cell(row=row_idx, column=c_i).border = BORDER_THIN
  ws.row_dimensions[row_idx].height = 25
  row_idx += 1

  row_moy, row_std = row_idx + 1, row_idx + 3
  col_map = {"aff": "E", "temp": "F", "fc7": "H", "fc28": "I"}

  if start_data_row <= end_data_row:
    f = lambda k, fn: (
        f"={fn}({col_map[k]}{start_data_row}:{col_map[k]}{end_data_row})"
    )
    f_min_aff, f_moy_aff, f_max_aff = (
        f("aff", "MIN"),
        f("aff", "AVERAGE"),
        f("aff", "MAX"),
    )
    f_min_temp, f_moy_temp, f_max_temp = (
        f("temp", "MIN"),
        f("temp", "AVERAGE"),
        f("temp", "MAX"),
    )
    f_min_fc7, f_moy_fc7, f_max_fc7 = (
        f("fc7", "MIN"),
        f("fc7", "AVERAGE"),
        f("fc7", "MAX"),
    )
    f_min_fc28, f_moy_fc28, f_max_fc28 = (
        f("fc28", "MIN"),
        f("fc28", "AVERAGE"),
        f("fc28", "MAX"),
    )
    f_std_fc28 = f("fc28", "ECARTYPE.STANDARD")
    f_cv_fc28 = (
        f"=SIERREUR(({col_map['fc28']}{row_std}/{col_map['fc28']}{row_moy})*100,"
        " 0)"
    )
  else:
    f_min_aff = f_moy_aff = f_max_aff = f_min_temp = f_moy_temp = f_max_temp = (
        f_min_fc7
    ) = f_moy_fc7 = f_max_fc7 = f_min_fc28 = f_moy_fc28 = f_max_fc28 = (
        f_std_fc28
    ) = f_cv_fc28 = "-"

  stat_rows = [
      ("MIN", f_min_aff, f_min_temp, f_min_fc7, f_min_fc28),
      ("MOY", f_moy_aff, f_moy_temp, f_moy_fc7, f_moy_fc28),
      ("MAX", f_max_aff, f_max_temp, f_max_fc7, f_max_fc28),
      ("σ", "-", "-", "-", f_std_fc28),
      ("CV %", "-", "-", "-", f_cv_fc28),
  ]

  for label, v_aff, v_temp, v_fc7, v_fc28 in stat_rows:
    c_lbl = ws.cell(row=row_idx, column=1, value=label)
    c_lbl.font, c_lbl.alignment = FONT_BOLD, ALIGN_CENTER

    for idx, val in enumerate([v_aff, v_temp, v_fc7, v_fc28]):
      c_start = 2 + idx * 2
      c_end = c_start + 1
      if c_start <= nb_cols:
        if c_end <= nb_cols:
          ws.merge_cells(
              start_row=row_idx,
              start_column=c_start,
              end_row=row_idx,
              end_column=c_end,
          )
        c = ws.cell(row=row_idx, column=c_start, value=val)
        c.font, c.alignment = FONT_NORMAL, ALIGN_CENTER
        if isinstance(val, str) and val.startswith("="):
          c.number_format = "0.00" if label in ["σ", "CV %"] else "0.0"

    for c_i in range(1, nb_cols + 1):
      ws.cell(row=row_idx, column=c_i).border = BORDER_THIN
    ws.row_dimensions[row_idx].height = 22
    row_idx += 1

  _add_signatures(ws, row_idx + 2, mid_col_idx, nb_cols)

  ws.column_dimensions["A"].width, ws.column_dimensions["C"].width = 14, 40.0
  for col_l in ["B", "D", "E", "F", "G", "H", "I", "J"]:
    ws.column_dimensions[col_l].width = 12

  wb.save(output)
  output.seek(0)
  return output.getvalue()


# =========================================================
# 2. CHARGEMENT & TRAITEMENT SUPABASE
# =========================================================
def load_and_process_controle_data(supabase):
  res_ecrasement = (
      supabase.table("suivi_controle_beton")
      .select("*")
      .order("id", desc=True)
      .execute()
  )
  df_raw = (
      pd.DataFrame(res_ecrasement.data)
      if res_ecrasement and res_ecrasement.data
      else pd.DataFrame()
  )
  if df_raw.empty:
    return pd.DataFrame()

  res_betonnage = supabase.table("suivi_betonnage").select("*").execute()
  df_betonnage = (
      pd.DataFrame(res_betonnage.data)
      if res_betonnage and res_betonnage.data
      else pd.DataFrame()
  )

  col_mapping = {
      "affaissement": "affaissement_mm",
      "temperature": "temp_beton_C",
      "temp_beton": "temp_beton_C",
  }
  df_raw = df_raw.rename(
      columns={k: v for k, v in col_mapping.items() if k in df_raw.columns}
  )

  for col in ["affaissement_mm", "temp_beton_C"]:
    if col not in df_raw.columns:
      df_raw[col] = None

  if not df_betonnage.empty:
    df_betonnage = df_betonnage.rename(
        columns={
            "affaissement": "affaissement_mm_b",
            "temperature": "temp_beton_C_b",
            "prelevement": "ref_controle_b",
            "date_livraison": "date_coulee_b",
            "ouvrage": "ouvrage_b",
        }
    )

  def clean_echeance(val):
    v = str(val).lower().strip()
    return (
        "3 jours"
        if "3" in v
        else "7 jours" if "7" in v else "28 jours" if "28" in v else v
    )

  df_raw["echeance_clean"] = (
      df_raw["echeance"].apply(clean_echeance)
      if "echeance" in df_raw.columns
      else ""
  )
  if "fc_mpa" in df_raw.columns:
    df_raw["fc_mpa"] = pd.to_numeric(df_raw["fc_mpa"], errors="coerce")

  base_group_cols = ["ref_controle", "date_coulee", "classe_beton", "ouvrage"]
  existing_group_cols = [c for c in base_group_cols if c in df_raw.columns]
  if not existing_group_cols:
    return pd.DataFrame()

  date_col = (
      "date_coulee"
      if "date_coulee" in df_raw.columns
      else "date_ecrasement" if "date_ecrasement" in df_raw.columns else None
  )
  df_raw["date_dt"] = (
      pd.to_datetime(df_raw[date_col], errors="coerce")
      if date_col
      else pd.NaT
  )

  pivot_rows = []
  for group_key, group_df in df_raw.groupby(existing_group_cols, dropna=False):
    row_dict = group_df.iloc[0][existing_group_cols].to_dict()

    aff_idx, temp_idx = (
        group_df["affaissement_mm"].dropna().first_valid_index(),
        group_df["temp_beton_C"].dropna().first_valid_index(),
    )
    aff_val = (
        group_df.loc[aff_idx, "affaissement_mm"] if aff_idx is not None else None
    )
    temp_val = (
        group_df.loc[temp_idx, "temp_beton_C"] if temp_idx is not None else None
    )

    if (pd.isna(aff_val) or pd.isna(temp_val)) and not df_betonnage.empty:
      ref, dt_str, ovr = (
          str(row_dict.get("ref_controle", "")).strip(),
          str(row_dict.get("date_coulee", "")).strip(),
          str(row_dict.get("ouvrage", "")).strip(),
      )
      matched_b = pd.DataFrame()
      if ref and "ref_controle_b" in df_betonnage.columns:
        matched_b = df_betonnage[
            df_betonnage["ref_controle_b"].astype(str).str.strip() == ref
        ]
      if (
          matched_b.empty
          and dt_str
          and ovr
          and "date_coulee_b" in df_betonnage.columns
          and "ouvrage_b" in df_betonnage.columns
      ):
        matched_b = df_betonnage[
            (df_betonnage["date_coulee_b"].astype(str).str.strip() == dt_str)
            & (df_betonnage["ouvrage_b"].astype(str).str.strip() == ovr)
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
      row_dict[f"fc_mpa_{ech}"] = (
          round(vals.mean(), 1) if not vals.empty else None
      )

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
      "date_dt",
  ]
  df_pivoted = df_pivoted[[c for c in desired_order if c in df_pivoted.columns]]

  return df_pivoted.rename(
      columns={
          "ref_controle": "Réf. Contrôle",
          "date_coulee": "Date Coulée",
          "classe_beton": "Classe Béton",
          "ouvrage": "Ouvrage",
          "affaissement_mm": "Affaissement (mm)",
          "temp_beton_C": "Temp. Béton (°C)",
          "fc_mpa_3 jours": "Moy. Fc (MPa) [3 Jours]",
          "fc_mpa_7 jours": "Moy. Fc (MPa) [7 Jours]",
          "fc_mpa_28 jours": "Moy. Fc (MPa) [28 Jours]",
      }
  )


def format_controle_dataframe(df_filtered):
  return (
      df_filtered.drop(columns=["date_dt"])
      if "date_dt" in df_filtered.columns
      else df_filtered.copy()
  )


def compute_statistics_df(df_display):
  cols_target = [
      "Affaissement (mm)",
      "Temp. Béton (°C)",
      "Moy. Fc (MPa) [7 Jours]",
      "Moy. Fc (MPa) [28 Jours]",
  ]
  stats_data = {"Indicateur": ["MIN", "MOY", "MAX", "σ", "CV %"]}

  for col in cols_target:
    if col in df_display.columns:
      s = pd.to_numeric(df_display[col], errors="coerce").dropna()
      if not s.empty:
        v_min, v_moy, v_max = (
            round(s.min(), 1),
            round(s.mean(), 1),
            round(s.max(), 1),
        )
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
  main_tab_betonnage, main_tab_controle = st.tabs(
      ["🏗️ Synthèse de Suivi de Bétonnage", "🧪 Synthèse de Contrôle Béton"]
  )

  # --- TAB 1: BETONNAGE ---
  with main_tab_betonnage:
    st.subheader("Bilan du Suivi de Bétonnage")
    tab_j_b, tab_m_b = st.tabs(["📅 Bilan Journalier", "📅 Bilan Mensuel"])

    with tab_j_b:
      st.markdown("### Filtrage par jour et par classe de béton")
      col1, col2 = st.columns(2)
      with col1:
        selected_date = st.date_input(
            "Sélectionnez une date :", value=date.today(), key="b_date_j"
        )

      try:
        res = (
            supabase.table("suivi_betonnage")
            .select("*")
            .eq("date_livraison", str(selected_date))
            .execute()
        )
        data = res.data if res else []
        classes_j = ["Toutes"]
        if data:
          df_temp = pd.DataFrame(data)
          if "classe_beton" in df_temp.columns:
            classes_j += sorted(
                list(df_temp["classe_beton"].dropna().unique())
            )

        with col2:
          selected_class = st.selectbox(
              "Filtrer par classe de béton :", classes_j, key="b_class_j"
          )

        if data:
          df = pd.DataFrame(data)
          if selected_class != "Toutes":
            df = df[df["classe_beton"] == selected_class]

          if df.empty:
            st.info("Aucun coulage enregistré pour les critères sélectionnés.")
          else:
            cols_drop = [
                c
                for c in [
                    "id",
                    "created_at",
                    "created",
                    "heure_fin_coulage",
                    "client",
                    "centrale_beton",
                    "technicien",
                    "observations",
                    "nb_eprouvettes",
                    "duree_transport",
                    "duree_transport_min",
                    "Durée de transport",
                ]
                if c in df.columns
            ]
            df_display = df.drop(columns=cols_drop).rename(
                columns={
                    "date_livraison": "Date Livraison",
                    "heure_arrivee": "Heure d'arrivée",
                    "bl_num": "N° BL",
                    "ouvrage": "Ouvrage",
                    "quantite_m3": "Quantité (m³)",
                    "classe_beton": "Classe",
                    "temperature": "Temp. Béton (°C)",
                    "temperature_ambiante": "Temp. Ambiante (°C)",
                    "affaissement": "Affaissement (mm)",
                    "prelevement": "Prélèvement",
                    "meteo": "Météo",
                }
            )

            desired = [
                "Date Livraison",
                "Heure d'arrivée",
                "N° BL",
                "Ouvrage",
                "Quantité (m³)",
                "Classe",
                "Temp. Béton (°C)",
                "Temp. Ambiante (°C)",
                "Affaissement (mm)",
                "Prélèvement",
                "Météo",
            ]
            final_cols = [c for c in desired if c in df_display.columns]
            df_display = df_display[
                final_cols + [c for c in df_display.columns if c not in final_cols]
            ]

            st.markdown("---")
            k1, k2 = st.columns(2)
            k1.metric(
                "Volume Total", f"{df_display['Quantité (m³)'].sum():.1f} m³"
            )
            if "Affaissement (mm)" in df_display.columns:
              k2.metric(
                  "Affaissement Moyen",
                  f"{pd.to_numeric(df_display['Affaissement (mm)'], errors='coerce').mean():.0f}"
                  " mm",
              )
            st.markdown("---")

            st.download_button(
                label="📥 Télécharger la Synthèse Excel Bétonnage (A4 Portrait)",
                data=generate_excel_synthesis_betonnage(
                    df_display,
                    f"Journée du {selected_date.strftime('%d/%m/%Y')}",
                    is_mensuel=False,
                ),
                file_name=f"Synthese_Betonnage_{selected_date}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
            st.dataframe(df_display, use_container_width=True)
        else:
          st.info("Aucun coulage enregistré pour les critères sélectionnés.")
      except Exception as e:
        st.error(f"Erreur de chargement : {e}")

    with tab_m_b:
      st.markdown("### Bilan mensuel agrégé avec colonnes Min / Max séparées")
      col_m1, col_m2 = st.columns(2)
      with col_m1:
        annee = date.today().year
        mois_selected = st.selectbox(
            "Sélectionnez le mois :",
            MOIS_LISTE,
            index=date.today().month - 1,
            key="b_mois_m",
        )
        mois_num = MOIS_LISTE.index(mois_selected) + 1

      try:
        dernier_jour = (
            31
            if mois_num in [1, 3, 5, 7, 8, 10, 12]
            else (30 if mois_num in [4, 6, 9, 11] else 28)
        )
        res_m = (
            supabase.table("suivi_betonnage")
            .select("*")
            .gte("date_livraison", f"{annee}-{mois_num:02d}-01")
            .lte("date_livraison", f"{annee}-{mois_num:02d}-{dernier_jour}")
            .execute()
        )
        data_m = res_m.data if res_m else []

        classes_m = ["Toutes"]
        if data_m:
          df_m_temp = pd.DataFrame(data_m)
          if "classe_beton" in df_m_temp.columns:
            classes_m += sorted(
                list(df_m_temp["classe_beton"].dropna().unique())
            )

        with col_m2:
          selected_class_m = st.selectbox(
              "Filtrer par classe de béton (Mensuel) :",
              classes_m,
              key="b_class_m",
          )

        if data_m:
          df_m = pd.DataFrame(data_m)
          if selected_class_m != "Toutes":
            df_m = df_m[df_m["classe_beton"] == selected_class_m]

          if df_m.empty:
            st.info("Aucun coulage enregistré pour ce mois.")
          else:
            for col in [
                "quantite_m3",
                "temperature",
                "temperature_ambiante",
                "affaissement",
            ]:
              df_m[col] = pd.to_numeric(df_m[col], errors="coerce")
            df_m["date_dt"] = pd.to_datetime(
                df_m["date_livraison"], errors="coerce"
            )

            grouped_rows = []
            for (classe, ovr), group in df_m.groupby(
                ["classe_beton", "ouvrage"], dropna=False
            ):
              d_min, d_max = group["date_dt"].min(), group["date_dt"].max()
              date_str = (
                  "-"
                  if pd.isna(d_min)
                  else (
                      d_min.strftime("%d/%m/%Y")
                      if d_min == d_max
                      else (
                          f"{d_min.strftime('%d/%m/%Y')} -"
                          f" {d_max.strftime('%d/%m/%Y')}"
                      )
                  )
              )

              get_mm = lambda s: (
                  ("-", "-")
                  if s.dropna().empty
                  else (
                      int(round(s.dropna().min())),
                      int(round(s.dropna().max())),
                  )
              )
              aff_min, aff_max = get_mm(group["affaissement"])
              tb_min, tb_max = get_mm(group["temperature"])
              ta_min, ta_max = get_mm(group["temperature_ambiante"])

              grouped_rows.append({
                  ("Période", ""): date_str,
                  ("Classe", ""): classe,
                  ("Ouvrage", ""): ovr,
                  ("Quantité (m³)", ""): round(group["quantite_m3"].sum(), 1),
                  ("Affaissement", "Min"): aff_min,
                  ("Affaissement", "Max"): aff_max,
                  ("Temp. Béton (°C)", "Min"): tb_min,
                  ("Temp. Béton (°C)", "Max"): tb_max,
                  ("Temp. Ambiante (°C)", "Min"): ta_min,
                  ("Temp. Ambiante (°C)", "Max"): ta_max,
              })

            df_m_display = pd.DataFrame(grouped_rows)
            df_m_display.columns = pd.MultiIndex.from_tuples(
                df_m_display.columns
            )

            st.markdown("---")
            st.metric(
                "Volume Cumulé du Mois",
                f"{df_m_display[('Quantité (m³)', '')].sum():.1f} m³",
            )
            st.markdown("---")

            st.download_button(
                label=(
                    "📥 Télécharger la Synthèse Mensuelle Excel Bétonnage (A4"
                    " Portrait)"
                ),
                data=generate_excel_synthesis_betonnage(
                    df_m_display,
                    f"Mois de {mois_selected} {annee}",
                    is_mensuel=True,
                ),
                file_name=(
                    f"Synthese_Mensuelle_Betonnage_{mois_selected}_{annee}.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
            st.dataframe(df_m_display, use_container_width=True)
        else:
          st.info("Aucun coulage enregistré pour ce mois.")
      except Exception as e:
        st.error(f"Erreur de chargement : {e}")

  # --- TAB 2: CONTROLE ---
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
        selected_date_c = st.date_input(
            "Sélectionnez une date :", value=date.today(), key="c_date_j"
        )
      with col2:
        selected_class_cj = st.selectbox(
            "Filtrer par classe de béton :", classes_dispo, key="c_class_j"
        )

      if df_merged.empty:
        st.info("Aucune donnée disponible.")
      else:
        df_j_c = df_merged[df_merged["date_dt"].dt.date == selected_date_c]
        if (
            selected_class_cj != "Toutes"
            and "Classe Béton" in df_j_c.columns
        ):
          df_j_c = df_j_c[df_j_c["Classe Béton"] == selected_class_cj]

        if df_j_c.empty:
          st.info("Aucun contrôle enregistré pour les critères sélectionnés.")
        else:
          df_display_cj = format_controle_dataframe(df_j_c)

          st.markdown("---")
          st.metric("Nombre de Prélèvements", f"{len(df_display_cj)}")
          st.markdown("---")

          st.download_button(
              label="📥 Télécharger la Synthèse Contrôle Excel (A4 Portrait)",
              data=generate_excel_synthesis_controle(
                  df_display_cj,
                  f"Journée du {selected_date_c.strftime('%d/%m/%Y')}",
              ),
              file_name=f"Synthese_Controle_Beton_{selected_date_c}.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )
          st.dataframe(df_display_cj, use_container_width=True)

          st.markdown("### 📈 Synthèse Statistique")
          st.dataframe(
              compute_statistics_df(df_display_cj), use_container_width=True
          )

    with tab_m_c:
      st.markdown("### Bilan mensuel par classe de béton")
      col_m1, col_m2 = st.columns(2)
      with col_m1:
        annee_c = date.today().year
        mois_selected_c = st.selectbox(
            "Sélectionnez le mois :",
            MOIS_LISTE,
            index=date.today().month - 1,
            key="c_mois_m",
        )
        mois_num_c = MOIS_LISTE.index(mois_selected_c) + 1
      with col_m2:
        selected_class_cm = st.selectbox(
            "Filtrer par classe de béton :", classes_dispo, key="c_class_m"
        )

      if df_merged.empty:
        st.info("Aucune donnée disponible.")
      else:
        df_m_c = df_merged[
            (df_merged["date_dt"].dt.year == annee_c)
            & (df_merged["date_dt"].dt.month == mois_num_c)
        ]
        if (
            selected_class_cm != "Toutes"
            and "Classe Béton" in df_m_c.columns
        ):
          df_m_c = df_m_c[df_m_c["Classe Béton"] == selected_class_cm]

        if df_m_c.empty:
          st.info("Aucun contrôle enregistré pour ce mois.")
        else:
          df_display_cm = format_controle_dataframe(df_m_c)

          st.markdown("---")
          st.metric("Total Prélèvements du Mois", f"{len(df_display_cm)}")
          st.markdown("---")

          st.download_button(
              label=(
                  "📥 Télécharger la Synthèse Mensuelle Contrôle Excel (A4"
                  " Portrait)"
              ),
              data=generate_excel_synthesis_controle(
                  df_display_cm, f"Mois de {mois_selected_c} {annee_c}"
              ),
              file_name=(
                  f"Synthese_Mensuelle_Controle_{mois_selected_c}_{annee_c}.xlsx"
              ),
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )
          st.dataframe(df_display_cm, use_container_width=True)

          st.markdown("### 📈 Synthèse Statistique Mensuelle")
          st.dataframe(
              compute_statistics_df(df_display_cm), use_container_width=True
          )
