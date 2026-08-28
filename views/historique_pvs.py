import io
import re
import unicodedata
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
import streamlit as st


# ==============================================================================
# 1. GESTION UTILISATEURS & SUPABASE
# ==============================================================================
def connecter_utilisateur(supabase, nom_utilisateur, mot_de_passe):
  """Vérifie l'utilisateur et récupère ses droits depuis la table 'users'."""
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
      st.session_state.update({
          "user_logged": True,
          "user": user_info,
          "username": user_info.get("username"),
          "role": user_info.get("role"),
          "user_role": user_info.get("role"),
          "can_edit": bool(user_info.get("can_edit", False)),
      })
      return True
    st.error("Nom d'utilisateur ou mot de passe incorrect.")
    return False
  except Exception as e:
    st.error(f"Erreur lors de la connexion : {e}")
    return False


def verifier_doublon_num_reception(
    supabase, num_reception, current_beton_id=None
):
  """Vérifie l'existence unique du numéro de réception dans 'suivi_betonnage'."""
  num_clean = str(num_reception or "").strip()
  if not num_clean or num_clean.upper() in ["-", "NONE", "NAN", "N/A"]:
    return False
  try:
    res = (
        supabase.table("suivi_betonnage")
        .select("id, num_reception")
        .eq("num_reception", num_clean)
        .execute()
    )
    for m in res.data or []:
      if current_beton_id is None or int(m.get("id")) != int(current_beton_id):
        return True
  except Exception as e:
    st.warning(f"Note doublons : {e}")
  return False


def extraire_num_bl(*sources):
  """Extrait le N° BL depuis les sources de données."""
  clefs = {
      "num_bl",
      "bl",
      "num_bon_livraison",
      "n_bl",
      "bon_livraison",
      "num_bl_p",
      "n_bon",
      "bon_de_livraison",
      "code_bl",
  }
  invalid = {"N/A", "NONE", "NAN", "-", ""}
  for src in sources:
    if isinstance(src, dict):
      for k, v in src.items():
        if (k in clefs or "bl" in k.lower() or "bon" in k.lower()) and v:
          v_str = str(v).strip()
          if v_str.upper() not in invalid:
            return v_str
    elif isinstance(src, str):
      match = re.search(r"BL\s*:\s*([^\|]+)", src, re.IGNORECASE)
      if match:
        v_str = match.group(1).strip()
        if v_str.upper() not in invalid:
          return v_str
  return "-"


def calculer_age_jours(date_fab, date_ess, age_defaut=None):
  """Calcule l'âge en jours entre deux dates ou nettoie la valeur existante."""
  try:
    if date_fab and date_ess and str(date_ess).strip().lower() != "en cours":
      df = datetime.strptime(str(date_fab).strip()[:10], "%Y-%m-%d")
      de = datetime.strptime(str(date_ess).strip()[:10], "%Y-%m-%d")
      diff = (de - df).days
      if diff >= 0:
        return diff
  except Exception:
    pass

  if age_defaut is not None:
    val_clean = (
        str(age_defaut)
        .lower()
        .replace("jours", "")
        .replace("jour", "")
        .replace("j", "")
        .strip()
    )
    if val_clean.isdigit():
      return int(val_clean)

  return 28


def nettoyer_nom_fichier(chaine):
  """Remplace les caractères interdits pour les noms de fichiers OS."""
  if not chaine:
    return "PV"
  clean = re.sub(r'[\\/*?:"<>|]', "-", str(chaine).strip())
  return re.sub(r"\s+", "_", clean)


def formater_date_nom_fichier(dt_str):
  """Convertit 'YYYY-MM-DD' en 'DD-MM-YYYY' pour le nom du fichier."""
  if not dt_str or str(dt_str).strip() in ["-", "None", "NaN", "N/A"]:
    return "date_inconnue"
  try:
    dt_obj = datetime.strptime(str(dt_str).strip()[:10], "%Y-%m-%d")
    return dt_obj.strftime("%d-%m-%Y")
  except Exception:
    return str(dt_str).replace("/", "-")


# ==============================================================================
# 2. GENERATION DES PROCES-VERBAUX (PDF & EXCEL LPEE)
# ==============================================================================
def generer_pv_pdf(export_data, infos_header):
  """Génère le PV d'écrasement en PDF (Format officiel LPEE)."""
  buf = io.BytesIO()
  left_m = right_m = 0.3 * inch
  top_m = bottom_m = 0.4 * inch
  doc = SimpleDocTemplate(
      buf,
      pagesize=A4,
      leftMargin=left_m,
      rightMargin=right_m,
      topMargin=top_m,
      bottomMargin=bottom_m,
  )

  page_width = A4[0] - left_m - right_m
  base_widths = [16, 12, 12, 10, 18, 14, 12, 12]
  total_units = sum(base_widths)
  col_widths = [page_width * (w / total_units) for w in base_widths]

  DARK = colors.HexColor("#1F4E78")
  TABLE_BG = colors.HexColor("#D9E1F2")
  LABEL_BG = colors.HexColor("#F2F2F2")
  WHITE = colors.white
  BLACK = colors.black

  def P(text, size=7.5, bold=False, align="CENTER", color=BLACK):
    align_map = {"CENTER": TA_CENTER, "LEFT": TA_LEFT, "RIGHT": TA_RIGHT}
    style = ParagraphStyle(
        name="cell",
        fontName="Helvetica-Bold" if bold else "Helvetica",
        fontSize=size,
        leading=size * 1.2,
        alignment=align_map.get(align, TA_CENTER),
        textColor=color,
    )
    return Paragraph(str(text), style)

  default_bl = extraire_num_bl(infos_header)

  def clean_na(val, fallback=default_bl):
    v = str(val).strip() if val is not None else ""
    return fallback if v.upper() in ["N/A", "NONE", "NAN", "", "-"] else val

  def blank_row():
    return ["" for _ in range(8)]

  data = []
  spans, bg, fonts, aligns, valigns = [], [], [], [], []

  # ---- Row 0 : LPEE / CTR CSB | RE N° | Réf ----
  r = blank_row()
  r[0] = "LPEE / CTR CSB"
  r[4] = "RE N° :"
  r[5] = clean_na(infos_header.get("re_num"), "25/260/LGV/")
  ref_h1 = clean_na(
      infos_header.get("num_reception")
      or infos_header.get("ref_controle")
      or infos_header.get("reference"),
      "B/406",
  )
  r[7] = ref_h1
  data.append(r)
  row0 = len(data) - 1
  spans += [(0, row0, 3, row0), (5, row0, 6, row0)]
  bg.append((0, row0, 3, row0, DARK))
  fonts.append((0, row0, 3, row0, "Helvetica-Bold", 9, WHITE))
  fonts.append((4, row0, 4, row0, "Helvetica-Bold", 8.5, BLACK))
  aligns.append((5, row0, 6, row0, "RIGHT"))
  fonts.append((7, row0, 7, row0, "Helvetica-Bold", 8.5, BLACK))
  aligns.append((7, row0, 7, row0, "LEFT"))

  # ---- Rows 1-2 : Laboratoire de Contrôle Externe | DOSSIER / CLIENT ----
  r = blank_row()
  r[0] = "Laboratoire de Contrôle Externe"
  r[4] = "DOSSIER :"
  r[5] = clean_na(infos_header.get("dossier"), "2025-260-05985-2025-0247")
  data.append(r)
  row1 = len(data) - 1

  r = blank_row()
  r[4] = "CLIENT :"
  r[5] = clean_na(infos_header.get("client"), "TGCC")
  data.append(r)
  row2 = len(data) - 1

  spans.append((0, row1, 3, row2))
  bg.append((0, row1, 3, row2, DARK))
  fonts.append((0, row1, 3, row2, "Helvetica-Bold", 9, WHITE))
  spans += [(5, row1, 7, row1), (5, row2, 7, row2)]
  fonts.append((4, row1, 4, row1, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((4, row2, 4, row2, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((5, row1, 7, row1, "Helvetica", 8.5, BLACK))
  fonts.append((5, row2, 7, row2, "Helvetica-Bold", 8.5, BLACK))

  # ---- Row 3 : Titre ----
  r = blank_row()
  r[0] = "RAPPORT D'ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
  data.append(r)
  row3 = len(data) - 1
  spans.append((0, row3, 7, row3))
  bg.append((0, row3, 7, row3, DARK))
  fonts.append((0, row3, 7, row3, "Helvetica-Bold", 11, WHITE))

  # ---- Row 4 : Compression / Traction ----
  r = blank_row()
  r[0] = "[X] COMPRESSION NF EN 12390-3 (2019)"
  r[4] = "[ ] TRACTION PAR FENDAGE NF EN 12390-6 (2023)"
  data.append(r)
  row4 = len(data) - 1
  spans += [(0, row4, 3, row4), (4, row4, 7, row4)]
  fonts.append((0, row4, 7, row4, "Helvetica-Bold", 8.5, BLACK))

  # ---- Row 5 : Presse / Classe ----
  r = blank_row()
  r[0] = "Presse : Marque: Controls"
  r[6] = "Classe : A"
  data.append(r)
  row5 = len(data) - 1
  spans += [(0, row5, 5, row5), (6, row5, 7, row5)]
  fonts.append((0, row5, 7, row5, "Helvetica-Bold", 8.5, BLACK))
  aligns.append((0, row5, 5, row5, "RIGHT"))

  # ---- Row 6 : Date / Lieu de prélèvement ----
  date_fab_header = clean_na(infos_header.get("date_coulee"), "-")
  r = blank_row()
  r[0] = "Date de\nprélèvement"
  r[1] = str(date_fab_header)
  r[2] = "Lieu de\nprélèvement"
  r[4] = P(
      clean_na(
          infos_header.get("lieu_prelevement", infos_header.get("ouvrage")), "-"
      ),
      size=8.5,
      bold=True,
  )
  data.append(r)
  row6 = len(data) - 1
  spans += [(2, row6, 3, row6), (4, row6, 7, row6)]
  bg += [(0, row6, 0, row6, LABEL_BG), (2, row6, 3, row6, LABEL_BG)]
  fonts.append((0, row6, 0, row6, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((1, row6, 1, row6, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((2, row6, 3, row6, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((4, row6, 7, row6, "Helvetica", 8.5, BLACK))

  # ---- Row 7 : Chantier / Type de béton ----
  r = blank_row()
  r[0] = "Chantier"
  r[1] = P(
      clean_na(
          infos_header.get("chantier"),
          "LGV-Travaux d'exécution de terrassement, ouvrages d'art et"
          " rétablissement de communication entre PK 5+500 et PK"
          " 10+000-GARE CASA SUD.",
      ),
      size=7,
  )
  r[4] = "Type de béton"
  r[6] = str(clean_na(infos_header.get("classe_beton"), "C35/45")).upper()
  data.append(r)
  row7 = len(data) - 1
  spans += [(1, row7, 3, row7), (4, row7, 5, row7), (6, row7, 7, row7)]
  bg += [(0, row7, 0, row7, LABEL_BG), (4, row7, 5, row7, LABEL_BG)]
  fonts.append((0, row7, 0, row7, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((1, row7, 3, row7, "Helvetica", 7.5, BLACK))
  fonts.append((4, row7, 5, row7, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((6, row7, 7, row7, "Helvetica-Bold", 8.5, BLACK))

  # ---- Row 8 : Centrale / Dimensions ----
  r = blank_row()
  r[0] = clean_na(infos_header.get("centrale"), "Centrale à Béton")
  r[2] = "- Dimensions"
  r[3] = clean_na(infos_header.get("forme"), "Cylindrique 150x300")
  data.append(r)
  row8 = len(data) - 1
  spans += [(0, row8, 1, row8), (3, row8, 7, row8)]
  bg.append((0, row8, 1, row8, LABEL_BG))
  fonts.append((0, row8, 1, row8, "Helvetica-Bold", 8.5, BLACK))
  aligns.append((2, row8, 2, row8, "LEFT"))
  fonts.append((3, row8, 7, row8, "Helvetica-Bold", 8.5, BLACK))

  # ---- Row 9 : Affaissement / Mode confection ----
  r = blank_row()
  r[0] = P("Affaissement au cône d'abrams NF EN 12350-2", size=7)
  r[2] = str(clean_na(infos_header.get("affaissement"), "-"))
  r[3] = P("- Mode confection", size=7, align="LEFT")
  r[4] = "Par vibration NF EN 12390-2 (2019)"
  data.append(r)
  row9 = len(data) - 1
  spans += [(0, row9, 1, row9), (4, row9, 7, row9)]
  bg.append((0, row9, 1, row9, LABEL_BG))
  fonts.append((0, row9, 1, row9, "Helvetica", 7.5, BLACK))
  fonts.append((2, row9, 2, row9, "Helvetica-Bold", 8.5, BLACK))
  aligns.append((3, row9, 3, row9, "LEFT"))
  fonts.append((4, row9, 7, row9, "Helvetica-Bold", 8.5, BLACK))

  # ---- Row 10 : Température / Mode conservation ----
  r = blank_row()
  r[0] = "Température °C"
  r[2] = str(clean_na(infos_header.get("temperature"), "-"))
  r[3] = P("- Mode conservation", size=7, align="LEFT")
  r[4] = P(
      "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C"
      " ± 2°C",
      size=7.5,
      bold=True,
  )
  data.append(r)
  row10 = len(data) - 1
  spans += [(0, row10, 1, row10), (4, row10, 7, row10)]
  bg.append((0, row10, 1, row10, LABEL_BG))
  fonts.append((0, row10, 1, row10, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((2, row10, 2, row10, "Helvetica-Bold", 8.5, BLACK))
  aligns.append((3, row10, 3, row10, "LEFT"))
  fonts.append((4, row10, 7, row10, "Helvetica-Bold", 7.5, BLACK))

  # ---- Row 11 : Prélèvement effectué par / N° BL ----
  tech = clean_na(
      infos_header.get("technicien_prelevement")
      or infos_header.get("preleve_par")
      or infos_header.get("technicien"),
      "Technicien LPEE",
  )
  r = blank_row()
  r[0] = P(f"prélèvement effectué par {tech}", size=7)
  r[3] = "N° de bon de livraison"
  r[5] = default_bl
  data.append(r)
  row11 = len(data) - 1
  spans += [(0, row11, 2, row11), (3, row11, 4, row11), (5, row11, 7, row11)]
  bg += [(0, row11, 2, row11, LABEL_BG), (3, row11, 4, row11, LABEL_BG)]
  fonts.append((0, row11, 2, row11, "Helvetica", 7.5, BLACK))
  fonts.append((3, row11, 4, row11, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((5, row11, 7, row11, "Helvetica-Bold", 8.5, BLACK))

  # ---- Rows 12-13 : entête du tableau de résultats ----
  r = blank_row()
  r[0] = "Réf,"
  r[1] = "Date"
  r[3] = "Age (jours)"
  r[4] = "Charge rupture(KN)"
  r[5] = "Résistance (MPa)"
  data.append(r)
  row12 = len(data) - 1

  r = blank_row()
  r[1] = "Fabri"
  r[2] = "Essai"
  r[5] = "Compression"
  r[6] = "Traction"
  r[7] = "Moyenne"
  data.append(r)
  row13 = len(data) - 1

  spans += [
      (0, row12, 0, row13),
      (1, row12, 2, row12),
      (3, row12, 3, row13),
      (4, row12, 4, row13),
      (5, row12, 7, row12),
  ]
  bg.append((0, row12, 7, row13, TABLE_BG))
  fonts.append((0, row12, 7, row13, "Helvetica-Bold", 8.5, BLACK))

  # ---- Lignes de résultats ----
  row_indices_body = []
  groupes_lots = {}
  for item in export_data:
    f_kn = float(item.get("force_kn", 0.0) or 0.0)
    is_en_cours = (
        str(item.get("statut", "")).lower() == "en cours" or f_kn == 0.0
    )
    dt_essai = item.get("date_essai")
    age_val = calculer_age_jours(date_fab_header, dt_essai, item.get("age"))

    date_essai_affichage = "-"
    if (
        not is_en_cours
        and dt_essai
        and str(dt_essai).strip() not in ["-", "", "None", "NaN"]
    ):
      date_essai_affichage = str(clean_na(dt_essai, "-"))
    else:
      try:
        df_obj = datetime.strptime(
            str(date_fab_header).strip()[:10], "%Y-%m-%d"
        )
        date_essai_affichage = (
            df_obj + timedelta(days=int(age_val))
        ).strftime("%Y-%m-%d")
      except Exception:
        date_essai_affichage = "-"

    r = blank_row()
    r[0] = str(item.get("repere_eprouvette", "B/01"))
    r[1] = str(date_fab_header)
    r[2] = date_essai_affichage
    r[3] = str(age_val)
    if is_en_cours:
      r[4] = "En cours"
      r[5] = "En cours"
    else:
      r[4] = f"{f_kn:.1f}"
      r[5] = f"{float(item.get('fc_mpa', 0.0)):.1f}"
    r[6] = "-"
    data.append(r)
    r_idx = len(data) - 1
    row_indices_body.append(r_idx)
    fonts.append((0, r_idx, 7, r_idx, "Helvetica", 8.5, BLACK))

    cle = f"{age_val}_{dt_essai}"
    groupes_lots.setdefault(
        cle, {"lignes": [], "en_cours": is_en_cours, "age": age_val}
    )["lignes"].append(r_idx)

  # Fusion des moyennes
  a_des_28j, moyenne_28j_val, est_en_cours_28j = False, None, False
  for gdata in groupes_lots.values():
    lignes, age = gdata["lignes"], gdata["age"]
    start_r, end_r = min(lignes), max(lignes)
    if start_r != end_r:
      spans.append((7, start_r, 7, end_r))
    if gdata["en_cours"]:
      data[start_r][7] = "En cours"
    else:
      vals = []
      for li in lignes:
        try:
          vals.append(float(data[li][5]))
        except (ValueError, TypeError):
          pass
      moy = round(sum(vals) / len(vals), 1) if vals else 0.0
      data[start_r][7] = f"{moy:.1f}"
      if int(age) >= 28:
        moyenne_28j_val = moy
    fonts.append((7, start_r, 7, end_r, "Helvetica-Bold", 8.5, BLACK))
    if int(age) >= 28:
      a_des_28j = True
      if gdata["en_cours"]:
        est_en_cours_28j = True

  # ---- Commentaire de conformité ----
  seuil = next(
      (
          s
          for k, s in [
              ("C25/30", 25.0),
              ("C30/37", 30.0),
              ("C35/45", 35.0),
              ("C40/50", 40.0),
          ]
          if k
          in str(clean_na(infos_header.get("classe_beton"), "C35/45")).upper()
      ),
      35.0,
  )
  if not a_des_28j or est_en_cours_28j or moyenne_28j_val is None:
    comment_valeur = (
        "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT."
    )
  elif moyenne_28j_val >= seuil:
    comment_valeur = "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"
  else:
    comment_valeur = "PERFORMANCES MECANIQUES NON CONFORMES"

  r = blank_row()
  r[0] = "Commentaire :"
  r[1] = P(comment_valeur, size=8.5, bold=True, align="LEFT")
  data.append(r)
  row_comment = len(data) - 1
  spans.append((1, row_comment, 7, row_comment))
  bg.append((0, row_comment, 0, row_comment, LABEL_BG))
  fonts.append((0, row_comment, 0, row_comment, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((1, row_comment, 7, row_comment, "Helvetica-Bold", 8.5, BLACK))
  aligns.append((0, row_comment, 0, row_comment, "LEFT"))
  aligns.append((1, row_comment, 7, row_comment, "LEFT"))

  # ---- Visas ----
  r = blank_row()
  r[1] = "Visa Responsable d'essai"
  r[5] = "Visa Chef du laboratoire"
  data.append(r)
  row_visa_titre = len(data) - 1
  spans += [
      (1, row_visa_titre, 3, row_visa_titre),
      (5, row_visa_titre, 7, row_visa_titre),
  ]
  fonts.append(
      (1, row_visa_titre, 3, row_visa_titre, "Helvetica-Bold", 8.5, BLACK)
  )
  fonts.append(
      (5, row_visa_titre, 7, row_visa_titre, "Helvetica-Bold", 8.5, BLACK)
  )

  r = blank_row()
  r[1] = "O.IKKEN"
  r[5] = "H.BAALLAL"
  data.append(r)
  row_visa_nom = len(data) - 1
  spans += [
      (1, row_visa_nom, 3, row_visa_nom),
      (5, row_visa_nom, 7, row_visa_nom),
  ]
  fonts.append((1, row_visa_nom, 3, row_visa_nom, "Helvetica-Bold", 9, BLACK))
  fonts.append((5, row_visa_nom, 7, row_visa_nom, "Helvetica-Bold", 9, BLACK))
  valigns += [
      (1, row_visa_nom, 3, row_visa_nom, "TOP"),
      (5, row_visa_nom, 7, row_visa_nom, "TOP"),
  ]

  # ---- Construction de la table ----
  rows_auto_hauteur = {row6, row7, row9, row10, row11, row_comment}
  row_heights = []
  for i in range(len(data)):
    if i in rows_auto_hauteur:
      row_heights.append(None)
    elif i == row_visa_nom:
      row_heights.append(48)
    else:
      row_heights.append(16)

  style_cmds = [
      ("GRID", (0, 0), (-1, -1), 0.5, BLACK),
      ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ("ALIGN", (0, 0), (-1, -1), "CENTER"),
      ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
      ("FONTSIZE", (0, 0), (-1, -1), 8),
      ("TOPPADDING", (0, 0), (-1, -1), 2),
      ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
      ("LEFTPADDING", (0, 0), (-1, -1), 2),
      ("RIGHTPADDING", (0, 0), (-1, -1), 2),
  ]
  for c1, r1, c2, r2 in spans:
    style_cmds.append(("SPAN", (c1, r1), (c2, r2)))
  for c1, r1, c2, r2, color in bg:
    style_cmds.append(("BACKGROUND", (c1, r1), (c2, r2), color))
  for c1, r1, c2, r2, fname, fsize, fcolor in fonts:
    style_cmds.append(("FONTNAME", (c1, r1), (c2, r2), fname))
    style_cmds.append(("FONTSIZE", (c1, r1), (c2, r2), fsize))
    style_cmds.append(("TEXTCOLOR", (c1, r1), (c2, r2), fcolor))
  for c1, r1, c2, r2, al in aligns:
    style_cmds.append(("ALIGN", (c1, r1), (c2, r2), al))
  for c1, r1, c2, r2, va in valigns:
    style_cmds.append(("VALIGN", (c1, r1), (c2, r2), va))
  table_style = TableStyle(style_cmds)

  page_height_dispo = A4[1] - top_m - bottom_m
  table1 = Table(data, colWidths=col_widths, rowHeights=row_heights)
  table1.setStyle(table_style)
  _, hauteur_naturelle = table1.wrap(page_width, page_height_dispo * 10)

  if 0 < hauteur_naturelle < page_height_dispo:
    hauteurs_reelles = list(table1._rowHeights)
    facteur = min((page_height_dispo * 0.97) / hauteur_naturelle, 3.5)
    row_heights_final = [h * facteur for h in hauteurs_reelles]

    table_verif = Table(
        data, colWidths=col_widths, rowHeights=row_heights_final
    )
    table_verif.setStyle(table_style)
    _, hauteur_finale = table_verif.wrap(page_width, page_height_dispo * 10)
    if hauteur_finale > page_height_dispo:
      correction = (page_height_dispo * 0.97) / hauteur_finale
      row_heights_final = [h * correction for h in row_heights_final]
  else:
    row_heights_final = row_heights

  table = Table(data, colWidths=col_widths, rowHeights=row_heights_final)
  table.setStyle(table_style)

  doc.build([table])
  buf.seek(0)
  return buf


def exporter_pv_excel(export_data, infos_header):
  """Génère un fichier Excel du PV respectant scrupuleusement le layout LPEE."""
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "PV Écrasement"
  ws.views.sheetView[0].showGridLines = True

  # Styles
  font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
  font_bold = Font(name="Calibri", size=9, bold=True)
  font_regular = Font(name="Calibri", size=9)
  font_title = Font(name="Calibri", size=12, bold=True, color="FFFFFF")

  fill_dark = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  fill_label = PatternFill(
      start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
  )
  fill_table_hdr = PatternFill(
      start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
  )

  thin_side = Side(border_style="thin", color="000000")
  thin_border = Border(
      left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
  )

  align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
  align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
  align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

  default_bl = extraire_num_bl(infos_header)

  def clean_na(val, fallback=default_bl):
    v = str(val).strip() if val is not None else ""
    return fallback if v.upper() in ["N/A", "NONE", "NAN", "", "-"] else val

  date_fab_header = clean_na(infos_header.get("date_coulee"), "-")
  ref_h1 = clean_na(
      infos_header.get("num_reception")
      or infos_header.get("ref_controle")
      or infos_header.get("reference"),
      "B/406",
  )
  tech = clean_na(
      infos_header.get("technicien_prelevement")
      or infos_header.get("preleve_par")
      or infos_header.get("technicien"),
      "Technicien LPEE",
  )

  # Row 1
  ws.merge_cells("A1:D1")
  ws["A1"] = "LPEE / CTR CSB"
  ws["A1"].fill = fill_dark
  ws["A1"].font = font_header
  ws["A1"].alignment = align_center

  ws["E1"] = "RE N° :"
  ws["E1"].font = font_bold
  ws["E1"].alignment = align_right

  ws.merge_cells("F1:G1")
  ws["F1"] = clean_na(infos_header.get("re_num"), "25/260/LGV/")
  ws["F1"].font = font_regular
  ws["F1"].alignment = align_right

  ws["H1"] = ref_h1
  ws["H1"].font = font_bold
  ws["H1"].alignment = align_left

  # Row 2 & 3
  ws.merge_cells("A2:D3")
  ws["A2"] = "Laboratoire de Contrôle Externe"
  ws["A2"].fill = fill_dark
  ws["A2"].font = font_header
  ws["A2"].alignment = align_center

  ws["E2"] = "DOSSIER :"
  ws["E2"].font = font_bold
  ws["E2"].alignment = align_left
  ws.merge_cells("F2:H2")
  ws["F2"] = clean_na(infos_header.get("dossier"), "2025-260-05985-2025-0247")
  ws["F2"].font = font_regular

  ws["E3"] = "CLIENT :"
  ws["E3"].font = font_bold
  ws["E3"].alignment = align_left
  ws.merge_cells("F3:H3")
  ws["F3"] = clean_na(infos_header.get("client"), "TGCC")
  ws["F3"].font = font_bold

  # Row 4 : Titre
  ws.merge_cells("A4:H4")
  ws["A4"] = "RAPPORT D'ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
  ws["A4"].fill = fill_dark
  ws["A4"].font = font_title
  ws["A4"].alignment = align_center

  # Row 5 : Compression / Traction
  ws.merge_cells("A5:D5")
  ws["A5"] = "[X] COMPRESSION NF EN 12390-3 (2019)"
  ws["A5"].font = font_bold
  ws["A5"].alignment = align_center

  ws.merge_cells("E5:H5")
  ws["E5"] = "[ ] TRACTION PAR FENDAGE NF EN 12390-6 (2023)"
  ws["E5"].font = font_bold
  ws["E5"].alignment = align_center

  # Row 6 : Presse
  ws.merge_cells("A6:F6")
  ws["A6"] = "Presse : Marque: Controls"
  ws["A6"].font = font_bold
  ws["A6"].alignment = align_right

  ws.merge_cells("G6:H6")
  ws["G6"] = "Classe : A"
  ws["G6"].font = font_bold
  ws["G6"].alignment = align_center

  # Row 7 : Date & Lieu
  ws["A7"] = "Date de\nprélèvement"
  ws["A7"].fill = fill_label
  ws["A7"].font = font_bold
  ws["A7"].alignment = align_center

  ws["B7"] = str(date_fab_header)
  ws["B7"].font = font_bold
  ws["B7"].alignment = align_center

  ws.merge_cells("C7:D7")
  ws["C7"] = "Lieu de\nprélèvement"
  ws["C7"].fill = fill_label
  ws["C7"].font = font_bold
  ws["C7"].alignment = align_center

  ws.merge_cells("E7:H7")
  ws["E7"] = clean_na(
      infos_header.get("lieu_prelevement", infos_header.get("ouvrage")), "-"
  )
  ws["E7"].font = font_bold
  ws["E7"].alignment = align_center

  # Row 8 : Chantier & Type de béton
  ws["A8"] = "Chantier"
  ws["A8"].fill = fill_label
  ws["A8"].font = font_bold
  ws["A8"].alignment = align_center

  ws.merge_cells("B8:D8")
  ws["B8"] = clean_na(
      infos_header.get("chantier"),
      "LGV-Travaux d'exécution de terrassement, ouvrages d'art et"
      " rétablissement de communication entre PK 5+500 et PK 10+000-GARE CASA"
      " SUD.",
  )
  ws["B8"].font = font_regular
  ws["B8"].alignment = align_center

  ws.merge_cells("E8:F8")
  ws["E8"] = "Type de béton"
  ws["E8"].fill = fill_label
  ws["E8"].font = font_bold
  ws["E8"].alignment = align_center

  ws.merge_cells("G8:H8")
  ws["G8"] = str(clean_na(infos_header.get("classe_beton"), "C35/45")).upper()
  ws["G8"].font = font_bold
  ws["G8"].alignment = align_center

  # Row 9 : Centrale & Dimensions
  ws.merge_cells("A9:B9")
  ws["A9"] = clean_na(infos_header.get("centrale"), "Centrale à Béton")
  ws["A9"].fill = fill_label
  ws["A9"].font = font_bold
  ws["A9"].alignment = align_center

  ws["C9"] = "- Dimensions"
  ws["C9"].font = font_bold
  ws["C9"].alignment = align_left

  ws.merge_cells("D9:H9")
  ws["D9"] = clean_na(infos_header.get("forme"), "Cylindrique 150x300")
  ws["D9"].font = font_bold
  ws["D9"].alignment = align_center

  # Row 10 : Affaissement
  ws.merge_cells("A10:B10")
  ws["A10"] = "Affaissement au cône d'abrams NF EN 12350-2"
  ws["A10"].fill = fill_label
  ws["A10"].font = font_regular
  ws["A10"].alignment = align_center

  ws["C10"] = str(clean_na(infos_header.get("affaissement"), "-"))
  ws["C10"].font = font_bold
  ws["C10"].alignment = align_center

  ws["D10"] = "- Mode confection"
  ws["D10"].font = font_regular
  ws["D10"].alignment = align_left

  ws.merge_cells("E10:H10")
  ws["E10"] = "Par vibration NF EN 12390-2 (2019)"
  ws["E10"].font = font_bold
  ws["E10"].alignment = align_center

  # Row 11 : Température
  ws.merge_cells("A11:B11")
  ws["A11"] = "Température °C"
  ws["A11"].fill = fill_label
  ws["A11"].font = font_bold
  ws["A11"].alignment = align_center

  ws["C11"] = str(clean_na(infos_header.get("temperature"), "-"))
  ws["C11"].font = font_bold
  ws["C11"].alignment = align_center

  ws["D11"] = "- Mode conservation"
  ws["D11"].font = font_regular
  ws["D11"].alignment = align_left

  ws.merge_cells("E11:H11")
  ws["E11"] = (
      "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C"
      " ± 2°C"
  )
  ws["E11"].font = font_bold
  ws["E11"].alignment = align_center

  # Row 12 : Tech & BL
  ws.merge_cells("A12:C12")
  ws["A12"] = f"prélèvement effectué par {tech}"
  ws["A12"].fill = fill_label
  ws["A12"].font = font_regular
  ws["A12"].alignment = align_center

  ws.merge_cells("D12:E12")
  ws["D12"] = "N° de bon de livraison"
  ws["D12"].fill = fill_label
  ws["D12"].font = font_bold
  ws["D12"].alignment = align_center

  ws.merge_cells("F12:H12")
  ws["F12"] = default_bl
  ws["F12"].font = font_bold
  ws["F12"].alignment = align_center

  # Row 13 & 14 : En-têtes du Tableau de résultats
  for row in [13, 14]:
    for col in range(1, 9):
      cell = ws.cell(row=row, column=col)
      cell.fill = fill_table_hdr
      cell.font = font_bold
      cell.alignment = align_center

  ws.merge_cells("A13:A14")
  ws["A13"] = "Réf,"

  ws.merge_cells("B13:C13")
  ws["B13"] = "Date"
  ws["B14"] = "Fabri"
  ws["C14"] = "Essai"

  ws.merge_cells("D13:D14")
  ws["D13"] = "Age (jours)"

  ws.merge_cells("E13:E14")
  ws["E13"] = "Charge rupture(KN)"

  ws.merge_cells("F13:H13")
  ws["F13"] = "Résistance (MPa)"
  ws["F14"] = "Compression"
  ws["G14"] = "Traction"
  ws["H14"] = "Moyenne"

  # Injection des résultats & formules Excel
  cur_row = 15
  groupes_lots = {}
  a_des_28j, moyenne_28j_val, est_en_cours_28j = False, None, False

  for item in export_data:
    f_kn = float(item.get("force_kn", 0.0) or 0.0)
    is_en_cours = (
        str(item.get("statut", "")).lower() == "en cours" or f_kn == 0.0
    )
    dt_essai = item.get("date_essai")
    age_val = calculer_age_jours(date_fab_header, dt_essai, item.get("age"))

    if (
        not is_en_cours
        and dt_essai
        and str(dt_essai).strip() not in ["-", "", "None", "NaN"]
    ):
      date_essai_affichage = str(clean_na(dt_essai, "-"))
    else:
      try:
        df_obj = datetime.strptime(
            str(date_fab_header).strip()[:10], "%Y-%m-%d"
        )
        date_essai_affichage = (
            df_obj + timedelta(days=int(age_val))
        ).strftime("%Y-%m-%d")
      except Exception:
        date_essai_affichage = "-"

    ws[f"A{cur_row}"] = str(item.get("repere_eprouvette", "B/01"))
    ws[f"B{cur_row}"] = str(date_fab_header)
    ws[f"C{cur_row}"] = date_essai_affichage
    ws[f"D{cur_row}"] = age_val

    if is_en_cours:
      ws[f"E{cur_row}"] = "En cours"
      ws[f"F{cur_row}"] = "En cours"
    else:
      ws[f"E{cur_row}"] = f_kn
      ws[f"E{cur_row}"].number_format = "0.0"
      # Formule Excel pour la résistance : (Force * 10) / Section
      sec = float(item.get("section") or 176.71)
      ws[f"F{cur_row}"] = f"=(E{cur_row}*10)/{sec}"
      ws[f"F{cur_row}"].number_format = "0.0"

    ws[f"G{cur_row}"] = "-"

    for c in ["A", "B", "C", "D", "E", "F", "G"]:
      ws[f"{c}{cur_row}"].font = font_regular
      ws[f"{c}{cur_row}"].alignment = align_center

    cle = f"{age_val}_{dt_essai}"
    groupes_lots.setdefault(
        cle, {"lignes": [], "en_cours": is_en_cours, "age": age_val}
    )["lignes"].append(cur_row)
    cur_row += 1

  # Fusions et formules de moyenne Excel (colonne H)
  for gdata in groupes_lots.values():
    lignes, age = gdata["lignes"], gdata["age"]
    start_r, end_r = min(lignes), max(lignes)
    if start_r != end_r:
      ws.merge_cells(f"H{start_r}:H{end_r}")

    cell_h = ws[f"H{start_r}"]
    cell_h.font = font_bold
    cell_h.alignment = align_center

    if gdata["en_cours"]:
      cell_h.value = "En cours"
    else:
      cell_h.value = f"=AVERAGE(F{start_r}:F{end_r})"
      cell_h.number_format = "0.0"

      # Calcul de valeur secours pour commentaire si besoin
      vals = [
          float(export_data[i - 15].get("fc_mpa", 0.0))
          for i in lignes
          if float(export_data[i - 15].get("force_kn", 0.0) or 0.0) > 0
      ]
      if vals:
        moyenne_28j_val = round(sum(vals) / len(vals), 1)

    if int(age) >= 28:
      a_des_28j = True
      if gdata["en_cours"]:
        est_en_cours_28j = True

  # Commentaire de conformité
  seuil = next(
      (
          s
          for k, s in [
              ("C25/30", 25.0),
              ("C30/37", 30.0),
              ("C35/45", 35.0),
              ("C40/50", 40.0),
          ]
          if k
          in str(clean_na(infos_header.get("classe_beton"), "C35/45")).upper()
      ),
      35.0,
  )

  if not a_des_28j or est_en_cours_28j or moyenne_28j_val is None:
    comment_valeur = (
        "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT."
    )
  elif moyenne_28j_val >= seuil:
    comment_valeur = "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"
  else:
    comment_valeur = "PERFORMANCES MECANIQUES NON CONFORMES"

  ws[f"A{cur_row}"] = "Commentaire :"
  ws[f"A{cur_row}"].fill = fill_label
  ws[f"A{cur_row}"].font = font_bold
  ws[f"A{cur_row}"].alignment = align_left

  ws.merge_cells(f"B{cur_row}:H{cur_row}")
  ws[f"B{cur_row}"] = comment_valeur
  ws[f"B{cur_row}"].font = font_bold
  ws[f"B{cur_row}"].alignment = align_left

  cur_row += 1

  # Visas
  ws.merge_cells(f"B{cur_row}:D{cur_row}")
  ws[f"B{cur_row}"] = "Visa Responsable d'essai"
  ws[f"B{cur_row}"].font = font_bold
  ws[f"B{cur_row}"].alignment = align_center

  ws.merge_cells(f"F{cur_row}:H{cur_row}")
  ws[f"F{cur_row}"] = "Visa Chef du laboratoire"
  ws[f"F{cur_row}"].font = font_bold
  ws[f"F{cur_row}"].alignment = align_center

  cur_row += 1
  ws.row_dimensions[cur_row].height = 40

  ws.merge_cells(f"B{cur_row}:D{cur_row}")
  ws[f"B{cur_row}"] = "O.IKKEN"
  ws[f"B{cur_row}"].font = font_bold
  ws[f"B{cur_row}"].alignment = Alignment(horizontal="center", vertical="top")

  ws.merge_cells(f"F{cur_row}:H{cur_row}")
  ws[f"F{cur_row}"] = "H.BAALLAL"
  ws[f"F{cur_row}"].font = font_bold
  ws[f"F{cur_row}"].alignment = Alignment(horizontal="center", vertical="top")

  # Application universelle des bordures sur la grille complète
  for r in range(1, cur_row + 1):
    for c in range(1, 9):
      ws.cell(row=r, column=c).border = thin_border

  # Définition des largeurs de colonnes optimales
  col_widths = {
      "A": 16,
      "B": 14,
      "C": 14,
      "D": 12,
      "E": 20,
      "F": 16,
      "G": 14,
      "H": 14,
  }
  for col_letter, width in col_widths.items():
    ws.column_dimensions[col_letter].width = width

  # Mise en page pour l'impression (1 page de large)
  ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
  ws.page_setup.paperSize = ws.PAPERSIZE_A4
  ws.sheet_properties.pageSetUpPr.fitToPage = True
  ws.page_setup.fitToWidth = 1
  ws.page_setup.fitToHeight = 0

  buf = io.BytesIO()
  wb.save(buf)
  buf.seek(0)
  return buf


def exporter_dataframe_excel(df, date_chaine):
  """Export standard DataFrame vers Excel."""
  buf = io.BytesIO()
  with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name=f"Planning_{date_chaine}"[:31])
  buf.seek(0)
  return buf


# ==============================================================================
# 3. HELPER SUPABASE
# ==============================================================================
def obtenir_historique_betonnage(supabase, betonnage_id):
  """Récupère l'ensemble des essais pour un même bétonnage."""
  if not betonnage_id:
    return []
  try:
    res = (
        supabase.table("suivi_controle_beton")
        .select("*")
        .eq("betonnage_id", betonnage_id)
        .order("id")
        .execute()
    )
    return res.data or []
  except Exception:
    return []


def obtenir_infos_betonnage_parent(supabase, betonnage_id):
  """Récupère la fiche parent de suivi_betonnage."""
  if not betonnage_id:
    return {}
  try:
    res = (
        supabase.table("suivi_betonnage")
        .select("*")
        .eq("id", betonnage_id)
        .execute()
    )
    return res.data[0] if res.data else {}
  except Exception:
    return {}


def determiner_ref_controle(supabase, betonnage_id, info_betonnage, sample_ep):
  """Calcule la référence de contrôle prioritaire."""
  key = f"ref_controle_beton_{betonnage_id}"
  if key in st.session_state and st.session_state[key]:
    return st.session_state[key]

  num_rec = (info_betonnage or {}).get("num_reception")
  if num_rec and str(num_rec).strip().upper() not in [
      "",
      "-",
      "NONE",
      "NAN",
      "N/A",
  ]:
    ref = str(num_rec).strip()
  else:
    ref = (
        (info_betonnage or {}).get("ref_controle")
        or (sample_ep or {}).get("ref_controle")
        or f"REF-{betonnage_id}-{(info_betonnage or {}).get('ouvrage', 'N/A')}"
    ).strip()

  st.session_state[key] = ref
  return ref


# ==============================================================================
# 4. APPLICATION STREAMLIT
# ==============================================================================
def show(supabase):
  st.title("📋 Historique & Procès-Verbaux d'Écrasement (NF EN 12390)")

  user_info = st.session_state.get("user", {})
  role = str(
      st.session_state.get("user_role")
      or st.session_state.get("role")
      or user_info.get("role", "")
  ).lower()
  can_edit = st.session_state.get("can_edit", False) or bool(
      user_info.get("can_edit", False)
  )
  is_admin = (
      role in ["admin", "responsable_labo"]
      or st.session_state.get("is_admin", False)
  )

  if (
      role not in ["laboratoire", "labo", "admin", "responsable_labo", "qualite"]
      and not is_admin
      and not can_edit
  ):
    st.error("⛔ **Accès Restreint**")
    st.warning(
        "Ce module est réservé exclusivement au personnel du **Laboratoire de"
        " Contrôle**."
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
    if not res_all.data:
      st.info("ℹ️ Aucun historique disponible dans la base de données.")
      return

    df_all = pd.DataFrame(res_all.data)

    unique_b_ids = [
        b_id for b_id in df_all["betonnage_id"].unique() if pd.notnull(b_id)
    ]
    unique_parents = {
        b_id: obtenir_infos_betonnage_parent(supabase, b_id)
        for b_id in unique_b_ids
    }

    def est_valide_val(v):
      if isinstance(v, bool):
        return v
      if isinstance(v, str):
        v_norm = (
            unicodedata.normalize("NFKD", v.strip().lower())
            .encode("ascii", "ignore")
            .decode("ascii")
        ).strip()
        if v_norm in ["true", "1", "ok", "oui"]:
          return True
        if any(mot in v_norm for mot in ["invalide", "non valide", "rejet"]):
          return False
        return "valide" in v_norm
      return False

    def verifier_pv_valide_et_signe(row):
      b_id = row.get("betonnage_id")
      parent = unique_parents.get(b_id) or {}

      f_kn = row.get("force_kn")
      try:
        a_force = pd.notnull(f_kn) and float(f_kn) > 0
      except (ValueError, TypeError):
        a_force = False

      statut_valide = (
          est_valide_val(parent.get("statut_pv"))
          or est_valide_val(parent.get("validation_admin"))
          or est_valide_val(row.get("statut_pv"))
          or est_valide_val(row.get("validation_admin"))
      )

      return a_force and statut_valide

    mask_valides = df_all.apply(verifier_pv_valide_et_signe, axis=1)
    df_valides = df_all[mask_valides].copy()

    st.markdown("##### 📥 Re-télécharger un Procès-Verbal")

    b_ids_dans_liste = (
        set(df_valides["betonnage_id"].dropna().unique())
        if not df_valides.empty
        else set()
    )
    lots_manquants = []
    for b_id, info_b in unique_parents.items():
      statut_admin_valide = est_valide_val(
          info_b.get("statut_pv")
      ) or est_valide_val(info_b.get("validation_admin"))
      if statut_admin_valide and b_id not in b_ids_dans_liste:
        rows_lot = df_all[df_all["betonnage_id"] == b_id]
        a_au_moins_une_force = (
            bool((rows_lot["force_kn"].fillna(0).astype(float) > 0).any())
            if not rows_lot.empty
            else False
        )
        lots_manquants.append({
            "Lot ID": b_id,
            "Statut (admin)": info_b.get("statut_pv"),
            "Au moins 1 force > 0 ?": "Oui" if a_au_moins_une_force else "Non",
        })

    if lots_manquants:
      with st.expander(
          f"🔧 {len(lots_manquants)} PV marqué(s) validé(s) mais absent(s) de"
          " la liste ci-dessous",
          expanded=True,
      ):
        st.caption(
            "Un PV validé n'apparaît dans la liste de téléchargement que si"
            " au moins une éprouvette de ce lot a une **Force (kN) > 0**. Ces"
            " lots sont marqués validés côté admin mais aucune éprouvette du"
            " lot n'a encore de force enregistrée :"
        )
        st.dataframe(
            pd.DataFrame(lots_manquants),
            use_container_width=True,
            hide_index=True,
        )

    if df_valides.empty:
      st.info(
          "ℹ️ Aucun Procès-Verbal **validé** n'est disponible pour le"
          " téléchargement."
      )
    else:
      c_r1, c_r2 = st.columns(2)
      recherche_pv = c_r1.text_input(
          "🔍 Rechercher (réf, ouvrage, classe...)",
          placeholder="Ex: gare casa sud, B/394...",
          key="search_input_pv",
      )
      recherche_date_pv = c_r2.text_input(
          "📅 Rechercher par Date d'écrasement",
          placeholder="Ex: 2026-08-08",
          key="search_date_pv",
      )

      groupes_valides = {}
      for _, row in df_valides.iterrows():
        b_id = row.get("betonnage_id")
        info_b = unique_parents.get(b_id) or {}
        ref_ctrl = determiner_ref_controle(
            supabase, b_id, info_b, row.to_dict()
        )
        classe = (
            row.get("classe_beton")
            or (info_b.get("classe_beton") or info_b.get("classe") if info_b else "-")
            or "-"
        )

        cle = (
            f"Référence : {ref_ctrl} | Classe : {classe} | Ouvrage :"
            f" {row.get('ouvrage', '-')} | Échéance : {row.get('echeance', '28 jours')}"
            f" (Date : {row.get('date_ecrasement', '-')}) | Lot ID #{b_id}"
        )
        groupes_valides.setdefault(cle, []).append(row.to_dict())

      pvs_filtrés = [
          k
          for k in groupes_valides.keys()
          if (not recherche_pv or recherche_pv.lower() in k.lower())
          and (not recherche_date_pv or recherche_date_pv in k)
      ]

      if not pvs_filtrés:
        st.warning("Aucun PV validé ne correspond à votre recherche.")
      else:
        choix_pv = st.selectbox(
            "Sélectionnez le PV à consulter :",
            pvs_filtrés,
            key="select_pv_hist",
        )
        lot_hist = groupes_valides[choix_pv]
        sample_h = lot_hist[0]
        b_id_h = sample_h.get("betonnage_id")

        info_b_h = unique_parents.get(b_id_h) or {}
        essais_h = obtenir_historique_betonnage(supabase, b_id_h) or lot_hist

        date_coulee_h = info_b_h.get("date_coulee") or sample_h.get(
            "date_coulee"
        )

        export_data_h = []
        for item in essais_h:
          sec = float(item.get("section") or 176.71)
          f_kn = float(item.get("force_kn") or 0.0)
          fc = float(
              item.get("fc_mpa")
              or (round((f_kn * 10.0) / sec, 1) if f_kn > 0 else 0.0)
          )
          ref_p = str(item.get("ref_controle") or "").strip()
          rep_s = str(item.get("repere_eprouvette", f"/{item['id']}")).strip()
          dt_essai_item = item.get("date_ecrasement", "-")

          age_real = calculer_age_jours(
              date_coulee_h, dt_essai_item, item.get("age")
          )

          export_data_h.append({
              "repere_eprouvette": f"{ref_p}{rep_s}" if ref_p else rep_s,
              "forme": item.get("forme", "Cylindrique 150x300"),
              "section": sec,
              "force_kn": f_kn,
              "fc_mpa": fc,
              "date_essai": dt_essai_item,
              "age": age_real,
              "statut": "En cours" if f_kn == 0 else "Réalisé",
          })

        num_bl_h = extraire_num_bl(sample_h, info_b_h, choix_pv)
        ouv_h = info_b_h.get("ouvrage") or sample_h.get("ouvrage")
        ref_ctrl_h = determiner_ref_controle(
            supabase, b_id_h, info_b_h, sample_h
        )

        infos_header_h = {
            "re_num": "25/260/LGV/ B/",
            "dossier": "2025-260-05985-2025-0247",
            "client": "TGCC",
            "num_reception": ref_ctrl_h,
            "ref_controle": ref_ctrl_h,
            "num_bl": num_bl_h,
            "ouvrage": ouv_h,
            "lieu_prelevement": ouv_h,
            "classe_beton": sample_h.get("classe_beton", "C35/45"),
            "date_coulee": date_coulee_h,
            "affaissement": (
                info_b_h.get("affaissement") or info_b_h.get("slump")
            ),
            "temperature": (
                info_b_h.get("temperature") or info_b_h.get("temp_beton")
            ),
            "forme": sample_h.get("forme", "Cylindrique 150x300"),
            "centrale": (
                info_b_h.get("centrale")
                or info_b_h.get("centrale_beton")
                or sample_h.get("centrale")
            ),
            "observations": (
                info_b_h.get("observations_admin")
                or sample_h.get("observations")
            ),
            "technicien_prelevement": (
                info_b_h.get("technicien_prelevement")
                or info_b_h.get("preleve_par")
                or info_b_h.get("technicien")
                or sample_h.get("technicien")
            ),
        }

        nom_rec_clean = nettoyer_nom_fichier(ref_ctrl_h)
        date_fab_clean = formater_date_nom_fichier(date_coulee_h)

      col_pdf, col_excel = st.columns(2)

with col_pdf:
  nom_fichier_pv_pdf = f"PV_{nom_rec_clean}_{date_fab_clean}.pdf"
  st.download_button(
      label="📄 Télécharger le PV (PDF)",
      data=generer_pv_pdf(export_data_h, infos_header_h),
      file_name=nom_fichier_pv_pdf,
      mime="application/pdf",
      use_container_width=True,
      key=f"btn_pdf_{b_id_h}_{sample_h.get('echeance', '')}",
  )

with col_excel:
  nom_fichier_pv_xlsx = f"PV_{nom_rec_clean}_{date_fab_clean}.xlsx"
  st.download_button(
      label="📊 Télécharger le PV (Excel)",
      data=exporter_pv_excel(export_data_h, infos_header_h),
      file_name=nom_fichier_pv_xlsx,
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      use_container_width=True,
      key=f"btn_excel_{b_id_h}_{sample_h.get('echeance', '')}",
  )
    # Base de données globale
    st.markdown("---")
    st.markdown("##### 📊 Base de données globale")

    df_all["ref_controle"] = df_all.apply(
        lambda r: determiner_ref_controle(
            supabase,
            r.get("betonnage_id"),
            unique_parents.get(r.get("betonnage_id")),
            r.to_dict(),
        ),
        axis=1,
    )

    df_all["affaissement_mm"] = df_all["betonnage_id"].map(
        lambda b: (unique_parents.get(b) or {}).get("affaissement")
        or (unique_parents.get(b) or {}).get("slump")
        or "-"
    )
    df_all["temp_beton_C"] = df_all["betonnage_id"].map(
        lambda b: (unique_parents.get(b) or {}).get("temperature")
        or (unique_parents.get(b) or {}).get("temp_beton")
        or "-"
    )

    df_all["statut_validation"] = df_all.apply(
        lambda r: (
            "✅ Validé & Signé"
            if verifier_pv_valide_et_signe(r)
            else "⏳ En attente"
        ),
        axis=1,
    )

    cols_ordre = [
        "id",
        "betonnage_id",
        "ref_controle",
        "repere_eprouvette",
        "num_bl",
        "ouvrage",
        "classe_beton",
        "statut_validation",
        "date_coulee",
        "affaissement_mm",
        "temp_beton_C",
        "echeance",
        "date_ecrasement",
        "fc_mpa",
        "technicien",
    ]
    exclus = {
        "forme",
        "section",
        "force_kn",
        "observations",
        "masse",
        "reference_controle",
        "refernce_controle",
        "num_reception",
    }

    cols_finales = [
        c
        for c in cols_ordre
        + [c for c in df_all.columns if c not in cols_ordre]
        if c not in exclus
    ]
    df_final = df_all[cols_finales]

    c_s1, c_s2 = st.columns(2)
    search_ref = c_s1.text_input(
        "🔍 Recherche par Réf. Contrôle",
        placeholder="Ex: REF-123-GARE CASA SUD",
    )
    search_date = c_s2.text_input(
        "📅 Recherche par Date de coulée", placeholder="Ex: 2026-08-24"
    )

    if search_ref:
      df_final = df_final[
          df_final["ref_controle"]
          .astype(str)
          .str.contains(search_ref, case=False, na=False)
      ]
    if search_date:
      df_final = df_final[
          df_final["date_coulee"]
          .astype(str)
          .str.contains(search_date, case=False, na=False)
      ]

    st.dataframe(df_final, use_container_width=True, hide_index=True)

    st.download_button(
        label="📊 Télécharger la base de données globale (Excel)",
        data=exporter_dataframe_excel(df_final, "Historique_Global"),
        file_name=f"Historique_Global_Beton_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="btn_download_hist_global",
    )

  except Exception as e:
    st.error(f"Erreur lors du chargement de l'historique global : {e}")
