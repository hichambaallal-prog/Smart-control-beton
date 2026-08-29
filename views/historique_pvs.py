import io
import re
import unicodedata
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle


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
  # Remplace les slashes, anti-slashes et caractères spéciaux par des tirets
  clean = re.sub(r'[\\/*?:"<>|]', "-", str(chaine).strip())
  # Supprime les espaces multiples
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
# 2. GÉNÉRATION DU PROCÈS-VERBAL PDF (FORMAT LPEE)
# ==============================================================================
@st.cache_data(show_spinner=False)
def generer_pv_pdf(export_data, infos_header):
  """Génère le PV d'écrasement en PDF, avec la même mise en page (mêmes
  sections, mêmes libellés, même grille) que l'ancienne version Excel.

  Mis en cache (st.cache_data) : sans ça, Streamlit régénère ce PDF à
  CHAQUE rerun du script — y compris pour une simple frappe dans un champ
  de recherche ailleurs sur la page — alors que le résultat est
  strictement identique tant que le PV sélectionné (et ses données) ne
  change pas. Le cache est automatiquement invalidé dès que les données
  d'entrée changent réellement (nouveau PV, force corrigée, etc.)."""
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
  base_widths = [16, 12, 12, 10, 18, 14, 12, 12]  # proportions A..H (comme Excel)
  total_units = sum(base_widths)
  col_widths = [page_width * (w / total_units) for w in base_widths]

  DARK = colors.HexColor("#1F4E78")
  TABLE_BG = colors.HexColor("#D9E1F2")
  LABEL_BG = colors.HexColor("#F2F2F2")
  WHITE = colors.white
  BLACK = colors.black

  def P(text, size=7.5, bold=False, align="CENTER", color=BLACK):
    """Cellule 'Paragraph' : contrairement à une simple chaîne de
    caractères, elle passe à la ligne automatiquement si le texte est trop
    long pour la largeur de la colonne (indispensable pour le Chantier,
    l'affaissement, etc. dont le texte dépasse largement une ligne)."""
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
  r[5] = clean_na(infos_header.get("re_num"), "25/260/LGV/ B/")
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
  r[0] = "ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
  data.append(r)
  row3 = len(data) - 1
  spans.append((0, row3, 7, row3))
  bg.append((0, row3, 7, row3, DARK))
  fonts.append((0, row3, 7, row3, "Helvetica-Bold", 11, WHITE))

  # ---- Row 4 : Compression / Traction ----
  r = blank_row()
  r[0] = "[X] COMPRESSION NF EN 12390-3 (2019)"
  r[4] = "[ ] TRACTION PAR FENDAGE NF EN 12390-6 (2019)"
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

  # ---- Lignes de résultats (une par éprouvette) ----
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

  # Fusion des moyennes (comme les cellules H fusionnées côté Excel)
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
  spans += [(1, row_visa_titre, 3, row_visa_titre), (5, row_visa_titre, 7, row_visa_titre)]
  fonts.append((1, row_visa_titre, 3, row_visa_titre, "Helvetica-Bold", 8.5, BLACK))
  fonts.append((5, row_visa_titre, 7, row_visa_titre, "Helvetica-Bold", 8.5, BLACK))

  r = blank_row()
  r[1] = "O.IKKEN"
  r[5] = "H.BAALLAL"
  data.append(r)
  row_visa_nom = len(data) - 1
  spans += [(1, row_visa_nom, 3, row_visa_nom), (5, row_visa_nom, 7, row_visa_nom)]
  fonts.append((1, row_visa_nom, 3, row_visa_nom, "Helvetica-Bold", 9, BLACK))
  fonts.append((5, row_visa_nom, 7, row_visa_nom, "Helvetica-Bold", 9, BLACK))
  valigns += [(1, row_visa_nom, 3, row_visa_nom, "TOP"), (5, row_visa_nom, 7, row_visa_nom, "TOP")]

  # ---- Construction de la table ----
  rows_auto_hauteur = {row6, row7, row9, row10, row11, row_comment}
  row_heights = []
  for i in range(len(data)):
    if i in rows_auto_hauteur:
      row_heights.append(None)  # calculé automatiquement selon le texte
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
  for (c1, r1, c2, r2) in spans:
    style_cmds.append(("SPAN", (c1, r1), (c2, r2)))
  for (c1, r1, c2, r2, color) in bg:
    style_cmds.append(("BACKGROUND", (c1, r1), (c2, r2), color))
  for (c1, r1, c2, r2, fname, fsize, fcolor) in fonts:
    style_cmds.append(("FONTNAME", (c1, r1), (c2, r2), fname))
    style_cmds.append(("FONTSIZE", (c1, r1), (c2, r2), fsize))
    style_cmds.append(("TEXTCOLOR", (c1, r1), (c2, r2), fcolor))
  for (c1, r1, c2, r2, al) in aligns:
    style_cmds.append(("ALIGN", (c1, r1), (c2, r2), al))
  for (c1, r1, c2, r2, va) in valigns:
    style_cmds.append(("VALIGN", (c1, r1), (c2, r2), va))
  table_style = TableStyle(style_cmds)

  # ---- Étirer le tableau pour qu'il couvre toute la page ----
  # Avec des hauteurs de ligne fixes, un PV à peu d'éprouvettes laisse un
  # grand vide sous le tableau à l'impression. On mesure d'abord la hauteur
  # réellement nécessaire (table1), puis on agrandit proportionnellement
  # TOUTES les lignes pour que le tableau final occupe toute la hauteur
  # imprimable — quel que soit le nombre d'éprouvettes.
  page_height_dispo = A4[1] - top_m - bottom_m

  table1 = Table(data, colWidths=col_widths, rowHeights=row_heights)
  table1.setStyle(table_style)
  _, hauteur_naturelle = table1.wrap(page_width, page_height_dispo * 10)

  if hauteur_naturelle > 0 and hauteur_naturelle < page_height_dispo:
    # table1._rowHeights contient les hauteurs réellement calculées (y
    # compris pour les lignes en hauteur automatique) après le wrap().
    hauteurs_reelles = list(table1._rowHeights)
    # Petite marge de sécurité (les cellules fusionnées ne redistribuent
    # pas toujours la hauteur de façon parfaitement linéaire).
    facteur = min((page_height_dispo * 0.97) / hauteur_naturelle, 3.5)
    row_heights_final = [h * facteur for h in hauteurs_reelles]

    # Vérification a posteriori : si malgré tout le résultat dépasse la
    # page (et déborderait sur une 2e page), on corrige le facteur une
    # dernière fois avant de construire la version définitive.
    table_verif = Table(data, colWidths=col_widths, rowHeights=row_heights_final)
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
  return buf.getvalue()




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


def obtenir_infos_betonnage_parents_bulk(supabase, betonnage_ids):
  """Charge en UNE seule requête les fiches parentes de plusieurs lots à la
  fois, au lieu d'une requête réseau par lot. C'est la principale cause de
  lenteur à l'ouverture de cette page : avec des dizaines de lots
  distincts, la version précédente déclenchait autant d'allers-retours
  réseau séquentiels rien que pour préparer l'affichage."""
  ids_valides = sorted({int(b) for b in betonnage_ids if pd.notnull(b)})
  if not ids_valides:
    return {}
  try:
    res = (
        supabase.table("suivi_betonnage")
        .select("*")
        .in_("id", ids_valides)
        .execute()
    )
    return {p["id"]: p for p in (res.data or [])}
  except Exception as e:
    st.warning(f"Note : chargement groupé des fiches parentes impossible ({e}).")
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
    unique_parents = obtenir_infos_betonnage_parents_bulk(supabase, unique_b_ids)

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

        date_coulee_h = info_b_h.get("date_coulee") or sample_h.get("date_coulee")

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

          # Calcul dynamique de l'âge spécifique pour chaque ligne
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

        # Formatage dynamique du nom du fichier : N°Réception_DateFabrication.pdf
        nom_rec_clean = nettoyer_nom_fichier(ref_ctrl_h)
        date_fab_clean = formater_date_nom_fichier(date_coulee_h)
        nom_fichier_pv = f"PV_{nom_rec_clean}_{date_fab_clean}.pdf"

        st.download_button(
            label=f"📄 Télécharger le PV ({nom_fichier_pv})",
            data=generer_pv_pdf(export_data_h, infos_header_h),
            file_name=nom_fichier_pv,
            mime="application/pdf",
            use_container_width=True,
            type="primary",
            key="btn_download_hist",
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
