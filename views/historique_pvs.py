import io
import re
import unicodedata
from datetime import date
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
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


# ==============================================================================
# 2. GÉNÉRATION DU PROCÈS-VERBAL EXCEL (FORMAT LPEE)
# ==============================================================================
def generer_pv_excel(export_data, infos_header):
  """Génère le PV d'écrasement Excel selon la mise en page LPEE."""
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "PV Écrasement LPEE"

  ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
  ws.page_setup.paperSize = ws.PAPERSIZE_A4
  ws.sheet_properties.pageSetUpPr.fitToPage = True
  ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
  ws.page_margins = PageMargins(
      left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.2, footer=0.2
  )

  # Styles
  f_bold = Font(name="Calibri", size=9, bold=True)
  f_bold_w = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
  f_title_w = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  f_reg = Font(name="Calibri", size=8.5)
  f_small = Font(name="Calibri", size=8)

  fill_dark = PatternFill("solid", fgColor="1F4E78")
  fill_table = PatternFill("solid", fgColor="D9E1F2")
  fill_label = PatternFill("solid", fgColor="F2F2F2")

  a_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
  a_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
  a_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
  a_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)

  thin = Side(border_style="thin", color="000000")
  b_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

  def set_cell(
      coord_or_cell,
      val=None,
      font=f_reg,
      align=a_center,
      fill=None,
      border=b_cell,
  ):
    cell = ws[coord_or_cell] if isinstance(coord_or_cell, str) else coord_or_cell
    if val is not None:
      cell.value = val
    cell.font, cell.alignment, cell.border = font, align, border
    if fill:
      cell.fill = fill

  default_bl = extraire_num_bl(infos_header)

  def clean_na(val, fallback=default_bl):
    v = str(val).strip() if val is not None else ""
    return fallback if v.upper() in ["N/A", "NONE", "NAN", "", "-"] else val

  # En-tête
  ws.merge_cells("A1:D1")
  set_cell("A1", "LPEE / CTR CSB", f_bold_w, fill=fill_dark)
  ws.merge_cells("A2:D3")
  set_cell("A2", "Laboratoire de Contrôle Externe", f_bold_w, fill=fill_dark)
  for r in range(1, 4):
    for c in range(1, 9):
      ws.cell(row=r, column=c).border = b_cell
      if c <= 4:
        ws.cell(row=r, column=c).fill = fill_dark

  set_cell("E1", "RE N° :", f_bold)
  ws.merge_cells("F1:G1")
  set_cell("F1", clean_na(infos_header.get("re_num"), "25/260/LGV/ B/"))

  ref_h1 = clean_na(
      infos_header.get("num_reception")
      or infos_header.get("ref_controle")
      or infos_header.get("reference"),
      "B/406",
  )
  set_cell("H1", ref_h1, f_bold)

  set_cell("E2", "DOSSIER :", f_bold)
  ws.merge_cells("F2:H2")
  set_cell(
      "F2",
      clean_na(infos_header.get("dossier"), "2025-260-05985-2025-0247"),
  )

  set_cell("E3", "CLIENT :", f_bold)
  ws.merge_cells("F3:H3")
  set_cell("F3", clean_na(infos_header.get("client"), "TGCC"), f_bold)

  # Titre & Normes
  ws.merge_cells("A4:H4")
  set_cell(
      "A4",
      "ESSAIS MECANIQUES SUR BETON HYDRAULIQUE",
      f_title_w,
      fill=fill_dark,
  )
  for c in range(1, 9):
    ws.cell(row=4, column=c).border = b_cell

  ws.merge_cells("A5:D5")
  set_cell("A5", "[X] COMPRESSION NF EN 12390-3 (2019)", f_bold)
  ws.merge_cells("E5:H5")
  set_cell("E5", "[ ] TRACTION PAR FENDAGE NF EN 12390-6 (2019)", f_bold)
  for c in range(1, 9):
    ws.cell(row=5, column=c).border = b_cell

  ws.merge_cells("A6:F6")
  set_cell("A6", "Presse : Marque: Controls", f_bold, a_right)
  ws.merge_cells("G6:H6")
  set_cell("G6", "Classe : A", f_bold)
  for c in range(1, 9):
    ws.cell(row=6, column=c).border = b_cell

  # Fiche technique
  set_cell("A7", "Date de\nprélèvement", f_bold, fill=fill_label)
  set_cell("B7", str(clean_na(infos_header.get("date_coulee"), "-")), f_bold)
  ws.merge_cells("C7:D7")
  set_cell("C7", "Lieu de\nprélèvement", f_bold, fill=fill_label)
  ws.merge_cells("E7:H7")
  set_cell(
      "E7",
      clean_na(
          infos_header.get("lieu_prelevement", infos_header.get("ouvrage")),
          "-",
      ),
  )

  set_cell("A8", "Chantier", f_bold, fill=fill_label)
  ws.merge_cells("B8:D8")
  set_cell(
      "B8",
      clean_na(
          infos_header.get("chantier"),
          "Augmentation de la capacité ferroviaire...",
      ),
      f_small,
  )
  ws.merge_cells("E8:F8")
  set_cell("E8", "Type de béton", f_bold, fill=fill_label)
  ws.merge_cells("G8:H8")
  set_cell(
      "G8",
      str(clean_na(infos_header.get("classe_beton"), "C35/45")).upper(),
      f_bold,
  )

  ws.merge_cells("A9:B9")
  set_cell(
      "A9",
      clean_na(infos_header.get("centrale"), "Centrale à Béton"),
      f_bold,
      fill=fill_label,
  )
  set_cell("C9", "- Dimensions", align=a_left)
  ws.merge_cells("D9:H9")
  set_cell(
      "D9",
      clean_na(infos_header.get("forme"), "Cylindrique 150x300"),
      f_bold,
  )

  ws.merge_cells("A10:B10")
  set_cell(
      "A10",
      "Affaissement au cône d'abrams NF EN 12350-2",
      f_small,
      fill=fill_label,
  )
  set_cell("C10", str(clean_na(infos_header.get("affaissement"), "-")), f_bold)
  set_cell("D10", "- Mode confection", align=a_left)
  ws.merge_cells("E10:H10")
  set_cell("E10", "Par vibration NF EN 12390-2 (2019)", f_bold)

  ws.merge_cells("A11:B11")
  set_cell("A11", "Température °C", fill=fill_label)
  set_cell("C11", str(clean_na(infos_header.get("temperature"), "-")), f_bold)
  set_cell("D11", "- Mode conservation", align=a_left)
  ws.merge_cells("E11:H11")
  set_cell(
      "E11",
      "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C ±"
      " 2°C",
      f_bold,
  )

  tech = clean_na(
      infos_header.get("technicien_prelevement")
      or infos_header.get("preleve_par")
      or infos_header.get("technicien"),
      "Technicien LPEE",
  )
  ws.merge_cells("A12:C12")
  set_cell("A12", f"prélèvement effectué par {tech}", f_small, fill=fill_label)
  ws.merge_cells("D12:E12")
  set_cell("D12", "N° de bon de livraison", fill=fill_label)
  ws.merge_cells("F12:H12")
  set_cell("F12", default_bl, f_bold)

  for r in range(7, 13):
    for c in range(1, 9):
      ws.cell(row=r, column=c).border = b_cell

  # Tableau des résultats
  headers = [
      ("A13:A14", "Réf,", f_bold),
      ("B13:C13", "Date", f_bold),
      ("B14", "Fabri", f_reg),
      ("C14", "Essai", f_reg),
      ("D13:D14", "Age (jours)", f_bold),
      ("E13:E14", "Charge rupture(KN)", f_bold),
      ("F13:H13", "Résistance (MPa)", f_bold),
      ("F14", "Compression", f_reg),
      ("G14", "Traction", f_reg),
      ("H14", "Moyenne", f_reg),
  ]
  for rng, text, font in headers:
    if ":" in rng:
      ws.merge_cells(rng)
    cell = rng.split(":")[0]
    set_cell(cell, text, font, fill=fill_table)

  for r in range(13, 15):
    for c in range(1, 9):
      ws.cell(row=r, column=c).border = b_cell

  row_start = 15
  groupes_lots = {}

  for idx, item in enumerate(export_data):
    r = row_start + idx
    f_kn = float(item.get("force_kn", 0.0) or 0.0)
    is_en_cours = (
        str(item.get("statut", "")).lower() == "en cours" or f_kn == 0.0
    )
    dt_essai = item.get("date_essai")
    age_val = (
        int(
            str(item.get("age", 7))
            .replace("j", "")
            .replace("jours", "")
            .strip()
        )
        if str(item.get("age", 7)).isdigit()
        else item.get("age", 7)
    )

    set_cell(f"A{r}", str(item.get("repere_eprouvette", "B/01")))
    set_cell(f"B{r}", str(clean_na(infos_header.get("date_coulee"), "-")))
    set_cell(f"C{r}", "En cours" if is_en_cours else str(clean_na(dt_essai, "-")))
    set_cell(f"D{r}", age_val)

    if is_en_cours:
      set_cell(f"E{r}", "En cours")
      set_cell(f"F{r}", "En cours")
    else:
      set_cell(f"E{r}", f_kn)
      ws[f"E{r}"].number_format = "0.0"
      set_cell(f"F{r}", float(item.get("fc_mpa", 0.0)))
      ws[f"F{r}"].number_format = "0.0"

    set_cell(f"G{r}", "-")
    for c in range(1, 9):
      ws.cell(row=r, column=c).border = b_cell

    cle = f"{item.get('age')}_{dt_essai}"
    groupes_lots.setdefault(
        cle, {"lignes": [], "en_cours": is_en_cours, "age": age_val}
    )["lignes"].append(r)

  # Fusion des moyennes & détection de la cellule moyenne à 28 jours
  a_des_28j, cel_moyenne_28j, est_en_cours_28j = False, None, False

  for data in groupes_lots.values():
    lignes, age = data["lignes"], data["age"]
    start_r, end_r = min(lignes), max(lignes)
    if start_r != end_r:
      ws.merge_cells(f"H{start_r}:H{end_r}")

    if data["en_cours"]:
      set_cell(f"H{start_r}", "En cours", f_bold)
    else:
      formula = (
          f"=ROUND(F{start_r}, 1)"
          if start_r == end_r
          else f"=ROUND(AVERAGE(F{start_r}:F{end_r}), 1)"
      )
      set_cell(f"H{start_r}", formula, f_bold)
      ws[f"H{start_r}"].number_format = "0.0"

    if str(age).isdigit() and int(age) >= 28:
      a_des_28j = True
      cel_moyenne_28j = f"H{start_r}"
      if data["en_cours"]:
        est_en_cours_28j = True

  # Commentaires & Visas
  next_r = row_start + len(export_data)
  set_cell(f"A{next_r}", "Commentaire :", f_bold, a_left, fill=fill_label)
  ws.merge_cells(f"B{next_r}:H{next_r}")

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

  # Construction sécurisée du commentaire pour éviter le crash de syntaxe XML d'Excel
  if not a_des_28j or est_en_cours_28j or not cel_moyenne_28j:
    comment_valeur = (
        "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT."
    )
  else:
    comment_valeur = (
        f'=IF(OR(ISBLANK({cel_moyenne_28j}), {cel_moyenne_28j}="En'
        ' cours"), "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES'
        f' ULTERIEUREMENT.", IF({cel_moyenne_28j}>={seuil}, "PERFORMANCES'
        ' MECANIQUES A 28 JOURS SONT CONFORMES", "PERFORMANCES MECANIQUES NON'
        ' CONFORMES"))'
    )

  set_cell(f"B{next_r}", comment_valeur, f_bold, a_left)
  for c in range(1, 9):
    ws.cell(row=next_r, column=c).border = b_cell

  # Visas
  r_titre = next_r + 2
  r_deb, r_fin = r_titre + 1, r_titre + 4
  ws.merge_cells(
      start_row=r_titre, start_column=2, end_row=r_titre, end_column=4
  )
  set_cell(ws.cell(row=r_titre, column=2), "Visa Responsable d'essai", f_bold)
  ws.merge_cells(start_row=r_deb, start_column=2, end_row=r_fin, end_column=4)
  set_cell(ws.cell(row=r_deb, column=2), "O.IKKEN", f_bold, a_top_center)

  ws.merge_cells(
      start_row=r_titre, start_column=6, end_row=r_titre, end_column=8
  )
  set_cell(ws.cell(row=r_titre, column=6), "Visa Chef du laboratoire", f_bold)
  ws.merge_cells(start_row=r_deb, start_column=6, end_row=r_fin, end_column=8)
  set_cell(ws.cell(row=r_deb, column=6), "H.BAALLAL", f_bold, a_top_center)

  # Dimensions
  heights = {7: 32, 8: 48, 10: 23, 11: 23, 9: 15, 12: 15, 13: 15, 14: 15}
  for r in range(1, r_fin + 1):
    ws.row_dimensions[r].height = heights.get(r, 28 if r >= 15 else 16)

  widths = {
      "A": 16,
      "B": 12,
      "C": 12,
      "D": 10,
      "E": 18,
      "F": 14,
      "G": 12,
      "H": 12,
  }
  for col, w in widths.items():
    ws.column_dimensions[col].width = w

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

          export_data_h.append({
              "repere_eprouvette": f"{ref_p}{rep_s}" if ref_p else rep_s,
              "forme": item.get("forme", "Cylindrique 150x300"),
              "section": sec,
              "force_kn": f_kn,
              "fc_mpa": fc,
              "date_essai": item.get("date_ecrasement", "-"),
              "age": (
                  str(item.get("age", "28"))
                  .replace(" jours", "")
                  .replace("j", "")
              ),
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
            "date_coulee": info_b_h.get("date_coulee")
            or sample_h.get("date_coulee"),
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

        st.download_button(
            label="📄 Télécharger le PV (Excel Format LPEE)",
            data=generer_pv_excel(export_data_h, infos_header_h),
            file_name=(
                f"PV_Ecrasement_LPEE_{num_bl_h if num_bl_h != '-' else 'BL'}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
