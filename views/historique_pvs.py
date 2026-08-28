import io
import re
import unicodedata
from datetime import datetime, date, timedelta
import pandas as pd
import streamlit as st

# ReportLab pour la génération du PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY


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


def verifier_doublon_num_reception(supabase, num_reception, current_beton_id=None):
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
        "num_bl", "bl", "num_bon_livraison", "n_bl", "bon_livraison",
        "num_bl_p", "n_bon", "bon_de_livraison", "code_bl"
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
    """Nettoie une chaîne pour qu'elle soit utilisable sans risque dans un nom de fichier."""
    chaine = str(chaine or "").strip().replace("/", "-").replace("\\", "-").replace(" ", "_")
    return re.sub(r"[^\w\-]", "", chaine)


# ==============================================================================
# 2. GÉNÉRATION DU PROCÈS-VERBAL PDF (FORMAT LPEE)
# ==============================================================================
def generer_pv_pdf(export_data, infos_header):
    """Génère le PV d'écrasement au format PDF vectoriel A4 selon la mise en page LPEE."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    story = []

    # Couleurs LPEE
    c_blue_dark = colors.HexColor("#1F4E78")
    c_blue_light = colors.HexColor("#D9E1F2")
    c_gray_light = colors.HexColor("#F2F2F2")
    c_black = colors.HexColor("#000000")

    # Styles Typo
    style_title = ParagraphStyle('TitleStyle', fontName='Helvetica-Bold', fontSize=10, textColor=colors.white, alignment=TA_CENTER)
    style_sub = ParagraphStyle('SubStyle', fontName='Helvetica-Bold', fontSize=8, textColor=colors.white, alignment=TA_CENTER)
    style_cell_bold = ParagraphStyle('CellBold', fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=TA_CENTER)
    style_cell_reg = ParagraphStyle('CellReg', fontName='Helvetica', fontSize=7.5, leading=9, alignment=TA_CENTER)
    style_cell_left = ParagraphStyle('CellLeft', fontName='Helvetica', fontSize=7.5, leading=9, alignment=TA_LEFT)
    style_cell_left_bold = ParagraphStyle('CellLeftBold', fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=TA_LEFT)
    style_cell_right_bold = ParagraphStyle('CellRightBold', fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=TA_RIGHT)

    default_bl = extraire_num_bl(infos_header)

    def clean_na(val, fallback=default_bl):
        v = str(val).strip() if val is not None else ""
        return fallback if v.upper() in ["N/A", "NONE", "NAN", "", "-"] else val

    # --- EN-TÊTE DU PV ---
    ref_h1 = clean_na(
        infos_header.get("num_reception")
        or infos_header.get("ref_controle")
        or infos_header.get("reference"),
        "B/406",
    )
    dossier_txt = clean_na(infos_header.get("dossier"), "2025-260-05985-2025-0247")
    client_txt = clean_na(infos_header.get("client"), "TGCC")
    re_num_txt = clean_na(infos_header.get("re_num"), "25/260/LGV/ B/")

    header_data = [
        [
            Paragraph("LPEE / CTR CSB", style_title), "", "", "",
            Paragraph("RE N° :", style_cell_bold),
            Paragraph(re_num_txt, style_cell_right_bold), "",
            Paragraph(str(ref_h1), style_cell_left_bold)
        ],
        [
            Paragraph("Laboratoire de Contrôle Externe", style_sub), "", "", "",
            Paragraph("DOSSIER :", style_cell_bold),
            Paragraph(dossier_txt, style_cell_left), "", ""
        ],
        [
            "", "", "", "",
            Paragraph("CLIENT :", style_cell_bold),
            Paragraph(f"<b>{client_txt}</b>", style_cell_left), "", ""
        ]
    ]

    t_header = Table(header_data, colWidths=[65, 65, 65, 65, 60, 90, 45, 100])
    t_header.setStyle(TableStyle([
        ('SPAN', (0, 0), (3, 0)),
        ('SPAN', (0, 1), (3, 2)),
        ('BACKGROUND', (0, 0), (3, 2), c_blue_dark),
        ('SPAN', (5, 0), (6, 0)),
        ('SPAN', (5, 1), (7, 1)),
        ('SPAN', (5, 2), (7, 2)),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, c_black),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 4))

    # --- TITRE PRINCIPAL & NORMES ---
    title_data = [
        [Paragraph("ESSAIS MECANIQUES SUR BETON HYDRAULIQUE", style_title)],
        [
            Paragraph("<b>[X] COMPRESSION NF EN 12390-3 (2019)</b>", ParagraphStyle('P1', parent=style_cell_bold, alignment=TA_LEFT)),
            Paragraph("<b>[ ] TRACTION PAR FENDAGE NF EN 12390-6 (2019)</b>", ParagraphStyle('P2', parent=style_cell_bold, alignment=TA_LEFT))
        ],
        [
            Paragraph("Presse : Marque: Controls", style_cell_right_bold),
            Paragraph("Classe : A", style_cell_bold)
        ]
    ]
    t_title = Table(title_data, colWidths=[277, 278])
    t_title.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('BACKGROUND', (0, 0), (1, 0), c_blue_dark),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, c_black),
    ]))
    story.append(t_title)
    story.append(Spacer(1, 4))

    # --- FICHE TECHNIQUE ---
    date_fab_header = clean_na(infos_header.get("date_coulee"), "-")
    lieu_prelev = clean_na(infos_header.get("lieu_prelevement", infos_header.get("ouvrage")), "-")
    chantier_txt = clean_na(infos_header.get("chantier"), "LGV-Gare Casa Sud")
    classe_b = str(clean_na(infos_header.get("classe_beton"), "C35/45")).upper()
    centrale_txt = clean_na(infos_header.get("centrale"), "Centrale à Béton")
    forme_txt = clean_na(infos_header.get("forme"), "Cylindrique 150x300")
    affaissement_txt = str(clean_na(infos_header.get("affaissement"), "-"))
    temp_txt = str(clean_na(infos_header.get("temperature"), "-"))
    tech_txt = clean_na(
        infos_header.get("technicien_prelevement")
        or infos_header.get("preleve_par")
        or infos_header.get("technicien"),
        "Technicien LPEE"
    )

    tech_data = [
        [
            Paragraph("<b>Date de<br/>prélèvement</b>", style_cell_bold),
            Paragraph(f"<b>{date_fab_header}</b>", style_cell_bold),
            Paragraph("<b>Lieu de<br/>prélèvement</b>", style_cell_bold), "",
            Paragraph(lieu_prelev, style_cell_left), "", "", ""
        ],
        [
            Paragraph("<b>Chantier</b>", style_cell_bold),
            Paragraph(chantier_txt, ParagraphStyle('SmallCh', parent=style_cell_left, fontSize=6.5, leading=7.5)), "", "",
            Paragraph("<b>Type de béton</b>", style_cell_bold), "",
            Paragraph(f"<b>{classe_b}</b>", style_cell_bold), ""
        ],
        [
            Paragraph(f"<b>{centrale_txt}</b>", style_cell_bold), "",
            Paragraph("- Dimensions", style_cell_left),
            Paragraph(f"<b>{forme_txt}</b>", style_cell_left_bold), "", "", "", ""
        ],
        [
            Paragraph("Affaissement au cône d'abrams NF EN 12350-2", ParagraphStyle('P3', parent=style_cell_left, fontSize=6.5)), "",
            Paragraph(f"<b>{affaissement_txt}</b>", style_cell_bold),
            Paragraph("- Mode confection", style_cell_left),
            Paragraph("<b>Par vibration NF EN 12390-2 (2019)</b>", style_cell_left), "", "", ""
        ],
        [
            Paragraph("Température °C", style_cell_bold), "",
            Paragraph(f"<b>{temp_txt}</b>", style_cell_bold),
            Paragraph("- Mode conservation", style_cell_left),
            Paragraph("<b>au laboratoire par immersion dans l'eau NF EN 12390-2 à 20°C ± 2°C</b>", style_cell_left), "", "", ""
        ],
        [
            Paragraph(f"prélèvement effectué par {tech_txt}", ParagraphStyle('P4', parent=style_cell_left, fontSize=6.5)), "", "",
            Paragraph("N° de bon de livraison", style_cell_bold), "",
            Paragraph(f"<b>{default_bl}</b>", style_cell_bold), "", ""
        ]
    ]

    t_tech = Table(tech_data, colWidths=[90, 70, 75, 80, 70, 55, 55, 60])
    t_tech.setStyle(TableStyle([
        ('SPAN', (2, 0), (3, 0)),
        ('SPAN', (4, 0), (7, 0)),
        ('SPAN', (1, 1), (3, 1)),
        ('SPAN', (4, 1), (5, 1)),
        ('SPAN', (6, 1), (7, 1)),
        ('SPAN', (0, 2), (1, 2)),
        ('SPAN', (3, 2), (7, 2)),
        ('SPAN', (0, 3), (1, 3)),
        ('SPAN', (4, 3), (7, 3)),
        ('SPAN', (0, 4), (1, 4)),
        ('SPAN', (4, 4), (7, 4)),
        ('SPAN', (0, 5), (2, 5)),
        ('SPAN', (3, 5), (4, 5)),
        ('SPAN', (5, 5), (7, 5)),
        ('BACKGROUND', (0, 0), (0, 0), c_gray_light),
        ('BACKGROUND', (2, 0), (3, 0), c_gray_light),
        ('BACKGROUND', (0, 1), (0, 1), c_gray_light),
        ('BACKGROUND', (4, 1), (5, 1), c_gray_light),
        ('BACKGROUND', (0, 2), (1, 2), c_gray_light),
        ('BACKGROUND', (0, 3), (1, 3), c_gray_light),
        ('BACKGROUND', (0, 4), (1, 4), c_gray_light),
        ('BACKGROUND', (0, 5), (2, 5), c_gray_light),
        ('BACKGROUND', (3, 5), (4, 5), c_gray_light),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, c_black),
    ]))
    story.append(t_tech)
    story.append(Spacer(1, 4))

    # --- TABLEAU DES RÉSULTATS D'ÉCRASEMENT ---
    rows_res = [
        [
            Paragraph("<b>Réf.</b>", style_cell_bold),
            Paragraph("<b>Date</b>", style_cell_bold), "",
            Paragraph("<b>Age<br/>(jours)</b>", style_cell_bold),
            Paragraph("<b>Charge<br/>rupture(KN)</b>", style_cell_bold),
            Paragraph("<b>Résistance (MPa)</b>", style_cell_bold), "", ""
        ],
        [
            "", Paragraph("Fabri", style_cell_reg), Paragraph("Essai", style_cell_reg),
            "", "", Paragraph("Compression", style_cell_reg), Paragraph("Traction", style_cell_reg), Paragraph("Moyenne", style_cell_reg)
        ]
    ]

    groupes_lots = {}
    row_index_start = 2

    for idx, item in enumerate(export_data):
        curr_r = row_index_start + idx
        f_kn = float(item.get("force_kn", 0.0) or 0.0)
        is_en_cours = (str(item.get("statut", "")).lower() == "en cours" or f_kn == 0.0)
        dt_essai = item.get("date_essai")

        age_val = calculer_age_jours(date_fab_header, dt_essai, item.get("age"))

        date_essai_affichage = "-"
        if not is_en_cours and dt_essai and str(dt_essai).strip() not in ["-", "", "None", "NaN"]:
            date_essai_affichage = str(clean_na(dt_essai, "-"))
        else:
            try:
                df_obj = datetime.strptime(str(date_fab_header).strip()[:10], "%Y-%m-%d")
                date_essai_affichage = (df_obj + timedelta(days=int(age_val))).strftime("%Y-%m-%d")
            except Exception:
                date_essai_affichage = "-"

        repere_txt = str(item.get("repere_eprouvette", "B/01"))
        force_str = "En cours" if is_en_cours else f"{f_kn:.1f}"
        fc_str = "En cours" if is_en_cours else f"{float(item.get('fc_mpa', 0.0)):.1f}"

        rows_res.append([
            Paragraph(repere_txt, style_cell_reg),
            Paragraph(str(date_fab_header), style_cell_reg),
            Paragraph(date_essai_affichage, style_cell_reg),
            Paragraph(str(age_val), style_cell_reg),
            Paragraph(force_str, style_cell_reg),
            Paragraph(fc_str, style_cell_reg),
            Paragraph("-", style_cell_reg),
            ""  # Moyenne calculée après fusion
        ])

        cle = f"{age_val}_{dt_essai}"
        groupes_lots.setdefault(cle, {"lignes": [], "en_cours": is_en_cours, "age": age_val})["lignes"].append(curr_r)

    t_res_style = [
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (2, 0)),
        ('SPAN', (3, 0), (3, 1)),
        ('SPAN', (4, 0), (4, 1)),
        ('SPAN', (5, 0), (7, 0)),
        ('BACKGROUND', (0, 0), (-1, 1), c_blue_light),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, c_black),
    ]

    a_des_28j, est_en_cours_28j, fc_28j_val = False, False, None

    for data in groupes_lots.values():
        lignes, age = data["lignes"], data["age"]
        start_r, end_r = min(lignes), max(lignes)

        if start_r != end_r:
            t_res_style.append(('SPAN', (7, start_r), (7, end_r)))

        if data["en_cours"]:
            rows_res[start_r][7] = Paragraph("<b>En cours</b>", style_cell_bold)
        else:
            valeurs_fc = []
            for r_idx in lignes:
                item_val = export_data[r_idx - row_index_start]
                valeurs_fc.append(float(item_val.get("fc_mpa", 0.0)))
            moy = sum(valeurs_fc) / len(valeurs_fc) if valeurs_fc else 0.0
            rows_res[start_r][7] = Paragraph(f"<b>{moy:.1f}</b>", style_cell_bold)

            if int(age) >= 28:
                fc_28j_val = moy

        if int(age) >= 28:
            a_des_28j = True
            if data["en_cours"]:
                est_en_cours_28j = True

    t_res = Table(rows_res, colWidths=[90, 70, 70, 50, 85, 65, 60, 65])
    t_res.setStyle(TableStyle(t_res_style))
    story.append(t_res)
    story.append(Spacer(1, 6))

    # --- COMMENTAIRE & VERDICT ---
    seuil = next(
        (s for k, s in [("C25/30", 25.0), ("C30/37", 30.0), ("C35/45", 35.0), ("C40/50", 40.0)] if k in classe_b),
        35.0
    )

    if not a_des_28j or est_en_cours_28j or fc_28j_val is None:
        comment_valeur = "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT."
    else:
        if fc_28j_val >= seuil:
            comment_valeur = "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES."
        else:
            comment_valeur = "PERFORMANCES MECANIQUES NON CONFORMES."

    comm_data = [
        [
            Paragraph("<b>Commentaire :</b>", style_cell_bold),
            Paragraph(f"<b>{comment_valeur}</b>", style_cell_left_bold)
        ]
    ]
    t_comm = Table(comm_data, colWidths=[90, 465])
    t_comm.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), c_gray_light),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, c_black),
    ]))
    story.append(t_comm)
    story.append(Spacer(1, 15))

    # --- VISAS DE SIGNATURE ---
    visas_data = [
        [
            "", Paragraph("<b>Visa Responsable d'essai</b>", style_cell_bold), "",
            "", Paragraph("<b>Visa Chef du laboratoire</b>", style_cell_bold), ""
        ],
        [
            "", Paragraph("<b>O.IKKEN</b>", ParagraphStyle('V1', parent=style_cell_bold, alignment=TA_CENTER)), "",
            "", Paragraph("<b>H.BAALLAL</b>", ParagraphStyle('V2', parent=style_cell_bold, alignment=TA_CENTER)), ""
        ]
    ]
    t_visas = Table(visas_data, colWidths=[30, 200, 65, 30, 200, 30], rowHeights=[20, 50])
    t_visas.setStyle(TableStyle([
        ('SPAN', (1, 0), (1, 0)),
        ('SPAN', (4, 0), (4, 0)),
        ('SPAN', (1, 1), (1, 1)),
        ('SPAN', (4, 1), (4, 1)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (1, 0), (1, 1), 0.5, c_black),
        ('GRID', (4, 0), (4, 1), 0.5, c_black),
    ]))

    story.append(KeepTogether(t_visas))

    # Génération
    doc.build(story)
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
    if num_rec and str(num_rec).strip().upper() not in ["", "-", "NONE", "NAN", "N/A"]:
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

        st.markdown("##### 📥 Re-télécharger un Procès-Verbal (PDF)")

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
                f"🔧 {len(lots_manquants)} PV marqué(s) validé(s) mais absent(s) de la liste ci-dessous",
                expanded=True,
            ):
                st.caption(
                    "Un PV validé n'apparaît dans la liste de téléchargement que si au moins une éprouvette de ce lot a une **Force (kN) > 0**."
                )
                st.dataframe(
                    pd.DataFrame(lots_manquants),
                    use_container_width=True,
                    hide_index=True,
                )

        if df_valides.empty:
            st.info("ℹ️ Aucun Procès-Verbal **validé** n'est disponible pour le téléchargement.")
        else:
            c_r1, c_r2 = st.columns(2)
            recherche_pv = c_r1.text_input(
                "🔍 Rechercher (réf, ouvrage, classe...)",
                placeholder="Ex: gare casa sud, B/421...",
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
                ref_ctrl = determiner_ref_controle(supabase, b_id, info_b, row.to_dict())
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

                    age_real = calculer_age_jours(date_coulee_h, dt_essai_item, item.get("age"))

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
                ref_ctrl_h = determiner_ref_controle(supabase, b_id_h, info_b_h, sample_h)

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
                    "affaissement": (info_b_h.get("affaissement") or info_b_h.get("slump")),
                    "temperature": (info_b_h.get("temperature") or info_b_h.get("temp_beton")),
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

                # Formatage du nom de fichier PDF : RefReception-DateFabrication.pdf (ex: B_421-13-08-2026.pdf)
                ref_clean = nettoyer_nom_fichier(ref_ctrl_h)
                
                # Inversion de la date en DD-MM-YYYY si sous format YYYY-MM-DD
                dt_fab_str = str(date_coulee_h or "").strip()
                try:
                    if len(dt_fab_str) >= 10:
                        dt_obj = datetime.strptime(dt_fab_str[:10], "%Y-%m-%d")
                        dt_fab_formatted = dt_obj.strftime("%d-%m-%Y")
                    else:
                        dt_fab_formatted = dt_fab_str
                except Exception:
                    dt_fab_formatted = dt_fab_str

                nom_fichier_pdf = f"PV_{ref_clean}-{dt_fab_formatted}.pdf"

                st.download_button(
                    label=f"📄 Télécharger le PV en PDF ({nom_fichier_pdf})",
                    data=generer_pv_pdf(export_data_h, infos_header_h),
                    file_name=nom_fichier_pdf,
                    mime="application/pdf",
                    use_container_width=True,
                    type="primary",
                    key="btn_download_hist_pdf",
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
        "id", "betonnage_id", "ref_controle", "repere_eprouvette",
        "num_bl", "ouvrage", "classe_beton", "statut_validation",
        "date_coulee", "affaissement_mm", "temp_beton_C", "echeance",
        "date_ecrasement", "fc_mpa", "technicien",
    ]
    exclus = {
        "forme", "section", "force_kn", "observations", "masse",
        "reference_controle", "refernce_controle", "num_reception",
    }

    cols_finales = [
        c for c in cols_ordre + [c for c in df_all.columns if c not in cols_ordre]
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
