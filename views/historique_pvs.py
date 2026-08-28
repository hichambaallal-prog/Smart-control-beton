import io
import re
import unicodedata
from datetime import datetime, date, timedelta
import pandas as pd
import streamlit as st

# Importations ReportLab pour la génération du PDF
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether
)


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
        "num_bl", "bl", "num_bon_livraison", "n_bl",
        "bon_livraison", "num_bl_p", "n_bon", "bon_de_livraison", "code_bl"
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
# 2. GÉNÉRATION DU PROCÈS-VERBAL PDF (FORMAT LPEE)
# ==============================================================================
def generer_pv_pdf(export_data, infos_header):
    """Génère le PV d'écrasement PDF haute définition aux normes LPEE."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    elements = []

    # Couleurs LPEE
    DARK_BLUE = colors.HexColor("#1F4E78")
    LIGHT_BLUE = colors.HexColor("#D9E1F2")
    GRAY_LABEL = colors.HexColor("#F2F2F2")

    # Styles Typographiques
    styles = getSampleStyleSheet()
    
    style_normal = ParagraphStyle('Norm', fontName='Helvetica', fontSize=7.5, leading=9, alignment=1)
    style_normal_left = ParagraphStyle('NormLeft', fontName='Helvetica', fontSize=7.5, leading=9, alignment=0)
    style_bold = ParagraphStyle('Bld', fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=1)
    style_bold_left = ParagraphStyle('BldLeft', fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=0)
    style_bold_right = ParagraphStyle('BldRight', fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=2)
    style_title_white = ParagraphStyle('TitleWhite', fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1, textColor=colors.white)
    style_header_white = ParagraphStyle('HWhite', fontName='Helvetica-Bold', fontSize=9, leading=11, alignment=1, textColor=colors.white)

    default_bl = extraire_num_bl(infos_header)

    def clean_na(val, fallback=default_bl):
        v = str(val).strip() if val is not None else ""
        return fallback if v.upper() in ["N/A", "NONE", "NAN", "", "-"] else val

    ref_h1 = clean_na(
        infos_header.get("num_reception") or infos_header.get("ref_controle") or infos_header.get("reference"),
        "B/406"
    )

    # 1. En-tête LPEE (Tableau 3x8)
    data_header = [
        [
            Paragraph("<b>LPEE / CTR CSB</b>", style_header_white), "", "", "",
            Paragraph("<b>RE N° :</b>", style_bold_right),
            Paragraph("25/260/LGV/ B/", style_bold_right), "",
            Paragraph(f"<b>{ref_h1}</b>", style_bold_left)
        ],
        [
            Paragraph("<b>Laboratoire de Contrôle Externe</b>", style_header_white), "", "", "",
            Paragraph("<b>DOSSIER :</b>", style_bold_right),
            Paragraph(clean_na(infos_header.get("dossier"), "2025-260-05985-2025-0247"), style_normal_left), "", ""
        ],
        [
            "", "", "", "",
            Paragraph("<b>CLIENT :</b>", style_bold_right),
            Paragraph(f"<b>{clean_na(infos_header.get('client'), 'TGCC')}</b>", style_bold_left), "", ""
        ]
    ]

    t_header = Table(data_header, colWidths=[65, 65, 65, 65, 60, 90, 60, 85], rowHeights=[18, 14, 14])
    t_header.setStyle(TableStyle([
        ('SPAN', (0, 0), (3, 0)),
        ('SPAN', (0, 1), (3, 2)),
        ('SPAN', (5, 1), (7, 1)),
        ('SPAN', (5, 2), (7, 2)),
        ('BACKGROUND', (0, 0), (3, 2), DARK_BLUE),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    elements.append(t_header)

    # 2. Titre & Normes
    data_title = [
        [Paragraph("ESSAIS MECANIQUES SUR BETON HYDRAULIQUE", style_title_white)],
        [
            Paragraph("<b>[X] COMPRESSION NF EN 12390-3 (2019)</b>", style_bold),
            Paragraph("<b>[  ] TRACTION PAR FENDAGE NF EN 12390-6 (2019)</b>", style_bold)
        ],
        [
            Paragraph("<b>Presse : Marque: Controls</b>", style_bold_right),
            Paragraph("<b>Classe : A</b>", style_bold_left)
        ]
    ]

    t_title = Table(data_title, colWidths=[275, 280], rowHeights=[20, 16, 16])
    t_title.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('BACKGROUND', (0, 0), (1, 0), DARK_BLUE),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(t_title)

    # 3. Fiche Technique
    date_fab_header = clean_na(infos_header.get("date_coulee"), "-")
    tech = clean_na(
        infos_header.get("technicien_prelevement") or infos_header.get("preleve_par") or infos_header.get("technicien"),
        "Technicien LPEE"
    )

    data_tech = [
        [
            Paragraph("<b>Date de<br/>prélèvement</b>", style_bold),
            Paragraph(f"<b>{date_fab_header}</b>", style_bold),
            Paragraph("<b>Lieu de<br/>prélèvement</b>", style_bold), "",
            Paragraph(clean_na(infos_header.get("lieu_prelevement", infos_header.get("ouvrage")), "-"), style_normal_left), "", "", ""
        ],
        [
            Paragraph("<b>Chantier</b>", style_bold),
            Paragraph(clean_na(infos_header.get("chantier"), "LGV-Travaux d’exécution GARE CASA SUD."), style_normal_left), "", "",
            Paragraph("<b>Type de béton</b>", style_bold), "",
            Paragraph(f"<b>{str(clean_na(infos_header.get('classe_beton'), 'C35/45')).upper()}</b>", style_bold), ""
        ],
        [
            Paragraph(f"<b>{clean_na(infos_header.get('centrale'), 'Centrale à Béton')}</b>", style_bold), "",
            Paragraph("- Dimensions", style_normal_left),
            Paragraph(f"<b>{clean_na(infos_header.get('forme'), 'Cylindrique 150x300')}</b>", style_bold), "", "", "", ""
        ],
        [
            Paragraph("Affaissement au cône d'abrams NF EN 12350-2", style_normal), "",
            Paragraph(f"<b>{clean_na(infos_header.get('affaissement'), '-')}</b>", style_bold),
            Paragraph("- Mode confection", style_normal_left),
            Paragraph("<b>Par vibration NF EN 12390-2 (2019)</b>", style_bold), "", "", ""
        ],
        [
            Paragraph("Température °C", style_normal), "",
            Paragraph(f"<b>{clean_na(infos_header.get('temperature'), '-')}</b>", style_bold),
            Paragraph("- Mode conservation", style_normal_left),
            Paragraph("<b>au laboratoire par immersion dans l'eau NF EN 12390-2 à 20°C ± 2°C</b>", style_bold), "", "", ""
        ],
        [
            Paragraph(f"prélèvement effectué par {tech}", style_normal), "", "",
            Paragraph("N° de bon de livraison", style_normal), "",
            Paragraph(f"<b>{default_bl}</b>", style_bold), "", ""
        ]
    ]

    t_tech = Table(data_tech, colWidths=[70, 70, 65, 60, 70, 70, 75, 75], rowHeights=[22, 28, 16, 16, 16, 16])
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
        ('BACKGROUND', (0, 0), (0, 0), GRAY_LABEL),
        ('BACKGROUND', (2, 0), (3, 0), GRAY_LABEL),
        ('BACKGROUND', (0, 1), (0, 1), GRAY_LABEL),
        ('BACKGROUND', (4, 1), (5, 1), GRAY_LABEL),
        ('BACKGROUND', (0, 2), (1, 2), GRAY_LABEL),
        ('BACKGROUND', (0, 3), (1, 3), GRAY_LABEL),
        ('BACKGROUND', (0, 4), (1, 4), GRAY_LABEL),
        ('BACKGROUND', (0, 5), (2, 5), GRAY_LABEL),
        ('BACKGROUND', (3, 5), (4, 5), GRAY_LABEL),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    elements.append(t_tech)

    # 4. Tableau des Résultats
    data_res = [
        [
            Paragraph("<b>Réf.</b>", style_bold),
            Paragraph("<b>Date</b>", style_bold), "",
            Paragraph("<b>Age (jours)</b>", style_bold),
            Paragraph("<b>Charge rupture (kN)</b>", style_bold),
            Paragraph("<b>Résistance (MPa)</b>", style_bold), "", "",
        ],
        [
            "", Paragraph("Fabri", style_normal), Paragraph("Essai", style_normal),
            "", "", Paragraph("Compression", style_normal), Paragraph("Traction", style_normal), Paragraph("Moyenne", style_normal)
        ]
    ]

    groupes_lots = {}
    row_start_idx = 2

    for idx, item in enumerate(export_data):
        f_kn = float(item.get("force_kn", 0.0) or 0.0)
        is_en_cours = (str(item.get("statut", "")).lower() == "en cours" or f_kn == 0.0)
        dt_essai = item.get("date_essai")

        age_val = calculer_age_jours(date_fab_header, dt_essai, item.get("age"))

        if not is_en_cours and dt_essai and str(dt_essai).strip() not in ["-", "", "None", "NaN"]:
            date_essai_disp = str(clean_na(dt_essai, "-"))
        else:
            try:
                df_obj = datetime.strptime(str(date_fab_header).strip()[:10], "%Y-%m-%d")
                date_essai_disp = (df_obj + timedelta(days=int(age_val))).strftime("%Y-%m-%d")
            except Exception:
                date_essai_disp = "-"

        fc_val = float(item.get("fc_mpa", 0.0))
        str_kn = "En cours" if is_en_cours else f"{f_kn:.1f}"
        str_fc = "En cours" if is_en_cours else f"{fc_val:.1f}"

        data_res.append([
            Paragraph(str(item.get("repere_eprouvette", "B/01")), style_normal),
            Paragraph(str(date_fab_header), style_normal),
            Paragraph(date_essai_disp, style_normal),
            Paragraph(str(age_val), style_normal),
            Paragraph(str_kn, style_normal),
            Paragraph(str_fc, style_normal),
            Paragraph("-", style_normal),
            "" # Sera calculé/fusionné
        ])

        cle = f"{age_val}_{dt_essai}"
        groupes_lots.setdefault(cle, {"lignes": [], "en_cours": is_en_cours, "age": age_val, "fcs": []})
        groupes_lots[cle]["lignes"].append(row_start_idx + idx)
        if not is_en_cours:
            groupes_lots[cle]["fcs"].append(fc_val)

    # Calcul des Moyennes & Détection pour le commentaire
    a_des_28j, est_en_cours_28j, conformite_28j = False, False, True
    seuil = next(
        (s for k, s in [("C25/30", 25.0), ("C30/37", 30.0), ("C35/45", 35.0), ("C40/50", 40.0)]
         if k in str(clean_na(infos_header.get("classe_beton"), "C35/45")).upper()),
        35.0
    )

    t_res_style = [
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (2, 0)),
        ('SPAN', (3, 0), (3, 1)),
        ('SPAN', (4, 0), (4, 1)),
        ('SPAN', (5, 0), (7, 0)),
        ('BACKGROUND', (0, 0), (-1, 1), LIGHT_BLUE),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]

    for data in groupes_lots.values():
        lignes, age, is_ec = data["lignes"], data["age"], data["en_cours"]
        r_first, r_last = min(lignes), max(lignes)
        if r_first != r_last:
            t_res_style.append(('SPAN', (7, r_first), (7, r_last)))

        if is_ec:
            moy_text = "En cours"
        else:
            moy_val = round(sum(data["fcs"]) / len(data["fcs"]), 1) if data["fcs"] else 0.0
            moy_text = f"<b>{moy_val:.1f}</b>"

        data_res[r_first][7] = Paragraph(moy_text, style_bold)

        if int(age) >= 28:
            a_des_28j = True
            if is_ec:
                est_en_cours_28j = True
            else:
                if moy_val < seuil:
                    conformite_28j = False

    t_res = Table(data_res, colWidths=[80, 65, 65, 55, 80, 65, 65, 80])
    t_res.setStyle(TableStyle(t_res_style))
    elements.append(t_res)

    # 5. Commentaire
    if not a_des_28j or est_en_cours_28j:
        txt_comm = "PERFORMANCES MECANIQUES A 28 JOURS SERONT DONNES ULTERIEUREMENT."
    elif conformite_28j:
        txt_comm = "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES"
    else:
        txt_comm = "PERFORMANCES MECANIQUES NON CONFORMES"

    data_comment = [
        [
            Paragraph("<b>Commentaire :</b>", style_bold_left),
            Paragraph(f"<b>{txt_comm}</b>", style_bold_left)
        ]
    ]
    t_comment = Table(data_comment, colWidths=[90, 465], rowHeights=[22])
    t_comment.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), GRAY_LABEL),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(t_comment)

    elements.append(Spacer(1, 15))

    # 6. Visas et Signatures
    data_visas = [
        [
            "", Paragraph("<b>Visa Responsable d'essai</b>", style_bold),
            "", Paragraph("<b>Visa Chef du laboratoire</b>", style_bold), ""
        ],
        [
            "", Paragraph("<b>O.IKKEN</b>", style_bold),
            "", Paragraph("<b>H.BAALLAL</b>", style_bold), ""
        ]
    ]

    t_visas = Table(data_visas, colWidths=[40, 200, 75, 200, 40], rowHeights=[18, 45])
    t_visas.setStyle(TableStyle([
        ('GRID', (1, 0), (1, 1), 0.5, colors.black),
        ('GRID', (3, 0), (3, 1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))

    elements.append(KeepTogether(t_visas))

    # Construction du document
    doc.build(elements)
    buffer.seek(0)
    return buffer


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
        "", "-", "NONE", "NAN", "N/A"
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
                k for k in groupes_valides.keys()
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
                    "centrale": (info_b_h.get("centrale") or info_b_h.get("centrale_beton") or sample_h.get("centrale")),
                    "observations": (info_b_h.get("observations_admin") or sample_h.get("observations")),
                    "technicien_prelevement": (
                        info_b_h.get("technicien_prelevement")
                        or info_b_h.get("preleve_par")
                        or info_b_h.get("technicien")
                        or sample_h.get("technicien")
                    ),
                }

                # Nom du fichier dynamique : PV_N°Réception_DateFabrication.pdf
                nom_rec_clean = nettoyer_nom_fichier(ref_ctrl_h)
                date_fab_clean = formater_date_nom_fichier(date_coulee_h)
                nom_fichier_pdf = f"PV_{nom_rec_clean}_{date_fab_clean}.pdf"

                st.download_button(
                    label=f"📄 Télécharger le PV PDF ({nom_fichier_pdf})",
                    data=generer_pv_pdf(export_data_h, infos_header_h),
                    file_name=nom_fichier_pdf,
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
