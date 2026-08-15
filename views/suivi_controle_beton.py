import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from supabase import create_client

# ---------------------------------------------------------
# INITIALISATION SUPABASE
# ---------------------------------------------------------
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "https://your-supabase-url.supabase.co")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "your-supabase-key")

@st.cache_resource
def init_supabase():
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

# ---------------------------------------------------------
# FONCTION REQUÊTE : RAPPEL DES SAMPLES ANTÉRIEURS (7j, 3j)
# ---------------------------------------------------------
def obtenir_historique_betonnage(betonnage_id):
    """
    Récupère toutes les éprouvettes déjà écrasées (force_kn > 0)
    associées au même bétonnage (betonnage_id).
    """
    if not betonnage_id:
        return []
    try:
        res = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .eq("betonnage_id", betonnage_id)
            .not_.is_("force_kn", "null")
            .gt("force_kn", 0)
            .order("echeance", desc=False)
            .execute()
        )
        return res.data if res.data else []
    except Exception as e:
        st.warning(f"Impossible de récupérer l'historique du bétonnage #{betonnage_id} : {e}")
        return []

# ---------------------------------------------------------
# FONCTION DE GÉNÉRATION DU PV EXCEL (LPEE)
# ---------------------------------------------------------
def generer_pv_excel(export_data, infos_header):
    """
    Génère un fichier Excel mis en forme selon les standards LPEE.
    Contient l'historique complet (3j, 7j, 28j) du lot de béton.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PV_Ecrasement"
    ws.views.sheetView[0].showGridLines = True

    # Styles & Couleurs
    font_title = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_sub_header = Font(name="Calibri", size=10, bold=True, color="000000")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="D3D3D3")
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # Header Document LPEE
    ws.merge_cells("A1:G1")
    ws["A1"] = "LABORATOIRE PUBLIC D'ESSAIS ET D'ETUDES - LPEE"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:G2")
    ws["A2"] = f"PROCÈS-VERBAL D'ESSAI COMPRESSION BETON (Réf: {infos_header.get('re_num', 'N/A')})"
    ws["A2"].font = Font(name="Calibri", size=11, bold=True, italic=True)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    # Méta-données du projet
    meta_info = [
        ("Client :", infos_header.get("client", "TGCC"), "Dossier N° :", infos_header.get("dossier", "N/A")),
        ("Ouvrage :", infos_header.get("ouvrage", "N/A"), "N° Bon Livraison :", infos_header.get("num_bl", "N/A")),
        ("Classe Béton :", infos_header.get("classe_beton", "C30/37"), "Date Coulée :", infos_header.get("date_coulee", "N/A")),
        ("Affaissement (mm) :", infos_header.get("affaissement", "200"), "Température (°C) :", infos_header.get("temperature", "31")),
    ]

    r = 4
    for row in meta_info:
        ws.cell(row=r, column=1, value=row[0]).font = font_bold
        ws.cell(row=r, column=2, value=row[1]).font = font_regular
        ws.cell(row=r, column=4, value=row[2]).font = font_bold
        ws.cell(row=r, column=5, value=row[3]).font = font_regular
        r += 1

    r += 1
    # Entêtes du Tableau d'Écrasement
    headers = ["Repère Éprouvette", "Forme", "Section (cm²)", "Âge (Jours)", "Date Écrasement", "Force (kN)", "Résistance Fc (MPa)"]
    for col_num, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=col_num, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    start_table_row = r + 1
    r += 1

    # Remplissage des données d'écrasement
    for idx, item in enumerate(export_data):
        ws.cell(row=r, column=1, value=item.get("repere_eprouvette")).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=item.get("forme")).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=3, value=float(item.get("section", 176.71))).number_format = "0.00"
        ws.cell(row=r, column=4, value=str(item.get("age"))).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5, value=str(item.get("date_essai"))).alignment = Alignment(horizontal="center")
        
        f_cell = ws.cell(row=r, column=6, value=float(item.get("force_kn", 0.0)))
        f_cell.number_format = "#,##0.0"
        
        fc_cell = ws.cell(row=r, column=7, value=float(item.get("fc_mpa", 0.0)))
        fc_cell.number_format = "0.0"

        # Application Zebra & Bordures
        for c in range(1, 8):
            cell_item = ws.cell(row=r, column=c)
            cell_item.border = border_all
            cell_item.font = font_regular
            if idx % 2 == 1:
                cell_item.fill = fill_zebra
        r += 1

    # Formule Moyenne Fc
    ws.cell(row=r, column=5, value="RÉSISTANCE MOYENNE (28J) :").font = font_bold
    ws.cell(row=r, column=5).alignment = Alignment(horizontal="right")
    mean_cell = ws.cell(row=r, column=7, value=f"=AVERAGE(G{start_table_row}:G{r-1})")
    mean_cell.font = font_bold
    mean_cell.number_format = "0.0"
    mean_cell.border = border_all
    mean_cell.fill = fill_sub

    r += 2
    # Observations
    ws.cell(row=r, column=1, value="Observations / Conclusions :").font = font_bold
    ws.cell(row=r+1, column=1, value=infos_header.get("observations", "N/A")).font = font_regular

    # Ajustement automatique des largeurs de colonnes
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# APPLICATION STREAMLIT (INTERFACE UTILISATEUR)
# =========================================================
st.set_page_config(page_title="Contrôle Béton LPEE - LGV", layout="wide")

st.title("🏗️ Suivi & Contrôle Qualité Béton - Projet LGV CASA")

tab_saisie, tab_hist = st.tabs(["💥 2. Saisie & Édition des PV", "📋 Historique Général"])

# ---------------------------------------------------------
# PHASE 2 : SAISIE DES ÉCRASEMENTS ET PV
# ---------------------------------------------------------
with tab_saisie:
    st.subheader("💥 2. Saisie Groupée & Édition des PV d'Écrasement")

    eprouvettes_en_attente = []
    try:
        res_att = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .order("id", desc=False)
            .execute()
        )
        if res_att.data:
            eprouvettes_en_attente = [
                e
                for e in res_att.data
                if e.get("force_kn") is None
                or float(e.get("force_kn") or 0) == 0
            ]
    except Exception as e:
        st.error(f"Erreur de chargement des essais en attente : {e}")

    if not eprouvettes_en_attente:
        st.info("👍 Aucune éprouvette en attente d'écrasement.")
    else:
        groupes_lots = {}
        for ep in eprouvettes_en_attente:
            b_id_ep = ep.get("betonnage_id")
            ech_ep = ep.get("echeance", "28 jours")
            ouv_ep = ep.get("ouvrage", "N/A")
            dt_ecras = ep.get("date_ecrasement", "N/A")

            cle_groupe = (
                f"Ouvrage: {ouv_ep} | Échéance: {ech_ep} (Date: {dt_ecras})"
                f" | Bétonnage ID #{b_id_ep}"
            )

            if cle_groupe not in groupes_lots:
                groupes_lots[cle_groupe] = []
            groupes_lots[cle_groupe].append(ep)

        choix_lot = st.selectbox(
            "📦 Sélectionner le lot d'éprouvettes à écraser :",
            list(groupes_lots.keys()),
            key="select_lot_saisie",
        )
        lot_selected = groupes_lots[choix_lot]

        sample = lot_selected[0]
        betonnage_id = sample.get("betonnage_id")

        # Chargement automatique des écrasements passés (ex: 3j, 7j) du même bétonnage
        essais_anterieurs = obtenir_historique_betonnage(betonnage_id)

        col_l1, col_l2, col_l3, col_l4 = st.columns(4)
        col_l1.metric("Client", "TGCC")
        col_l2.metric("Projet", "LGV CASA")
        col_l3.metric("Ouvrage", str(sample.get("ouvrage")))
        col_l4.metric("Échéance Visée", str(sample.get("echeance")))

        st.markdown("---")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            tech_global = st.text_input(
                "Technicien / Opérateur",
                value="Technicien LPEE",
                key="tech_global",
            )
        with col_g2:
            obs_globale = st.text_input(
                "Commentaire / Observation",
                value="PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES",
                key="obs_global",
            )

        st.markdown("##### 📝 Saisie des mesures pour le lot en cours")

        lot_key = f"df_lot_{choix_lot}"

        if lot_key not in st.session_state:
            rows_list = []
            for ep in lot_selected:
                sec = float(ep.get("section") or 176.71)
                f_kn = float(ep.get("force_kn") or 0.0)
                fc = (
                    round((f_kn * 10.0) / sec, 1)
                    if sec > 0 and f_kn > 0
                    else 0.0
                )

                rows_list.append({
                    "ID": ep["id"],
                    "Repère": ep.get("repere_eprouvette", f"/{ep['id']}"),
                    "Forme d'éprouvette": str(ep.get("forme") or "Cylindrique 150x300"),
                    "_section": sec,
                    "Force (kN)": f_kn,
                    "Résistance Fc (MPa)": fc,
                })
            st.session_state[lot_key] = pd.DataFrame(rows_list)

        def update_fc():
            changes = st.session_state.data_editor_ecrasement.get("edited_rows", {})
            for row_idx, updated_cols in changes.items():
                if "Force (kN)" in updated_cols:
                    new_force = float(updated_cols["Force (kN)"] or 0.0)
                    sec = float(st.session_state[lot_key].at[row_idx, "_section"])
                    st.session_state[lot_key].at[row_idx, "Force (kN)"] = new_force
                    if sec > 0 and new_force > 0:
                        st.session_state[lot_key].at[row_idx, "Résistance Fc (MPa)"] = round((new_force * 10.0) / sec, 1)
                    else:
                        st.session_state[lot_key].at[row_idx, "Résistance Fc (MPa)"] = 0.0

        st.data_editor(
            st.session_state[lot_key],
            column_config={
                "ID": st.column_config.NumberColumn("ID", disabled=True),
                "Repère": st.column_config.TextColumn("Repère", disabled=True),
                "Forme d'éprouvette": st.column_config.TextColumn("Forme d'éprouvette", disabled=True),
                "_section": None,
                "Force (kN)": st.column_config.NumberColumn(
                    "⚡ Force (kN)",
                    help="Saisissez la force lue",
                    min_value=0.0,
                    max_value=3000.0,
                    step=0.1,
                    format="%.1f",
                ),
                "Résistance Fc (MPa)": st.column_config.NumberColumn(
                    "💥 Résistance Fc (MPa)", disabled=True, format="%.1f"
                ),
            },
            use_container_width=True,
            hide_index=True,
            key="data_editor_ecrasement",
            on_change=update_fc,
        )

        df_actuel = st.session_state[lot_key]
        forces_valides = df_actuel[df_actuel["Force (kN)"] > 0]
        if not forces_valides.empty:
            fc_moy = round(forces_valides["Résistance Fc (MPa)"].mean(), 1)
            st.success(f"📈 **Résistance moyenne du lot actuel : {fc_moy:.1f} MPa**")

        # --- PRÉPARATION DE L'EXCEL D'EXPORT AVEC RAPPEL D'HISTORIQUE ---
        export_data = []

        # 1. Ajout des anciens écrasements (ex: 7 jours / 3 jours)
        if essais_anterieurs:
            st.info(f"ℹ️ {len(essais_anterieurs)} essai(s) antérieur(s) retrouvé(s) pour ce même béton (Bétonnage ID #{betonnage_id}) et inclus dans le PV.")
            for ep_ant in essais_anterieurs:
                sec_a = float(ep_ant.get("section") or 176.71)
                f_a = float(ep_ant.get("force_kn") or 0.0)
                fc_a = float(ep_ant.get("fc_mpa") or round((f_a * 10.0) / sec_a, 1))

                export_data.append({
                    "repere_eprouvette": ep_ant.get("repere_eprouvette", "N/A"),
                    "forme": ep_ant.get("forme", "Cylindrique 150x300"),
                    "section": sec_a,
                    "force_kn": f_a,
                    "fc_mpa": fc_a,
                    "date_essai": ep_ant.get("date_ecrasement", "N/A"),
                    "age": str(ep_ant.get("echeance", "7")).replace(" jours", "").replace("j", ""),
                })

        # 2. Ajout du lot actuel en cours de saisie (ex: 28 jours)
        for _, row in df_actuel.iterrows():
            export_data.append({
                "repere_eprouvette": row["Repère"],
                "forme": row["Forme d'éprouvette"],
                "section": row["_section"],
                "force_kn": row["Force (kN)"],
                "fc_mpa": row["Résistance Fc (MPa)"],
                "date_essai": sample.get("date_ecrasement", "N/A"),
                "age": str(sample.get("echeance", "28")).replace(" jours", "").replace("j", ""),
            })

        infos_header = {
            "re_num": "25/260/LGV/ B/01",
            "dossier": "2025-260-05985-2025-0247",
            "client": "TGCC",
            "num_bl": sample.get("num_bl", "15479"),
            "ouvrage": sample.get("ouvrage", "N/A"),
            "classe_beton": sample.get("classe_beton", "C30/37"),
            "date_coulee": sample.get("date_coulee", "02/06/2025"),
            "affaissement": sample.get("affaissement", "200"),
            "temperature": sample.get("temperature", "31"),
            "observations": obs_globale,
        }

        excel_file = generer_pv_excel(export_data, infos_header)
        filename = f"PV_Ecrasement_LPEE_{sample.get('num_bl', 'BL')}.xlsx"

        st.markdown("---")
        col_b1, col_b2 = st.columns(2)

        with col_b1:
            btn_enregistrer = st.button(
                "💾 Valider et Enregistrer Tout le Lot",
                type="primary",
                use_container_width=True,
            )

        with col_b2:
            st.download_button(
                label="📄 Télécharger le PV d'écrasement (Format LPEE)",
                data=excel_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if btn_enregistrer:
            if (df_actuel["Force (kN)"] == 0).any():
                st.error("❌ Une ou plusieurs forces d'écrasement sont à 0.0 kN. Veuillez saisir toutes les forces.")
            else:
                succes_lot = 0
                for _, row in df_actuel.iterrows():
                    update_payload = {
                        "force_kn": float(row["Force (kN)"]),
                        "fc_mpa": float(row["Résistance Fc (MPa)"]),
                        "technicien": tech_global,
                        "observations": obs_globale,
                    }
                    try:
                        supabase.table("suivi_controle_beton").update(update_payload).eq("id", int(row["ID"])).execute()
                        succes_lot += 1
                    except Exception as e:
                        st.error(f"Erreur sur l'éprouvette {row['Repère']} : {e}")

                if succes_lot == len(df_actuel):
                    st.balloons()
                    st.success(f"✅ Lot de {succes_lot} éprouvettes enregistré dans Supabase ! Vous pouvez télécharger le PV ci-dessus.")

# ---------------------------------------------------------
# HISTORIQUE & RE-TÉLÉCHARGEMENT DES PV
# ---------------------------------------------------------
with tab_hist:
    st.subheader("📋 Historique Général & Re-téléchargement des PV")
    try:
        res_all = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .order("id", desc=True)
            .execute()
        )
        if res_all.data:
            df_all = pd.DataFrame(res_all.data)

            # Extraction des essais validés (avec force > 0)
            df_valides = df_all[
                (df_all["force_kn"].notnull()) & (df_all["force_kn"] > 0)
            ].copy()

            if not df_valides.empty:
                st.markdown("##### 📥 Re-télécharger un PV déjà validé")

                groupes_valides = {}
                for _, row in df_valides.iterrows():
                    b_id_ep = row.get("betonnage_id")
                    ech_ep = row.get("echeance", "28 jours")
                    ouv_ep = row.get("ouvrage", "N/A")
                    dt_ecras = row.get("date_ecrasement", "N/A")

                    cle_pv = (
                        f"Ouvrage: {ouv_ep} | Échéance: {ech_ep} (Date: {dt_ecras}) | Lot ID #{b_id_ep}"
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

                # Récupération de tout l'historique associé au bétonnage
                tous_essais_hist = obtenir_historique_betonnage(b_id_h)

                export_data_h = []
                # Si des essais antérieurs existent, on alimente le rapport complet
                items_a_exporter = tous_essais_hist if tous_essais_hist else lot_hist

                for item in items_a_exporter:
                    sec = float(item.get("section") or 176.71)
                    f_kn = float(item.get("force_kn") or 0.0)
                    fc = float(item.get("fc_mpa") or 0.0)

                    export_data_h.append({
                        "repere_eprouvette": item.get("repere_eprouvette", f"/{item['id']}"),
                        "forme": item.get("forme", "Cylindrique 150x300"),
                        "section": sec,
                        "force_kn": f_kn,
                        "fc_mpa": fc,
                        "date_essai": item.get("date_ecrasement", "N/A"),
                        "age": str(item.get("echeance", "28")).replace(" jours", "").replace("j", ""),
                    })

                infos_header_h = {
                    "re_num": "25/260/LGV/ B/01",
                    "dossier": "2025-260-05985-2025-0247",
                    "client": "TGCC",
                    "num_bl": sample_h.get("num_bl", "15479"),
                    "ouvrage": sample_h.get("ouvrage", "N/A"),
                    "classe_beton": sample_h.get("classe_beton", "C30/37"),
                    "date_coulee": sample_h.get("date_coulee", "N/A"),
                    "affaissement": sample_h.get("affaissement", "200"),
                    "temperature": sample_h.get("temperature", "31"),
                    "observations": sample_h.get(
                        "observations",
                        "PERFORMANCES MECANIQUES A 28 JOURS SONT CONFORMES",
                    ),
                }

                excel_pv_hist = generer_pv_excel(export_data_h, infos_header_h)
                file_name_h = f"PV_Ecrasement_RE-EXPORT_{sample_h.get('num_bl', 'BL')}.xlsx"

                st.download_button(
                    label="📄 Télécharger le PV ré-énoncé (Excel Format LPEE)",
                    data=excel_pv_hist,
                    file_name=file_name_h,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_download_hist",
                )

            st.markdown("---")
            st.markdown("##### 📊 Base de données complète")
            st.dataframe(df_all, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun enregistrement d'écrasement dans la base.")
    except Exception as e:
        st.error(f"Erreur lors du chargement de l'historique : {e}")
