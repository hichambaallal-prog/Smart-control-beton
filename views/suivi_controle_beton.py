import io
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st

# Importations openpyxl pour le style et la mise en page A4 Excel
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Importation sécurisée de PageMargins
try:
    from openpyxl.worksheet.margins import PageMargins
except ImportError:
    from openpyxl.worksheet.page_break import PageMargins


# =========================================================
# FONCTION : GÉNÉRATION DU PV EXCEL
# =========================================================
def generer_pv_excel(lot_data, infos_header):
    """Génère un fichier Excel mis en page au format A4 Portrait pour impression.

    - Client : TGCC
    - Projet : LGV CASA
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PV Écrasement"

    # --- CONFIGURATION IMPRESSION A4 PORTRAIT ---
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Marges réduites (CORRECT)
    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.6, bottom=0.6, header=0.3, footer=0.3
    )

    # --- STYLES ---
    font_title = Font(name="Arial", size=14, bold=True, color="1F497D")
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_norm = Font(name="Arial", size=10)

    fill_navy = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    fill_sub_header = PatternFill(
        start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"
    )

    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thick_bottom_side = Side(border_style="medium", color="1F497D")

    border_all = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )

    align_center = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )
    align_left = Alignment(horizontal="left", vertical="center")

    # --- ENTÊTE DU DOCUMENT / TITRE ---
    ws.merge_cells("A1:F1")
    ws["A1"] = "LABORATOIRE DE CONTRÔLE DE QUALITÉ DES BÉTONS"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:F2")
    ws["A2"] = (
        "PROCÈS-VERBAL D'ESSAI D'ÉCRASEMENT D'ÉPROUVETTES DE BÉTON (NF EN"
        " 12390-3)"
    )
    ws["A2"].font = Font(name="Arial", size=11, bold=True, color="333333")
    ws["A2"].alignment = align_center

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 18

    # Ligne de séparation
    for col in range(1, 7):
        ws.cell(row=3, column=col).border = Border(bottom=thick_bottom_side)

    # --- BLOC INFORMATIONS PROJET & CLIENT ---
    ws.merge_cells("A5:F5")
    ws["A5"] = " 📌 INFORMATIONS GÉNÉRALES DU PROJET & PRÉLÈVEMENT"
    ws["A5"].font = font_header
    ws["A5"].fill = fill_navy
    ws["A5"].alignment = align_left
    ws.row_dimensions[5].height = 20

    info_grid = [
        (
            "Client :",
            infos_header.get("client", "TGCC"),
            "Projet :",
            infos_header.get("projet", "LGV CASA"),
        ),
        (
            "N° Bon de Livraison (BL) :",
            infos_header.get("num_bl", "N/A"),
            "Ouvrage / Élément :",
            infos_header.get("ouvrage", "N/A"),
        ),
        (
            "Classe de Béton Spécifiée :",
            infos_header.get("classe_beton", "N/A"),
            "Échéance Visée :",
            infos_header.get("echeance", "N/A"),
        ),
        (
            "Date de Coulée :",
            str(infos_header.get("date_coulee", "N/A")),
            "Date d'Écrasement :",
            str(infos_header.get("date_ecrasement", "N/A")),
        ),
        (
            "Affaissement / Slump :",
            str(infos_header.get("affaissement", "N/A")),
            "Température Béton :",
            str(infos_header.get("temperature", "N/A")),
        ),
        (
            "Opérateur / Technicien :",
            infos_header.get("technicien", "Technicien LPEE"),
            "Observations :",
            infos_header.get("observations", "-"),
        ),
    ]

    row_idx = 6
    for label1, val1, label2, val2 in info_grid:
        ws.cell(row=row_idx, column=1, value=label1).font = font_bold
        ws.cell(row=row_idx, column=1).alignment = align_left
        ws.cell(row=row_idx, column=1).fill = fill_sub_header

        ws.cell(row=row_idx, column=2, value=val1).font = font_norm
        ws.cell(row=row_idx, column=2).alignment = align_left

        ws.cell(row=row_idx, column=4, value=label2).font = font_bold
        ws.cell(row=row_idx, column=4).alignment = align_left
        ws.cell(row=row_idx, column=4).fill = fill_sub_header

        ws.merge_cells(
            start_row=row_idx, start_column=5, end_row=row_idx, end_column=6
        )
        ws.cell(row=row_idx, column=5, value=val2).font = font_norm
        ws.cell(row=row_idx, column=5).alignment = align_left

        for c in range(1, 7):
            ws.cell(row=row_idx, column=c).border = border_all

        ws.row_dimensions[row_idx].height = 18
        row_idx += 1

    # --- TABLEAU DES RÉSULTATS D'ÉCRASEMENT ---
    row_idx += 2
    ws.merge_cells(
        start_row=row_idx, start_column=1, end_row=row_idx, end_column=6
    )
    ws.cell(row=row_idx, column=1, value=" 💥 RÉSULTATS DES ESSAIS D'ÉCRASEMENT").font = font_header
    ws.cell(row=row_idx, column=1).fill = fill_navy
    ws.cell(row=row_idx, column=1).alignment = align_left
    ws.row_dimensions[row_idx].height = 20

    row_idx += 1
    headers_tb = [
        "N° Éprouvette",
        "Forme / Dim.",
        "Section (cm²)",
        "Force F (kN)",
        "Résistance Fc (MPa)",
        "Conformité",
    ]
    for col_i, h_text in enumerate(headers_tb, 1):
        cell = ws.cell(row=row_idx, column=col_i, value=h_text)
        cell.font = font_bold
        cell.fill = fill_sub_header
        cell.alignment = align_center
        cell.border = border_all

    start_data_row = row_idx + 1
    row_idx += 1

    for item in lot_data:
        ws.cell(row=row_idx, column=1, value=item.get("repere_eprouvette", "")).alignment = align_center
        ws.cell(row=row_idx, column=2, value=item.get("forme", "")).alignment = align_center
        ws.cell(row=row_idx, column=3, value=item.get("section", 176.71)).alignment = align_center
        ws.cell(row=row_idx, column=4, value=item.get("force_kn", 0.0)).alignment = align_center
        ws.cell(row=row_idx, column=5, value=item.get("fc_mpa", 0.0)).alignment = align_center
        ws.cell(row=row_idx, column=6, value="Conforme").alignment = align_center

        for c in range(1, 7):
            ws.cell(row=row_idx, column=c).font = font_norm
            ws.cell(row=row_idx, column=c).border = border_all

        ws.row_dimensions[row_idx].height = 18
        row_idx += 1

    # Ligne Moyenne
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
    ws.cell(row=row_idx, column=1, value="Résistance Moyenne (MPa) :").alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(row=row_idx, column=1).font = font_bold
    
    end_data_row = row_idx - 1
    cell_moy = ws.cell(row=row_idx, column=5, value=f"=AVERAGE(E{start_data_row}:E{end_data_row})")
    cell_moy.font = font_bold
    cell_moy.alignment = align_center

    for c in range(1, 7):
        ws.cell(row=row_idx, column=c).border = border_all
        ws.cell(row=row_idx, column=c).fill = fill_sub_header

    # Ajustement largeur colonnes
    col_widths = [18, 22, 16, 16, 20, 16]
    for i, col_w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = col_w

    # RETOUR DU BUFFER BINAIRE
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================================================
# APPLICATION STREAMLIT (INTERFACE PRINCIPALE)
# =========================================================
tab_saisie, tab_hist = st.tabs(["📝 Saisie Écrasement", "📋 Historique"])

# ---------------------------------------------------------
# PHASE 2 : SAISIE GROUPÉE PAR ÉCHÉANCE / AGE
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
            "📦 Sélectionner le lot d'éprouvettes :",
            list(groupes_lots.keys()),
            key="select_lot_saisie",
        )
        lot_selected = groupes_lots[choix_lot]

        sample = lot_selected[0]
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
                "Observations générales",
                value="Rupture satisfaisante (NF EN 12390-3).",
                key="obs_global",
            )

        st.markdown("##### 📝 Saisie des mesures pour le lot")

        lot_key = f"df_lot_{choix_lot}"
        editor_key = f"editor_{choix_lot}"  # Clé dynamique unique par lot

        # Initialisation de st.session_state pour ce lot précis
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
                    "Repère": ep.get(
                        "repere_eprouvette", f"EP-{ep['id']}"
                    ),
                    "Forme d'éprouvette": str(
                        ep.get("forme") or "Cylindrique 150x300"
                    ),
                    "_section": sec,
                    "Force (kN)": f_kn,
                    "Résistance Fc (MPa)": fc,
                })
            st.session_state[lot_key] = pd.DataFrame(rows_list)

        # Callback de mise à jour sécurisé
        def update_fc():
            if editor_key in st.session_state:
                changes = st.session_state[editor_key].get("edited_rows", {})
                for row_idx, updated_cols in changes.items():
                    if "Force (kN)" in updated_cols:
                        try:
                            new_force = float(updated_cols["Force (kN)"] or 0.0)
                        except ValueError:
                            new_force = 0.0

                        sec = float(
                            st.session_state[lot_key].at[row_idx, "_section"]
                        )
                        st.session_state[lot_key].at[row_idx, "Force (kN)"] = new_force
                        
                        if sec > 0 and new_force > 0:
                            st.session_state[lot_key].at[
                                row_idx, "Résistance Fc (MPa)"
                            ] = round((new_force * 10.0) / sec, 1)
                        else:
                            st.session_state[lot_key].at[
                                row_idx, "Résistance Fc (MPa)"
                            ] = 0.0

        edited_df = st.data_editor(
            st.session_state[lot_key],
            column_config={
                "ID": st.column_config.NumberColumn("ID", disabled=True),
                "Repère": st.column_config.TextColumn("Repère", disabled=True),
                "Forme d'éprouvette": st.column_config.TextColumn(
                    "Forme d'éprouvette", disabled=True
                ),
                "_section": None,
                "Force (kN)": st.column_config.NumberColumn(
                    "⚡ Force (kN)",
                    help="Saisissez la force de rupture lue sur la presse",
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
            key=editor_key,
            on_change=update_fc,
        )

        df_actuel = st.session_state[lot_key]
        forces_valides = df_actuel[df_actuel["Force (kN)"] > 0]
        if not forces_valides.empty:
            fc_moy = round(forces_valides["Résistance Fc (MPa)"].mean(), 1)
            st.success(
                "📈 **Résistance moyenne calculée pour les éprouvettes"
                f" saisies : {fc_moy:.1f} MPa**"
            )
        else:
            st.warning(
                "👈 Veuillez remplir la colonne **Force (kN)** pour chaque"
                " éprouvette."
            )

        col_b1, col_b2 = st.columns(2)

        with col_b1:
            if st.button(
                "💾 Valider et Enregistrer Tout le Lot",
                type="primary",
                use_container_width=True,
            ):
                if (df_actuel["Force (kN)"] == 0).any():
                    st.error(
                        "❌ Attention : Une ou plusieurs éprouvettes ont"
                        " encore une force de 0.0 kN."
                    )
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
                            supabase.table("suivi_controle_beton").update(
                                update_payload
                            ).eq("id", int(row["ID"])).execute()
                            succes_lot += 1
                        except Exception as e:
                            st.error(
                                "Erreur sur l'éprouvette"
                                f" {row['Repère']} : {e}"
                            )

                    if succes_lot == len(df_actuel):
                        if lot_key in st.session_state:
                            del st.session_state[lot_key]
                        st.success(
                            f"✅ Lot de {succes_lot} éprouvettes"
                            " enregistré avec succès !"
                        )
                        st.rerun()

        with col_b2:
            export_data = []
            for _, row in df_actuel.iterrows():
                export_data.append({
                    "repere_eprouvette": row["Repère"],
                    "forme": row["Forme d'éprouvette"],
                    "section": row["_section"],
                    "force_kn": row["Force (kN)"],
                    "fc_mpa": row["Résistance Fc (MPa)"],
                })

            infos_header = {
                "client": "TGCC",
                "projet": "LGV CASA",
                "num_bl": sample.get("num_bl", "N/A"),
                "ouvrage": sample.get("ouvrage", "N/A"),
                "classe_beton": sample.get("classe_beton", "N/A"),
                "echeance": sample.get("echeance", "N/A"),
                "date_coulee": sample.get("date_coulee", "N/A"),
                "date_ecrasement": sample.get("date_ecrasement", "N/A"),
                "affaissement": sample.get("affaissement", "N/A"),
                "temperature": sample.get("temperature", "N/A"),
                "technicien": tech_global,
                "observations": obs_globale,
            }

            excel_file = generer_pv_excel(export_data, infos_header)
            filename = (
                f"PV_Ecrasement_TGCC_LGV_CASA_{sample.get('num_bl', 'BL')}_{sample.get('echeance', '28j')}.xlsx"
            )

            st.download_button(
                label="📄 Télécharger le PV d'écrasement (Excel A4)",
                data=excel_file,
                file_name=filename,
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

# ---------------------------------------------------------
# HISTORIQUE ET SUIVI GLOBAL
# ---------------------------------------------------------
with tab_hist:
    st.subheader("📋 Historique Général des Contrôles de Béton")

    try:
        res_all = (
            supabase.table("suivi_controle_beton")
            .select("*")
            .order("id", desc=True)
            .execute()
        )
        if res_all.data:
            df_all = pd.DataFrame(res_all.data)
            cols_display = [
                "id",
                "num_bl",
                "ouvrage",
                "classe_beton",
                "date_coulee",
                "echeance",
                "date_ecrasement",
                "repere_eprouvette",
                "forme",
                "force_kn",
                "fc_mpa",
                "technicien",
            ]
            cols_valid = [c for c in cols_display if c in df_all.columns]
            st.dataframe(
                df_all[cols_valid],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info(
                "Aucun enregistrement d'écrasement dans la base de"
                " données."
            )
    except Exception as e:
        st.error(f"Erreur lors du chargement de l'historique : {e}")
