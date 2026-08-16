import streamlit as st
import pandas as pd
from datetime import datetime, date
import io

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# FONCTION GENERATION EXCEL FORMAT A4 PORTRAIT
# =========================================================
def generate_excel_synthesis(df_data, titre_periode):
    """Génère un fichier Excel en A4 Portrait adapté au tableau de synthèse du contrôle béton."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Synthèse Contrôle Béton"

    # --- 1. CONFIGURATION D'IMPRESSION A4 PORTRAIT ---
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    # Marges d'impression
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

    # --- 2. PALETTE DE COULEURS ET POLICES ---
    color_primary = "1F4E79"    # Bleu LPEE / Marine
    color_header = "2D572C"     # Vert/Gris foncé entête
    color_card_bg = "F7F9FA"    # Fond clair cartes info
    color_kpi_bg = "EDF2F8"     # Fond KPI

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
    total_border = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    # Calcul des dimensions
    nb_cols = max(len(df_data.columns), 7)
    last_col_letter = get_column_letter(nb_cols)
    mid_col_idx = nb_cols // 2
    mid_col_letter = get_column_letter(mid_col_idx)
    next_mid_letter = get_column_letter(mid_col_idx + 1)

    # --- 3. BANNIÈRE EN-TÊTE LPEE ---
    ws.merge_cells(f"A1:{last_col_letter}2")
    cell_title = ws["A1"]
    cell_title.value = "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - CTR-CSB\nRAPPORT DE SYNTHÈSE DU CONTRÔLE BÉTON"
    cell_title.font = font_title
    cell_title.fill = fill_title
    cell_title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 25

    # --- 4. BLOC INFOS CLIENT & PROJET ---
    ws.merge_cells(f"A4:{mid_col_letter}4")
    cell_c = ws["A4"]
    cell_c.value = "   CLIENT :   TGCC"
    cell_c.font = font_bold
    cell_c.fill = fill_card
    cell_c.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"{next_mid_letter}4:{last_col_letter}4")
    cell_p = ws[f"{next_mid_letter}4"]
    cell_p.value = "   PROJET :   LGV CASA SUD"
    cell_p.font = font_bold
    cell_p.fill = fill_card
    cell_p.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"A5:{mid_col_letter}5")
    cell_per = ws["A5"]
    cell_per.value = f"   PÉRIODE :   {titre_periode}"
    cell_per.font = font_bold
    cell_per.fill = fill_card
    cell_per.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells(f"{next_mid_letter}5:{last_col_letter}5")
    cell_d = ws[f"{next_mid_letter}5"]
    cell_d.value = f"   DATE ÉDITION :   {datetime.now().strftime('%d/%m/%Y')}"
    cell_d.font = font_bold
    cell_d.fill = fill_card
    cell_d.alignment = Alignment(horizontal="left", vertical="center")

    for r in range(4, 6):
        ws.row_dimensions[r].height = 28
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    # --- 5. RÉSUMÉ DU NOMBRE DE CONTRÔLES ---
    row_idx = 7
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📊 RÉSUMÉ GLOBAL"
    ws[f"A{row_idx}"].font = font_section
    ws.row_dimensions[row_idx].height = 26

    row_idx += 1
    total_samples = len(df_data)

    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    cell_k_lbl = ws[f"A{row_idx}"]
    cell_k_lbl.value = "Nombre Total de Contrôles"
    cell_k_lbl.font = font_bold
    cell_k_lbl.fill = fill_kpi
    cell_k_lbl.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"A{row_idx+1}:{last_col_letter}{row_idx+1}")
    cell_k_val = ws[f"A{row_idx+1}"]
    cell_k_val.value = f"{total_samples} prélèvement(s)"
    cell_k_val.font = font_kpi_val
    cell_k_val.fill = fill_kpi
    cell_k_val.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(row_idx, row_idx + 2):
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    ws.row_dimensions[row_idx].height = 24
    ws.row_dimensions[row_idx+1].height = 30
    row_idx += 3

    # --- 6. TABLEAU DES DONNÉES ---
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📋 DÉTAIL DES ESSAIS ET ÉCRASEMENTS"
    ws[f"A{row_idx}"].font = font_section
    ws.row_dimensions[row_idx].height = 26
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

    # --- 7. PIED DE PAGE : ZONE DE SIGNATURES ---
    ws.merge_cells(f"A{row_idx}:{mid_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "Responsables d'essai"
    ws[f"A{row_idx}"].font = font_bold
    ws[f"A{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(f"{next_mid_letter}{row_idx}:{last_col_letter}{row_idx}")
    ws[f"{next_mid_letter}{row_idx}"] = "Chef du Laboratoire"
    ws[f"{next_mid_letter}{row_idx}"].font = font_bold
    ws[f"{next_mid_letter}{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_idx].height = 25

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
        ws.row_dimensions[r].height = 20

    for r in range(row_idx - 1, row_idx + 4):
        for c in range(1, mid_col_idx + 1):
            ws.cell(row=r, column=c).border = thin_border
        for c in range(mid_col_idx + 1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    # --- 8. LARGEUR SUR MESURE DES COLONNES ---
    col_widths = [20, 16, 16, 16, 18, 18, 18]
    for col_idx, width in enumerate(col_widths, 1):
        if col_idx <= nb_cols:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# FONCTION AUXILIAIRE DE TRAITEMENT ET CONSOLIDATION
# =========================================================
def load_and_process_data(supabase):
    """Charge les données de bétonnage et d'écrasement, puis calcule les résistances moyennes."""
    res_betonnage = supabase.table("suivi_betonnage").select("*").execute()
    res_ecrasement = supabase.table("suivi_controle_beton").select("*").execute()

    df_beton = pd.DataFrame(res_betonnage.data) if res_betonnage.data else pd.DataFrame()
    df_ecrasement = pd.DataFrame(res_ecrasement.data) if res_ecrasement.data else pd.DataFrame()

    if df_beton.empty:
        return pd.DataFrame()

    # Détermination de la colonne de référence et conversion de la date
    ref_col = "ref_controle" if "ref_controle" in df_beton.columns else ("prelevement" if "prelevement" in df_beton.columns else None)
    date_col = "date_livraison" if "date_livraison" in df_beton.columns else ("date_prelevement" if "date_prelevement" in df_beton.columns else None)

    if date_col:
        df_beton["date_dt"] = pd.to_datetime(df_beton[date_col], errors="coerce")
    else:
        df_beton["date_dt"] = pd.NaT

    # Calcul des résistances moyennes par échéance
    if not df_ecrasement.empty and ref_col and "ref_controle" in df_ecrasement.columns and "resistance" in df_ecrasement.columns and "echeance" in df_ecrasement.columns:
        df_ecrasement["resistance"] = pd.to_numeric(df_ecrasement["resistance"], errors="coerce")
        
        res_3j = df_ecrasement[df_ecrasement["echeance"].astype(str).str.lower().isin(["3j", "3 j", "3d", "3 jours"])].groupby("ref_controle")["resistance"].mean().rename("res_3j")
        res_7j = df_ecrasement[df_ecrasement["echeance"].astype(str).str.lower().isin(["7j", "7 j", "7d", "7 jours"])].groupby("ref_controle")["resistance"].mean().rename("res_7j")
        res_28j = df_ecrasement[df_ecrasement["echeance"].astype(str).str.lower().isin(["28j", "28 j", "28d", "28 jours"])].groupby("ref_controle")["resistance"].mean().rename("res_28j")

        df_merged = df_beton.merge(res_3j, left_on=ref_col, right_index=True, how="left")
        df_merged = df_merged.merge(res_7j, left_on=ref_col, right_index=True, how="left")
        df_merged = df_merged.merge(res_28j, left_on=ref_col, right_index=True, how="left")
    else:
        df_merged = df_beton
        df_merged["res_3j"] = None
        df_merged["res_7j"] = None
        df_merged["res_28j"] = None

    df_merged["_ref_col"] = df_merged[ref_col] if ref_col else "-"
    return df_merged


def format_final_dataframe(df_filtered):
    """Formate les 7 colonnes requises pour l'affichage et l'exportation."""
    df_display = pd.DataFrame()
    df_display["1. Référence de Contrôle"] = df_filtered["_ref_col"].fillna("-")
    df_display["2. Date de Prélèvement"] = df_filtered["date_dt"].dt.strftime("%d/%m/%Y").fillna("-")
    df_display["3. Affaissement (cm)"] = df_filtered.get("affaissement", "-").fillna("-")
    df_display["4. Température (°C)"] = df_filtered.get("temperature", "-").fillna("-")
    
    df_display["5. Résistance moyenne 3J (MPa)"] = df_filtered["res_3j"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "-")
    df_display["6. Résistance moyenne 7J (MPa)"] = df_filtered["res_7j"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "-")
    df_display["7. Résistance moyenne 28J (MPa)"] = df_filtered["res_28j"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "-")
    
    return df_display


# =========================================================
# VUE PRINCIPALE STREAMLIT
# =========================================================
def show(supabase):
    st.title("📊 Récapitulatif et Synthèse du Contrôle Béton")
    st.caption("Synthèse consolidée avec filtres journaliers et mensuels par classe de béton")

    if supabase is None:
        st.error("❌ Connexion Supabase indisponible.")
        st.stop()

    tab_journalier, tab_mensuel = st.tabs(["📅 Bilan Journalier", "📅 Bilan Mensuel"])

    try:
        df_merged = load_and_process_data(supabase)
    except Exception as e:
        st.error(f"❌ Erreur lors de la récupération des données : {e}")
        st.stop()

    if df_merged.empty:
        st.warning("⚠️ Aucune donnée disponible dans le système.")
        st.stop()

    # Liste des classes de béton pour les filtres
    classes_dispo = ["Toutes"]
    if "classe_beton" in df_merged.columns:
        classes_dispo += sorted(list(df_merged["classe_beton"].dropna().unique()))

    # ---------------------------------------------------------
    # 1. BILAN JOURNALIER
    # ---------------------------------------------------------
    with tab_journalier:
        st.markdown("### Filtrage journalier par classe de béton")
        
        col1, col2 = st.columns(2)
        with col1:
            selected_date = st.date_input("Sélectionnez une date :", value=date.today(), key="j_date")
        with col2:
            selected_class_j = st.selectbox("Filtrer par classe de béton :", classes_dispo, key="j_class")

        df_j = df_merged[df_merged["date_dt"].dt.date == selected_date]
        if selected_class_j != "Toutes" and "classe_beton" in df_j.columns:
            df_j = df_j[df_j["classe_beton"] == selected_class_j]

        if df_j.empty:
            st.info("ℹ️ Aucun contrôle enregistré pour les critères sélectionnés.")
        else:
            df_display_j = format_final_dataframe(df_j)

            st.markdown("---")
            k1, k2 = st.columns(2)
            k1.metric("Nombre de Contrôles", f"{len(df_display_j)}")
            
            # Affaissement moyen si numérique
            aff_num = pd.to_numeric(df_j.get("affaissement", pd.Series()), errors="coerce")
            if not aff_num.dropna().empty:
                k2.metric("Affaissement Moyen", f"{aff_num.mean():.1f} cm")
            else:
                k2.metric("Affaissement Moyen", "-")

            st.markdown("---")

            excel_file_j = generate_excel_synthesis(df_display_j, f"Journée du {selected_date.strftime('%d/%m/%Y')}")
            st.download_button(
                label="📥 Télécharger la Synthèse Journalière Excel (A4 Portrait)",
                data=excel_file_j,
                file_name=f"Synthese_Controle_Beton_{selected_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.dataframe(df_display_j, use_container_width=True)

    # ---------------------------------------------------------
    # 2. BILAN MENSUEL
    # ---------------------------------------------------------
    with tab_mensuel:
        st.markdown("### Bilan mensuel global par classe de béton")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            annee = date.today().year
            mois_liste = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
            mois_selected = st.selectbox("Sélectionnez le mois :", mois_liste, index=date.today().month - 1, key="m_mois")
            mois_num = mois_liste.index(mois_selected) + 1
        
        with col_m2:
            selected_class_m = st.selectbox("Filtrer par classe de béton :", classes_dispo, key="m_class")

        df_m = df_merged[(df_merged["date_dt"].dt.year == annee) & (df_merged["date_dt"].dt.month == mois_num)]
        if selected_class_m != "Toutes" and "classe_beton" in df_m.columns:
            df_m = df_m[df_m["classe_beton"] == selected_class_m]

        if df_m.empty:
            st.info("ℹ️ Aucun contrôle enregistré pour ce mois.")
        else:
            df_display_m = format_final_dataframe(df_m)

            st.markdown("---")
            st.metric("Total Prélèvements / Contrôles du Mois", f"{len(df_display_m)}")
            st.markdown("---")

            excel_file_m = generate_excel_synthesis(df_display_m, f"Mois de {mois_selected} {annee}")
            st.download_button(
                label="📥 Télécharger la Synthèse Mensuelle Excel (A4 Portrait)",
                data=excel_file_m,
                file_name=f"Synthese_Mensuelle_Beton_{mois_selected}_{annee}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.dataframe(df_display_m, use_container_width=True)
