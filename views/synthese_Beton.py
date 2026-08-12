import streamlit as st
import pandas as pd
from datetime import datetime, date
import io

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# FONCTION GENERATION EXCEL FORMAT A4 PORTRAIT (ESPACÉ & PURGÉ)
# =========================================================
def generate_excel_synthesis(df_data, titre_periode):
    """Génère un fichier Excel en A4 Portrait sans les colonnes/KPIs Technicien, Observations et Éprouvettes."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Synthèse Béton"

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

    # --- 2. PALETTE DE COULEURS ET POLICES (CALIBRI 12) ---
    color_primary = "1F4E79"    # Bleu LPEE / Marine
    color_header = "2D572C"     # Vert/Gris foncé entête
    color_card_bg = "F7F9FA"    # Fond clair cartes info
    color_kpi_bg = "EDF2F8"     # Fond KPI

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

    # Calcul des dimensions
    nb_cols = max(len(df_data.columns), 6)
    last_col_letter = get_column_letter(nb_cols)
    mid_col_idx = nb_cols // 2
    mid_col_letter = get_column_letter(mid_col_idx)
    next_mid_letter = get_column_letter(mid_col_idx + 1)

    # --- 3. BANNIÈRE EN-TÊTE LPEE ---
    ws.merge_cells(f"A1:{last_col_letter}2")
    cell_title = ws["A1"]
    cell_title.value = "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - CTR-CSB\nRAPPORT DE SYNTHÈSE DU BÉTONNAGE"
    cell_title.font = font_title
    cell_title.fill = fill_title
    cell_title.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 28

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
        ws.row_dimensions[r].height = 32
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    # --- 5. RÉSUMÉ DU VOLUME TOTAL UNIQUEMENT ---
    row_idx = 7
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    ws[f"A{row_idx}"] = "📊 RÉSUMÉ GLOBAL"
    ws[f"A{row_idx}"].font = font_section
    ws.row_dimensions[row_idx].height = 30

    row_idx += 1
    vol_tot = df_data["Quantité (m³)"].sum() if "Quantité (m³)" in df_data.columns else 0

    # Titre KPI (Uniquement Volume Total)
    ws.merge_cells(f"A{row_idx}:{last_col_letter}{row_idx}")
    cell_k_lbl = ws[f"A{row_idx}"]
    cell_k_lbl.value = "Volume Total Béton"
    cell_k_lbl.font = font_bold
    cell_k_lbl.fill = fill_kpi
    cell_k_lbl.alignment = Alignment(horizontal="center", vertical="center")

    # Valeur KPI
    ws.merge_cells(f"A{row_idx+1}:{last_col_letter}{row_idx+1}")
    cell_k_val = ws[f"A{row_idx+1}"]
    cell_k_val.value = f"{vol_tot:.1f} m³"
    cell_k_val.font = font_kpi_val
    cell_k_val.fill = fill_kpi
    cell_k_val.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(row_idx, row_idx + 2):
        for c in range(1, nb_cols + 1):
            ws.cell(row=r, column=c).border = thin_border

    ws.row_dimensions[row_idx].height = 28
    ws.row_dimensions[row_idx+1].height = 36
    row_idx += 3

    # --- 6. TABLEAU DES DONNÉES ---
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
            cell.font = font_normal  # Calibri 12
            cell.border = thin_border
            if isinstance(val, (int, float)):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        ws.row_dimensions[row_idx].height = 36
        row_idx += 1

    # Ligne de TOTAL
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

    # --- 7. PIED DE PAGE : ZONE DE SIGNATURES ---
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

    # --- 8. LARGEUR SUR MESURE DES COLONNES RESTANTES ---
    col_width_map = {
        "Date Livraison": 16,
        "Heure d'arrivée": 15,
        "N° BL": 16,
        "Ouvrage": 22,
        "Quantité (m³)": 16,
        "Classe": 14,
        "Durée de transport": 18,
        "Temp. Béton": 15,
        "Temp. Ambiante": 16,
        "Affaissement": 15,
        "Prélèvement": 18,
        "Météo": 15
    }

    for col_idx, col_name in enumerate(headers, 1):
        col_letter = get_column_letter(col_idx)
        width = col_width_map.get(col_name, 16)
        ws.column_dimensions[col_letter].width = width

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# VUE PRINCIPALE STREAMLIT
# =========================================================
def show(supabase):
    st.title("📊 Récapitulatif et Synthèse du Bétonnage")
    
    tab_journalier, tab_mensuel = st.tabs(["📅 Bilan Journalier", "📅 Bilan Mensuel"])
    
    # ---------------------------------------------------------
    # 1. BILAN JOURNALIER
    # ---------------------------------------------------------
    with tab_journalier:
        st.markdown("### Filtrage par jour et par classe de béton")
        
        col1, col2 = st.columns(2)
        with col1:
            selected_date = st.date_input("Sélectionnez une date :", value=date.today())
        with col2:
            selected_class = st.selectbox(
                "Filtrer par classe de béton :", 
                ["Toutes", "C25/30", "C30/37", "C35/45", "C40/50", "C45/55"]
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
                    
                    excel_file = generate_excel_synthesis(df_display, f"Journée du {selected_date.strftime('%d/%m/%Y')}")
                    st.download_button(
                        label="📥 Télécharger la Synthèse Excel (Format A4 Portrait)",
                        data=excel_file,
                        file_name=f"Synthese_Beton_{selected_date}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.dataframe(df_display, use_container_width=True)
            else:
                st.info("Aucun coulage enregistré pour les critères sélectionnés.")
                
        except Exception as e:
            st.error(f"Erreur de chargement : {e}")

    # ---------------------------------------------------------
    # 2. BILAN MENSUEL
    # ---------------------------------------------------------
    with tab_mensuel:
        st.markdown("### Bilan mensuel global")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            annee = date.today().year
            mois_liste = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
            mois_selected = st.selectbox("Sélectionnez le mois :", mois_liste, index=date.today().month - 1)
            mois_num = mois_liste.index(mois_selected) + 1
            
        with col_m2:
            selected_class_m = st.selectbox(
                "Filtrer par classe de béton (Mensuel) :", 
                ["Toutes", "C25/30", "C30/37", "C35/45", "C40/50", "C45/55"],
                key="class_mensuel"
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
                    
                    excel_file_m = generate_excel_synthesis(df_m_display, f"Mois de {mois_selected} {annee}")
                    st.download_button(
                        label="📥 Télécharger la Synthèse Mensuelle Excel (Format A4 Portrait)",
                        data=excel_file_m,
                        file_name=f"Synthese_Mensuelle_{mois_selected}_{annee}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.dataframe(df_m_display, use_container_width=True)
            else:
                st.info("Aucun coulage enregistré pour ce mois.")
                
        except Exception as e:
            st.error(f"Erreur de chargement : {e}")
