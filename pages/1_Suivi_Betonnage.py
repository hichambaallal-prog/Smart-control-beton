import streamlit as st
import pandas as pd
from datetime import datetime, date
from supabase import create_client, Client
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. Configuration de la page
st.set_page_config(page_title="Suivi Béton - LGV Casa Sud (LPEE)", layout="wide")

# =========================================================
# FONCTION 1 : EXCEL RAPPORT JOURNALIER PAR CLASSE DE BÉTON & HISTORIQUE (LPEE)
# =========================================================
def generer_excel_lpee(df_jour, date_rapport="", est_historique=False):
    """
    Génère un fichier Excel (.xlsx) stylisé aux normes LPEE.
    Chaque classe/type de béton possède sa propre feuille (onglet) séparée dans le fichier journalier.
    Incorpore l'en-tête officiel LPEE CTR-CSB et le pied de page avec cases de signatures.
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Supprimer la feuille par défaut

    # Définition des Styles LPEE
    font_titre = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    font_sub_ctr = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    font_info = Font(name="Calibri", size=10, italic=True, color="333333")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_sig_title = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_conforme = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_conforme = Font(name="Calibri", size=10, bold=True, color="006100")
    
    fill_non_conforme = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_non_conforme = Font(name="Calibri", size=10, bold=True, color="9C0006")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # Identification des classes de béton présentes
    if df_jour.empty:
        classes = ["Rapport Béton"]
    elif est_historique or "classe_beton" not in df_jour.columns:
        classes = ["Registre Général"]
    else:
        classes = df_jour["classe_beton"].unique().tolist()
        if not classes:
            classes = ["Général"]

    for cls in classes:
        # Nom de l'onglet Excel (max 30 caractères sans caractères interdits)
        sheet_name = str(cls).replace("/", "-").replace("\\", "-")[:30]
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].showGridLines = True

        if est_historique or cls in ["Registre Général", "Rapport Béton"]:
            df_cls = df_jour.copy()
        else:
            df_cls = df_jour[df_jour["classe_beton"] == cls].copy()

        # --- 1. EN-TÊTE OFFICIEL LPEE & PROJET ---
        ws['A1'] = "LPEE - LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES"
        ws['A1'].font = font_titre
        ws['A1'].alignment = align_left

        ws['A2'] = "CENTRE TECHNIQUE RÉGIONAL DE CASABLANCA-SETTAT-BÉNI MELLAL (CTR-CSB)"
        ws['A2'].font = font_sub_ctr
        ws['A2'].alignment = align_left

        if est_historique:
            ws['A3'] = "Projet : LGV CASA SUD | Client : TGCC | Registre Général et Historique Complet"
        else:
            ws['A3'] = f"Projet : LGV CASA SUD | Client : TGCC | Classe de Béton : {cls} | Date : {date_rapport}"
        ws['A3'].font = font_info
        ws['A3'].alignment = align_left

        # Mapping des noms de colonnes
        col_mapping = {
            "N°": "N°",
            "num_bon_livraison": "N° Bon Livraison",
            "ouvrage": "Ouvrage",
            "element_betonne": "Élément Bétonné",
            "volume_beton": "Volume (m³)",
            "classe_beton": "Classe Béton",
            "heure_fin_production_cab": "Fin Prod. CAB",
            "heure_arrivee_chantier": "Arrivée Chantier",
            "tbf": "TBF (°C)",
            "ta": "TA (°C)",
            "affaissement": "Slump (mm)",
            "prelevement": "Prélèvement",
            "technicien": "Technicien",
            "statut": "STATUT"
        }

        df_export = df_cls.copy()
        df_export.insert(0, 'N°', range(1, len(df_export) + 1))
        cols_presentes = [c for c in list(col_mapping.keys()) if c in df_export.columns]

        start_row = 5
        # En-têtes (Ligne 5)
        ws.row_dimensions[start_row].height = 24
        for col_idx, col_key in enumerate(cols_presentes, start=1):
            cell = ws.cell(row=start_row, column=col_idx)
            cell.value = col_mapping[col_key]
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        # Remplissage des données
        current_row = start_row + 1
        for row_data in df_export[cols_presentes].values:
            ws.row_dimensions[current_row].height = 20
            for col_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = val
                cell.border = thin_border
                cell.alignment = align_center

                col_key = cols_presentes[col_idx - 1]
                if col_key == "statut":
                    val_str = str(val)
                    if "Conforme" in val_str and "Non" not in val_str:
                        cell.fill = fill_conforme
                        cell.font = font_conforme
                    else:
                        cell.fill = fill_non_conforme
                        cell.font = font_non_conforme
            current_row += 1

        # --- 2. PIED DE PAGE : CASES DE SIGNATURES (Responsable d'essai / Chef du laboratoire) ---
        sig_row = current_row + 2
        ws.row_dimensions[sig_row].height = 20

        # Case 1 : Responsable d'essai (Colonnes B à D)
        ws.merge_cells(start_row=sig_row, start_column=2, end_row=sig_row, end_column=4)
        c_resp = ws.cell(row=sig_row, column=2, value="Responsable d'essai")
        c_resp.font = font_sig_title
        c_resp.alignment = align_center

        for r in range(sig_row, sig_row + 4):
            for c in range(2, 5):
                cell = ws.cell(row=r, column=c)
                top = Side(style='thin', color='1F4E78') if r == sig_row else None
                bottom = Side(style='thin', color='1F4E78') if r == sig_row + 3 else None
                left = Side(style='thin', color='1F4E78') if c == 2 else None
                right = Side(style='thin', color='1F4E78') if c == 4 else None
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

        # Case 2 : Chef du laboratoire (Colonnes F à H ou fin du tableau)
        col_chef_start = min(8, max(5, len(cols_presentes) - 2))
        col_chef_end = col_chef_start + 2

        ws.merge_cells(start_row=sig_row, start_column=col_chef_start, end_row=sig_row, end_column=col_chef_end)
        c_chef = ws.cell(row=sig_row, column=col_chef_start, value="Chef du laboratoire")
        c_chef.font = font_sig_title
        c_chef.alignment = align_center

        for r in range(sig_row, sig_row + 4):
            for c in range(col_chef_start, col_chef_end + 1):
                cell = ws.cell(row=r, column=c)
                top = Side(style='thin', color='1F4E78') if r == sig_row else None
                bottom = Side(style='thin', color='1F4E78') if r == sig_row + 3 else None
                left = Side(style='thin', color='1F4E78') if c == col_chef_start else None
                right = Side(style='thin', color='1F4E78') if c == col_chef_end else None
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

        # Ajustement des largeurs de colonnes
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row < start_row:
                    continue
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(output)
    return output.getvalue()

# =========================================================
# FONCTION 2 : EXCEL RÉCAPITULATIF MENSUEL STYLISÉ
# =========================================================
def generer_excel_recap_mensuel_lpee(df_recap, mois):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Synthèse {mois.replace('/', '-')}"
    ws.views.sheetView[0].showGridLines = True

    font_titre = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    font_sub_ctr = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    font_card_label = Font(name="Calibri", size=10, bold=True, color="595959")
    font_card_val = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_total = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_card = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    total_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='medium', color='1F4E78'), bottom=Side(style='double', color='1F4E78')
    )
    
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # En-tête LPEE
    ws['A1'] = "LPEE - LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES"
    ws['A1'].font = font_titre

    ws['A2'] = "CENTRE TECHNIQUE RÉGIONAL DE CASABLANCA-SETTAT-BÉNI MELLAL (CTR-CSB)"
    ws['A2'].font = font_sub_ctr

    # Bloc Infos
    ws['A4'] = "PROJET :"
    ws['A4'].font = font_card_label
    ws['B4'] = "LGV CASA SUD"
    ws['B4'].font = font_card_val

    ws['D4'] = "CLIENT / ENTREPRISE :"
    ws['D4'].font = font_card_label
    ws['E4'] = "TGCC"
    ws['E4'].font = font_card_val

    ws['A5'] = "ORGANISME DE CONTRÔLE :"
    ws['A5'].font = font_card_label
    ws['B5'] = "LPEE - Laboratoire LGV Casa Sud"
    ws['B5'].font = font_card_val

    ws['D5'] = "PÉRIODE DE SYNTHÈSE :"
    ws['D5'].font = font_card_label
    ws['E5'] = mois
    ws['E5'].font = font_card_val

    for r in range(4, 6):
        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_card
            cell.border = thin_border

    # Tableau
    headers = list(df_recap.columns)
    start_row = 7
    ws.row_dimensions[start_row].height = 26

    for col_idx, h_name in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx)
        cell.value = h_name
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    current_row = start_row + 1
    for row_data in df_recap.values:
        ws.row_dimensions[current_row].height = 21
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = val
            cell.border = thin_border
            cell.alignment = align_center if col_idx in [1, 3, 4, 5, 6, 7] else align_left
        current_row += 1

    # Totaux
    if len(df_recap) > 0:
        ws.row_dimensions[current_row].height = 24
        c_tot = ws.cell(row=current_row, column=1, value="TOTAL / SYNTHÈSE MOIS")
        c_tot.font = font_total; c_tot.alignment = align_center; c_tot.fill = fill_total; c_tot.border = total_border
        
        ws.cell(row=current_row, column=2, value="—").fill = fill_total
        
        tot_ctrl = int(df_recap["Nombre de contrôles"].sum())
        c_ctrl = ws.cell(row=current_row, column=3, value=tot_ctrl)
        c_ctrl.font = font_total; c_ctrl.alignment = align_center; c_ctrl.fill = fill_total; c_ctrl.border = total_border

        for c_idx, val in enumerate([
            df_recap["Affaissement Min (mm)"].min(), df_recap["Affaissement Max (mm)"].max(),
            df_recap["TBF Min (°C)"].min(), df_recap["TBF Max (°C)"].max()
        ], start=4):
            c_cell = ws.cell(row=current_row, column=c_idx, value=val)
            c_cell.font = font_total; c_cell.alignment = align_center; c_cell.fill = fill_total; c_cell.border = total_border

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < start_row: continue
            if cell.value is not None: max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

    wb.save(output)
    return output.getvalue()

# =========================================================
# CALCUL STATUT BÉTON
# =========================================================
def evaluer_statut_beton(tbf, h_fin, h_arr, affaissement):
    try:
        cond_tbf = float(tbf) < 32.1
        cond_aff = 160 <= int(affaissement) <= 220
        fmt = "%H:%M"
        t_fin = datetime.strptime(str(h_fin).strip(), fmt)
        t_arr = datetime.strptime(str(h_arr).strip(), fmt)
        diff_minutes = (t_arr - t_fin).total_seconds() / 60
        if diff_minutes < 0: diff_minutes += 24 * 60
        cond_delai = 0 <= diff_minutes <= 120
        return "✅ Conforme" if (cond_tbf and cond_aff and cond_delai) else "⚠️ Non Conforme"
    except Exception:
        return "⚠️ Non Conforme"

# =========================================================
# 2. VALEURS FIXES & AUTHENTIFICATION
# =========================================================
DEFAULT_PROJET = "LGV Casa Sud"
DEFAULT_ENTREPRISE = "TGCC"
DEFAULT_CENTRALE = "TG PREFA"

PASSWORD_GENERAL = "lpee2026"
PASSWORD_ADMIN = "lpee_admin_2026"

if "authentifie" not in st.session_state: st.session_state["authentifie"] = False
if "is_admin" not in st.session_state: st.session_state["is_admin"] = False

if not st.session_state["authentifie"]:
    st.title("🔒 Accès Sécurisé - LPEE Smart Control Béton")
    pwd = st.text_input("Mot de passe", type="password")
    if st.button("Valider l'accès"):
        if pwd in [PASSWORD_GENERAL, PASSWORD_ADMIN]:
            st.session_state["authentifie"] = True
            if pwd == PASSWORD_ADMIN: st.session_state["is_admin"] = True
            st.rerun()
        else:
            st.error("❌ Mot de passe incorrect.")
    st.stop()

# =========================================================
# 3. SUPABASE
# =========================================================
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1"

try:
    supabase: Client = create_client(URL, CLE)
    response = supabase.table("controles_beton").select("*").execute()
    data_all = response.data or []
except Exception as e:
    st.error(f"Erreur Supabase : {e}")
    data_all = []

st.title("🚧 Suivi Journalier de Bétonnage - LGV Casa Sud")
st.markdown("##### LPEE - CTR-CSB | Projet : **LGV CASA SUD** | Client : **TGCC**")

# =========================================================
# 4. SÉLECTION JOURNÉE
# =========================================================
st.markdown("---")
col_date1, col_date2 = st.columns([1, 2])
with col_date1:
    date_choisie = st.date_input("📅 Choisir la journée de bétonnage :", value=date.today())
    str_date_choisie = date_choisie.strftime("%d/%m/%Y")

data_jour = [r for r in data_all if r.get("date_betonnage") == str_date_choisie]

with col_date2:
    st.info(f"📌 **Journée du {str_date_choisie}** | Total camions contrôlés : **{len(data_jour)}**")

# =========================================================
# 5. ONGLETS PRINCIPAUX
# =========================================================
tab_ajouter, tab_modifier, tab_supprimer, tab_historique, tab_recap_mensuel = st.tabs([
    "➕ Ajouter un camion", "✏️ Modifier (Admin)", "❌ Supprimer (Admin)", "📚 Historique Général", "📊 Récap Mensuel"
])

# --- ONGLET 1 : SAISIE ---
with tab_ajouter:
    st.subheader("📝 Saisie de suivi de bétonnage")
    with st.form("form_controle_ajouter"):
        c1, c2, c3 = st.columns(3)
        with c1:
            date_betonnage_saisie = st.date_input("📅 Date de bétonnage :", value=date_choisie, key="add_date_saisie")
            projet = st.text_input("Projet", value=DEFAULT_PROJET, disabled=True)
            entreprise = st.text_input("Entreprise / Client", value=DEFAULT_ENTREPRISE, disabled=True)
            centrale_beton = st.text_input("Centrale béton", value=DEFAULT_CENTRALE, disabled=True)
        with c2:
            ouvrage = st.text_input("Ouvrage", value="PRO745 OA1", key="add_ouvrage")
            element_betonne = st.text_input("Élément bétonné", value="Semelle C0", key="add_elem")
            volume_beton = st.text_input("Volume béton (m³)", value="8", key="add_vol")
        with c3:
            num_bon_livraison = st.text_input("N° bon livraison (BL)", value="BL2548", key="add_bl")
            classe_beton = st.text_input("Classe béton (ex: C30/37)", value="C30/37", key="add_classe")

        st.markdown("---")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            heure_fin_prod = st.text_input("Heure fin prod. CAB (HH:MM)", value="15:28", key="add_h_fin")
            heure_arrivee = st.text_input("Heure arrivée chantier (HH:MM)", value="16:31", key="add_h_arr")
        with col_m2:
            tbf = st.number_input("TBF (°C)", value=32.0, step=0.1, format="%.1f", key="add_tbf")
            ta = st.number_input("TA (°C)", value=28.9, step=0.1, format="%.1f", key="add_ta")
        with col_m3:
            affaissement = st.number_input("Affaissement (mm)", value=170, step=1, format="%d", key="add_aff")
            meteo = st.selectbox("Météo", ["Soleil", "Nuageux", "Pluie", "Vent"], key="add_meteo")
        with col_m4:
            prelevement = st.selectbox("Prélèvement", ["OUI", "NON"], key="add_prelev")
            technicien = st.text_input("Technicien Contrôleur", value="Ismail / Mohamed", key="add_tech")

        observations = st.text_area("Observations", value="RAS", key="add_obs")
        submit_add = st.form_submit_button("💾 Enregistrer le camion")

        if submit_add:
            statut_auto = evaluer_statut_beton(tbf, heure_fin_prod, heure_arrivee, affaissement)
            str_date_saisie = date_betonnage_saisie.strftime("%d/%m/%Y")
            data_to_insert = {
                "projet": DEFAULT_PROJET, "entreprise": DEFAULT_ENTREPRISE, "centrale_beton": DEFAULT_CENTRALE,
                "ouvrage": ouvrage, "element_betonne": element_betonne, "volume_beton": volume_beton,
                "num_bon_livraison": num_bon_livraison, "classe_beton": classe_beton,
                "date_betonnage": str_date_saisie, "meteo": meteo, "observations": observations,
                "heure_fin_production_cab": heure_fin_prod, "heure_arrivee_chantier": heure_arrivee,
                "tbf": round(tbf, 1), "ta": round(ta, 1), "affaissement": int(affaissement),
                "prelevement": prelevement, "technicien": technicien, "statut": statut_auto
            }
            try:
                supabase.table("controles_beton").insert(data_to_insert).execute()
                st.success(f"Camion BL : {num_bon_livraison} (Classe : {classe_beton}) enregistré pour le {str_date_saisie} !")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur d'enregistrement : {e}")

# --- ONGLET 2 : MODIFIER ---
with tab_modifier:
    if not st.session_state["is_admin"]:
        pwd_admin = st.text_input("Code Administrateur :", type="password", key="mod_pwd")
        if st.button("Débloquer"):
            if pwd_admin == PASSWORD_ADMIN:
                st.session_state["is_admin"] = True; st.rerun()
            else: st.error("❌ Code incorrect")
    else:
        if data_jour:
            options_edit = {f"BL: {r.get('num_bon_livraison')} | Classe: {r.get('classe_beton')} | Ouvrage: {r.get('ouvrage')}": r for r in data_jour}
            choix = st.selectbox("Sélectionner le camion :", list(options_edit.keys()))
            row_s = options_edit[choix]
            with st.form("form_edit"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    e_proj = st.text_input("Projet", value=str(row_s.get('projet') or DEFAULT_PROJET))
                    e_ent = st.text_input("Client", value=str(row_s.get('entreprise') or DEFAULT_ENTREPRISE))
                    e_cent = st.text_input("Centrale", value=str(row_s.get('centrale_beton') or DEFAULT_CENTRALE))
                with c2:
                    e_ouv = st.text_input("Ouvrage", value=str(row_s.get('ouvrage') or ''))
                    e_elem = st.text_input("Élément", value=str(row_s.get('element_betonne') or ''))
                    e_vol = st.text_input("Volume", value=str(row_s.get('volume_beton') or ''))
                with c3:
                    e_bl = st.text_input("BL", value=str(row_s.get('num_bon_livraison') or ''))
                    e_classe = st.text_input("Classe béton", value=str(row_s.get('classe_beton') or ''))

                st.markdown("---")
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                with col_m1:
                    e_h_fin = st.text_input("Fin Prod CAB", value=str(row_s.get('heure_fin_production_cab') or ''))
                    e_h_arr = st.text_input("Arrivée Chantier", value=str(row_s.get('heure_arrivee_chantier') or ''))
                with col_m2:
                    e_tbf = st.number_input("TBF", value=float(row_s.get('tbf') or 0.0), step=0.1)
                    e_ta = st.number_input("TA", value=float(row_s.get('ta') or 0.0), step=0.1)
                with col_m3:
                    e_aff = st.number_input("Affaissement", value=int(row_s.get('affaissement') or 0), step=1)
                    e_meteo = st.selectbox("Météo", ["Soleil", "Nuageux", "Pluie", "Vent"], index=0)
                with col_m4:
                    e_prelev = st.selectbox("Prélèvement", ["OUI", "NON"], index=0)
                    e_tech = st.text_input("Technicien", value=str(row_s.get('technicien') or ''))

                e_obs = st.text_area("Observations", value=str(row_s.get('observations') or ''))
                if st.form_submit_button("🔄 Mettre à jour"):
                    statut_up = evaluer_statut_beton(e_tbf, e_h_fin, e_h_arr, e_aff)
                    up_data = {
                        "projet": e_proj, "entreprise": e_ent, "centrale_beton": e_cent,
                        "ouvrage": e_ouv, "element_betonne": e_elem, "volume_beton": e_vol,
                        "num_bon_livraison": e_bl, "classe_beton": e_classe, "date_betonnage": str_date_choisie,
                        "meteo": e_meteo, "observations": e_obs, "heure_fin_production_cab": e_h_fin,
                        "heure_arrivee_chantier": e_h_arr, "tbf": round(e_tbf, 1), "ta": round(e_ta, 1),
                        "affaissement": int(e_aff), "prelevement": e_prelev, "technicien": e_tech, "statut": statut_up
                    }
                    supabase.table("controles_beton").update(up_data).eq("id", row_s["id"]).execute()
                    st.success("Données mises à jour !"); st.rerun()

# --- ONGLET 3 : SUPPRIMER ---
with tab_supprimer:
    if st.session_state["is_admin"] and data_jour:
        options_del = {f"BL: {r.get('num_bon_livraison')} | Classe: {r.get('classe_beton')}": r["id"] for r in data_jour}
        c_del = st.selectbox("Supprimer la ligne :", list(options_del.keys()))
        if st.button("🚨 Supprimer définitivement", type="primary"):
            supabase.table("controles_beton").delete().eq("id", options_del[c_del]).execute()
            st.success("Ligne supprimée !"); st.rerun()

# --- ONGLET 4 : HISTORIQUE GENERAL ---
with tab_historique:
    st.subheader("📚 Registre Général des Contrôles")
    if data_all:
        df_all = pd.DataFrame(data_all)
        cols_h = ["num_bon_livraison", "ouvrage", "element_betonne", "volume_beton", "classe_beton", "date_betonnage", "tbf", "affaissement", "statut"]
        df_h_v = df_all[[c for c in cols_h if c in df_all.columns]]
        st.dataframe(df_h_v, use_container_width=True)

# --- ONGLET 5 : RÉCAP MENSUEL ---
with tab_recap_mensuel:
    st.subheader("📊 Récapitulatif Mensuel LPEE - CTR-CSB")
    if data_all:
        df_r = pd.DataFrame(data_all)
        df_r["dt"] = pd.to_datetime(df_r["date_betonnage"], format="%d/%m/%Y", errors="coerce")
        df_v = df_r.dropna(subset=["dt"]).copy()
        if not df_v.empty:
            df_v["mois_annee"] = df_v["dt"].dt.strftime("%m/%Y")
            mois_sel = st.selectbox("📅 Choisir le mois :", sorted(df_v["mois_annee"].unique(), reverse=True))
            df_m = df_v[df_v["mois_annee"] == mois_sel].copy()
            df_m["ouvrage_elem"] = df_m["ouvrage"].fillna('') + " - " + df_m["element_betonne"].fillna('')
            
            df_recap = df_m.groupby(["dt", "date_betonnage", "ouvrage_elem"]).agg(
                nb_controles=("ouvrage_elem", "count"),
                aff_min=("affaissement", "min"), aff_max=("affaissement", "max"),
                tbf_min=("tbf", "min"), tbf_max=("tbf", "max")
            ).reset_index().sort_values(["dt"])

            df_f = df_recap[["date_betonnage", "ouvrage_elem", "nb_controles", "aff_min", "aff_max", "tbf_min", "tbf_max"]].rename(columns={
                "date_betonnage": "Date", "ouvrage_elem": "Ouvrage / Élément Bétonné",
                "nb_controles": "Nombre de contrôles", "aff_min": "Affaissement Min (mm)",
                "aff_max": "Affaissement Max (mm)", "tbf_min": "TBF Min (°C)", "tbf_max": "TBF Max (°C)"
            })
            st.dataframe(df_f, use_container_width=True)
            
            excel_m = generer_excel_recap_mensuel_lpee(df_f, mois_sel)
            st.download_button("📊 Télécharger le Récap Mensuel Officiel (.xlsx)", data=excel_m, file_name=f"Recap_Mensuel_LPEE_{mois_sel.replace('/', '_')}.xlsx")

# =========================================================
# 6. TABLEAUX ET FICHIER EXCEL SÉPARÉS PAR TYPE DE BÉTON
# =========================================================
st.markdown("---")
st.subheader(f"📋 Rapport Journalier de Bétonnage — Date : {str_date_choisie}")

if data_jour:
    df_jour = pd.DataFrame(data_jour)
    
    # Obtenir la liste des types de béton de la journée
    classes_du_jour = df_jour["classe_beton"].unique().tolist()
    
    # Métriques globales du jour
    vol_tot_j = pd.to_numeric(df_jour["volume_beton"], errors="coerce").fillna(0).sum()
    m1, m2, m3 = st.columns(3)
    with m1: st.metric("🏗️ Volume Total du Jour", f"{vol_tot_j:.2f} m³")
    with m2: st.metric("🚛 Total Camions (Jour)", f"{len(df_jour)}")
    with m3: st.metric("🏷️ Types de Béton Livrés", f"{len(classes_du_jour)} type(s)")

    st.markdown("---")
    
    # Sélecteur de type de béton pour l'affichage à l'écran
    cls_selectionnee = st.selectbox(
        "🔎 Filtrer l'affichage par type/classe de béton :", 
        ["TOUS LES TYPES"] + classes_du_jour, 
        key="select_classe_view"
    )

    if cls_selectionnee == "TOUS LES TYPES":
        df_visu = df_jour.copy()
    else:
        df_visu = df_jour[df_jour["classe_beton"] == cls_selectionnee].copy()

    cols_vis = ["num_bon_livraison", "ouvrage", "element_betonne", "volume_beton", "classe_beton", "heure_fin_production_cab", "heure_arrivee_chantier", "tbf", "ta", "affaissement", "prelevement", "technicien", "statut"]
    df_visu_display = df_visu[[c for c in cols_vis if c in df_visu.columns]]
    df_visu_display.index = range(1, len(df_visu_display) + 1)
    
    st.dataframe(df_visu_display, use_container_width=True)

    # Téléchargement du fichier Excel officiel
    # Remarque : Le fichier Excel téléchargé contiendra 1 ONGLET SÉPARÉ par type de béton !
    excel_jour_bytes = generer_excel_lpee(df_jour, date_rapport=str_date_choisie)
    nom_excel_jour = f"Rapport_Betonnage_LPEE_{date_choisie.strftime('%Y-%m-%d')}.xlsx"

    st.download_button(
        label=f"📥 Télécharger le Rapport Journalier Officiel LPEE (.xlsx) — Feuille séparée par type de béton",
        data=excel_jour_bytes,
        file_name=nom_excel_jour,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.warning(f"Aucun camion enregistré pour le {str_date_choisie}.")
