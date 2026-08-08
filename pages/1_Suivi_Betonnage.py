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
# FONCTION 1 : EXCEL DASHBOARD JOURNALIER & HISTORIQUE (LPEE)
# =========================================================
def generer_excel_lpee(df, date_rapport="", est_historique=False):
    """
    Génère un fichier Excel (.xlsx) stylisé aux normes LPEE pour le journalier / historique.
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rapport Bétonnage"

    # Afficher le quadrillage dans Excel
    ws.views.sheetView[0].showGridLines = True

    # Définition des Styles LPEE
    font_titre = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    font_sub = Font(name="Calibri", size=11, italic=True, color="595959")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_conforme = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_conforme = Font(name="Calibri", size=11, bold=True, color="006100")
    
    fill_non_conforme = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_non_conforme = Font(name="Calibri", size=11, bold=True, color="9C0006")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # En-tête du document
    ws['A1'] = "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE) - SMART CONTROL BÉTON"
    ws['A1'].font = font_titre
    ws['A1'].alignment = align_left

    if est_historique:
        ws['A2'] = "PROJET : LGV CASA SUD | CLIENT : TGCC | Registre Général et Historique Complet"
    else:
        ws['A2'] = f"PROJET : LGV CASA SUD | CLIENT : TGCC | Rapport Journalier du {date_rapport}"
    ws['A2'].font = font_sub
    ws['A2'].alignment = align_left

    # Mapping des noms de colonnes
    col_mapping = {
        "N°": "N°",
        "num_bon_livraison": "N° Bon Livraison",
        "ouvrage": "Ouvrage",
        "element_betonne": "Élément Bétonné",
        "volume_beton": "Volume (m³)",
        "classe_beton": "Classe Béton",
        "date_betonnage": "Date Bétonnage",
        "heure_fin_production_cab": "Fin Prod. CAB",
        "heure_arrivee_chantier": "Arrivée Chantier",
        "tbf": "TBF (°C)",
        "ta": "TA (°C)",
        "affaissement": "Slump (mm)",
        "prelevement": "Prélèvement",
        "technicien": "Technicien",
        "statut": "STATUT"
    }

    df_export = df.copy()
    df_export.insert(0, 'N°', range(1, len(df_export) + 1))
    
    cols_presentes = [c for c in list(col_mapping.keys()) if c in df_export.columns]

    start_row = 4
    # En-têtes (Ligne 4)
    for col_idx, col_key in enumerate(cols_presentes, start=1):
        cell = ws.cell(row=start_row, column=col_idx)
        cell.value = col_mapping[col_key]
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
        ws.row_dimensions[start_row].height = 24

    # Données (Lignes 5+)
    for row_idx, row_data in enumerate(df_export[cols_presentes].values, start=start_row + 1):
        ws.row_dimensions[row_idx].height = 20
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
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

    # Ajustement automatique de la largeur des colonnes
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < start_row:
                continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output)
    return output.getvalue()

# =========================================================
# FONCTION 2 : EXCEL RÉCAPITULATIF MENSUEL STYLISÉ (LPEE & TGCC)
# =========================================================
def generer_excel_recap_mensuel_lpee(df_recap, mois):
    """
    Génère un rapport mensuel haute qualité en format Excel avec en-tête LPEE, 
    informations projet/client (TGCC) et totaux généraux.
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Synthèse {mois.replace('/', '-')}"

    # Afficher le quadrillage
    ws.views.sheetView[0].showGridLines = True

    # Styles
    font_titre = Font(name="Calibri", size=15, bold=True, color="1F4E78")
    font_sub_titre = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    font_card_label = Font(name="Calibri", size=10, bold=True, color="595959")
    font_card_val = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_total = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_card = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    card_border = Border(
        left=Side(style='medium', color='1F4E78'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    total_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='medium', color='1F4E78'),
        bottom=Side(style='double', color='1F4E78')
    )
    
    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # --- 1. CARTOUCHE D'EN-TÊTE LPEE ---
    ws['A1'] = "LABORATOIRE PUBLIC D'ESSAIS ET D'ÉTUDES (LPEE)"
    ws['A1'].font = font_titre

    ws['A2'] = f"SYNTHÈSE MENSUELLE DU SUIVI DE BÉTONNAGE — {mois}"
    ws['A2'].font = font_sub_titre

    # --- 2. BLOC INFOS PROJET & CLIENT ---
    # Ligne 4 : Projet & Client
    ws['A4'] = "PROJET :"
    ws['A4'].font = font_card_label
    ws['B4'] = "LGV CASA SUD"
    ws['B4'].font = font_card_val

    ws['D4'] = "CLIENT / ENTREPRISE :"
    ws['D4'].font = font_card_label
    ws['E4'] = "TGCC"
    ws['E4'].font = font_card_val

    # Ligne 5 : Organisme & Période
    ws['A5'] = "ORGANISME DE CONTRÔLE :"
    ws['A5'].font = font_card_label
    ws['B5'] = "LPEE - Laboratoire LGV Casa Sud"
    ws['B5'].font = font_card_val

    ws['D5'] = "PÉRIODE DE SYNTHÈSE :"
    ws['D5'].font = font_card_label
    ws['E5'] = mois
    ws['E5'].font = font_card_val

    # Application du style de carte d'informations (lignes 4 et 5)
    for r in range(4, 6):
        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_card
            cell.border = thin_border

    # --- 3. EN-TÊTES DU TABLEAU (Ligne 7) ---
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

    # --- 4. REMPLISSAGE DES DONNÉES (Ligne 8+) ---
    current_row = start_row + 1
    for row_data in df_recap.values:
        ws.row_dimensions[current_row].height = 21
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = val
            cell.border = thin_border
            
            # Alignement selon le type de colonne
            if col_idx in [1, 3, 4, 5, 6, 7]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        current_row += 1

    # --- 5. LIGNE DE SYNTHÈSE ET TOTAL MENSUEL ---
    if len(df_recap) > 0:
        ws.row_dimensions[current_row].height = 24
        
        # Titre Synthèse
        c_tot_label = ws.cell(row=current_row, column=1, value="TOTAL / SYNTHÈSE MOIS")
        c_tot_label.font = font_total
        c_tot_label.alignment = align_center
        c_tot_label.fill = fill_total
        c_tot_label.border = total_border

        c_blank = ws.cell(row=current_row, column=2, value="—")
        c_blank.alignment = align_center
        c_blank.font = font_total
        c_blank.fill = fill_total
        c_blank.border = total_border

        # Somme du nombre de contrôles (camions)
        tot_ctrl = int(df_recap["Nombre de contrôles"].sum())
        c_ctrl = ws.cell(row=current_row, column=3, value=tot_ctrl)
        c_ctrl.font = font_total
        c_ctrl.alignment = align_center
        c_ctrl.fill = fill_total
        c_ctrl.border = total_border

        # Valeurs extrêmes du mois (Min/Max Slump et TBF)
        c_aff_min = ws.cell(row=current_row, column=4, value=df_recap["Affaissement Min (mm)"].min())
        c_aff_max = ws.cell(row=current_row, column=5, value=df_recap["Affaissement Max (mm)"].max())
        c_tbf_min = ws.cell(row=current_row, column=6, value=df_recap["TBF Min (°C)"].min())
        c_tbf_max = ws.cell(row=current_row, column=7, value=df_recap["TBF Max (°C)"].max())

        for c in [c_aff_min, c_aff_max, c_tbf_min, c_tbf_max]:
            c.font = font_total
            c.alignment = align_center
            c.fill = fill_total
            c.border = total_border

    # --- 6. AJUSTEMENT DE LA LARGEUR DES COLONNES ---
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row < start_row:
                continue
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 5, 15)

    wb.save(output)
    return output.getvalue()

# =========================================================
# FONCTION DE CALCUL AUTOMATIQUE DU STATUT (Formule Excel)
# =========================================================
def evaluer_statut_beton(tbf, h_fin, h_arr, affaissement):
    try:
        cond_tbf = float(tbf) < 32.1
        cond_aff = 160 <= int(affaissement) <= 220

        fmt = "%H:%M"
        t_fin = datetime.strptime(str(h_fin).strip(), fmt)
        t_arr = datetime.strptime(str(h_arr).strip(), fmt)

        diff_minutes = (t_arr - t_fin).total_seconds() / 60
        if diff_minutes < 0:
            diff_minutes += 24 * 60

        cond_delai = 0 <= diff_minutes <= 120

        if cond_tbf and cond_aff and cond_delai:
            return "✅ Conforme"
        else:
            return "⚠️ Non Conforme"
    except Exception:
        return "⚠️ Non Conforme"

# =========================================================
# 2. VALEURS FIXES & MOTS DE PASSE
# =========================================================
DEFAULT_PROJET = "LGV Casa Sud"
DEFAULT_ENTREPRISE = "TGCC"
DEFAULT_CENTRALE = "TG PREFA"

PASSWORD_GENERAL = "lpee2026"          # Accès technicien
PASSWORD_ADMIN = "lpee_admin_2026"     # Code administrateur

if "authentifie" not in st.session_state:
    st.session_state["authentifie"] = False

if "is_admin" not in st.session_state:
    st.session_state["is_admin"] = False

# Écran de connexion initial
if not st.session_state["authentifie"]:
    st.title("🔒 Accès Sécurisé - LPEE Smart Control Béton")
    st.warning("Veuillez saisir le mot de passe du laboratoire pour accéder à l'application.")
    
    pwd = st.text_input("Mot de passe", type="password")
    if st.button("Valider l'accès"):
        if pwd == PASSWORD_GENERAL or pwd == PASSWORD_ADMIN:
            st.session_state["authentifie"] = True
            if pwd == PASSWORD_ADMIN:
                st.session_state["is_admin"] = True
            st.rerun()
        else:
            st.error("❌ Mot de passe incorrect.")
    st.stop()

# =========================================================
# 3. CONNEXION À SUPABASE
# =========================================================
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "sb_publishable_m8g5mocsCDgk3JpS1lpuCQ_3wOPyet1"  # ⚠️ REMETTEZ VOTRE VRAIE CLÉ SUPABASE ICI

try:
    supabase: Client = create_client(URL, CLE)
    response = supabase.table("controles_beton").select("*").execute()
    data_all = response.data or []
except Exception as e:
    st.error(f"Erreur de connexion à la base de données : {e}")
    data_all = []

# En-tête de l'application
st.title("🚧 Suivi Journalier de Bétonnage - LGV Casa Sud")
st.markdown("### Laboratoire Public d'Essais et d'Études (LPEE) | Client : TGCC")

# =========================================================
# 4. SÉLECTION DE LA JOURNÉE DE BÉTONNAGE
# =========================================================
st.markdown("---")
col_date1, col_date2 = st.columns([1, 2])

with col_date1:
    date_choisie = st.date_input("📅 Choisir la journée de bétonnage :", value=date.today())
    str_date_choisie = date_choisie.strftime("%d/%m/%Y")

# Filtrer les données pour la journée sélectionnée
data_jour = [r for r in data_all if r.get("date_betonnage") == str_date_choisie]

with col_date2:
    st.info(f"📌 **Fichier du jour sélectionné : {str_date_choisie}** | Camions contrôlés aujourd'hui : **{len(data_jour)}**")

# =========================================================
# 5. GESTION PAR ONGLETS
# =========================================================
tab_ajouter, tab_modifier, tab_supprimer, tab_historique, tab_recap_mensuel = st.tabs([
    "➕ Ajouter un camion (Saisie du Jour)", 
    "✏️ Modifier un camion (Admin)", 
    "❌ Supprimer un camion (Admin)",
    "📚 Historique & Tous les Fichiers",
    "📊 Récap Mensuel"
])

# ---------------------------------------------------------
# --- ONGLET 1 : SAISIE D'UN CAMION (JOURNÉE SÉLECTIONNÉE) ---
# ---------------------------------------------------------
with tab_ajouter:
    st.subheader(f"📝 Saisie pour la journée du {str_date_choisie}")
    
    with st.form("form_controle_ajouter"):
        c1, c2, c3 = st.columns(3)
        
        with c1:
            projet = st.text_input("Projet", value=DEFAULT_PROJET, disabled=True)
            entreprise = st.text_input("Entreprise / Client", value=DEFAULT_ENTREPRISE, disabled=True)
            centrale_beton = st.text_input("Centrale béton", value=DEFAULT_CENTRALE, disabled=True)
            
        with c2:
            ouvrage = st.text_input("Ouvrage", value="PRO745 OA1", key="add_ouvrage")
            element_betonne = st.text_input("Élément bétonné", value="Semelle C0", key="add_elem")
            volume_beton = st.text_input("Volume béton (m³)", value="8", key="add_vol")
            
        with c3:
            num_bon_livraison = st.text_input("N° bon livraison (BL)", value="BL2548", key="add_bl")
            classe_beton = st.text_input("Classe béton", value="C30/37", key="add_classe")

        st.markdown("---")
        st.markdown("#### Mesures & Contrôles Chantier")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        
        with col_m1:
            heure_fin_prod = st.text_input("Heure fin prod. CAB (HH:MM)", value="15:28", key="add_h_fin")
            heure_arrivee = st.text_input("Heure arrivée chantier (HH:MM)", value="16:31", key="add_h_arr")
        with col_m2:
            tbf = st.number_input("TBF (°C)", value=32.0, step=0.1, format="%.1f", key="add_tbf")
            ta = st.number_input("TA (°C) - Ambiante", value=28.9, step=0.1, format="%.1f", key="add_ta")
        with col_m3:
            affaissement = st.number_input("Affaissement (mm)", value=170, step=1, format="%d", key="add_aff")
            meteo = st.selectbox("Météo", ["Soleil", "Nuageux", "Pluie", "Vent"], key="add_meteo")
        with col_m4:
            prelevement = st.selectbox("Prélèvement", ["OUI", "NON"], key="add_prelev")
            technicien = st.text_input("Technicien Contrôleur", value="Ismail / Mohamed", key="add_tech")

        observations = st.text_area("Observations", value="RAS", key="add_obs")

        submit_add = st.form_submit_button("💾 Enregistrer le camion dans la journée")
        
        if submit_add:
            statut_auto = evaluer_statut_beton(tbf, heure_fin_prod, heure_arrivee, affaissement)
            
            data_to_insert = {
                "projet": DEFAULT_PROJET,
                "entreprise": DEFAULT_ENTREPRISE,
                "centrale_beton": DEFAULT_CENTRALE,
                "ouvrage": ouvrage,
                "element_betonne": element_betonne,
                "volume_beton": volume_beton,
                "num_bon_livraison": num_bon_livraison,
                "classe_beton": classe_beton,
                "date_betonnage": str_date_choisie,
                "meteo": meteo,
                "observations": observations,
                "heure_fin_production_cab": heure_fin_prod,
                "heure_arrivee_chantier": heure_arrivee,
                "tbf": round(tbf, 1),
                "ta": round(ta, 1),
                "affaissement": int(affaissement),
                "prelevement": prelevement,
                "technicien": technicien,
                "statut": statut_auto
            }
            try:
                supabase.table("controles_beton").insert(data_to_insert).execute()
                st.success(f"Camion BL : {num_bon_livraison} enregistré par {technicien} ! Statut : {statut_auto}")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur d'enregistrement : {e}")

# ---------------------------------------------------------
# --- ONGLET 2 : MODIFIER (Réservé Admin) ---
# ---------------------------------------------------------
with tab_modifier:
    if not st.session_state["is_admin"]:
        st.warning("🔒 La modification est réservée au responsable du laboratoire.")
        pwd_admin_input = st.text_input("Saisissez le code Administrateur :", type="password", key="login_admin_mod")
        if st.button("Débloquer la modification"):
            if pwd_admin_input == PASSWORD_ADMIN:
                st.session_state["is_admin"] = True
                st.success("Accès Administrateur débloqué !")
                st.rerun()
            else:
                st.error("❌ Code Administrateur incorrect.")
    else:
        st.info("🔓 Mode Administrateur Actif")
        if data_jour and len(data_jour) > 0:
            options_edit = {
                f"BL : {row.get('num_bon_livraison')} | Arrivée : {row.get('heure_arrivee_chantier')} | Ouvrage : {row.get('ouvrage')}": row
                for row in data_jour
            }
            choix = st.selectbox("📌 Choisissez le camion à modifier pour cette journée :", list(options_edit.keys()), key="select_edit")
            row_selected = options_edit[choix]

            with st.form("form_controle_modifier"):
                st.markdown(f"#### Modification du camion (BL : {row_selected.get('num_bon_livraison')})")
                c1, c2, c3 = st.columns(3)
                
                with c1:
                    edit_projet = st.text_input("Projet", value=str(row_selected.get('projet') or DEFAULT_PROJET))
                    edit_entreprise = st.text_input("Entreprise / Client", value=str(row_selected.get('entreprise') or DEFAULT_ENTREPRISE))
                    edit_centrale = st.text_input("Centrale béton", value=str(row_selected.get('centrale_beton') or DEFAULT_CENTRALE))
                    
                with c2:
                    edit_ouvrage = st.text_input("Ouvrage", value=str(row_selected.get('ouvrage') or ''))
                    edit_elem = st.text_input("Élément bétonné", value=str(row_selected.get('element_betonne') or ''))
                    edit_vol = st.text_input("Volume béton", value=str(row_selected.get('volume_beton') or ''))
                    
                with c3:
                    edit_bl = st.text_input("N° bon livraison", value=str(row_selected.get('num_bon_livraison') or ''))
                    edit_classe = st.text_input("Classe béton", value=str(row_selected.get('classe_beton') or ''))

                st.markdown("---")
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                with col_m1:
                    edit_h_fin = st.text_input("Heure fin prod. CAB (HH:MM)", value=str(row_selected.get('heure_fin_production_cab') or ''))
                    edit_h_arr = st.text_input("Heure arrivée chantier (HH:MM)", value=str(row_selected.get('heure_arrivee_chantier') or ''))
                with col_m2:
                    edit_tbf = st.number_input("TBF (°C)", value=float(row_selected.get('tbf') or 0.0), step=0.1, format="%.1f", key="edit_tbf_input")
                    edit_ta = st.number_input("TA (°C)", value=float(row_selected.get('ta') or 0.0), step=0.1, format="%.1f", key="edit_ta_input")
                with col_m3:
                    edit_aff = st.number_input("Affaissement (mm)", value=int(row_selected.get('affaissement') or 0), step=1, format="%d", key="edit_aff_input")
                    meteo_opts = ["Soleil", "Nuageux", "Pluie", "Vent"]
                    m_idx = meteo_opts.index(row_selected.get('meteo')) if row_selected.get('meteo') in meteo_opts else 0
                    edit_meteo = st.selectbox("Météo", meteo_opts, index=m_idx, key="edit_meteo_select")
                with col_m4:
                    prelev_opts = ["OUI", "NON"]
                    p_idx = prelev_opts.index(row_selected.get('prelevement')) if row_selected.get('prelevement') in prelev_opts else 0
                    edit_prelev = st.selectbox("Prélèvement", prelev_opts, index=p_idx, key="edit_prelev_select")
                    edit_tech = st.text_input("Technicien Contrôleur", value=str(row_selected.get('technicien') or ''), key="edit_tech_input")

                edit_obs = st.text_area("Observations", value=str(row_selected.get('observations') or ''))

                submit_update = st.form_submit_button("🔄 Mettre à jour cette entrée")

                if submit_update:
                    statut_recalcule = evaluer_statut_beton(edit_tbf, edit_h_fin, edit_h_arr, edit_aff)
                    update_data = {
                        "projet": edit_projet, "entreprise": edit_entreprise, "centrale_beton": edit_centrale,
                        "ouvrage": edit_ouvrage, "element_betonne": edit_elem,
                        "volume_beton": edit_vol, "num_bon_livraison": edit_bl,
                        "classe_beton": edit_classe, "date_betonnage": str_date_choisie, "meteo": edit_meteo,
                        "observations": edit_obs, "heure_fin_production_cab": edit_h_fin,
                        "heure_arrivee_chantier": edit_h_arr, "tbf": round(edit_tbf, 1), "ta": round(edit_ta, 1),
                        "affaissement": int(edit_aff), "prelevement": edit_prelev, 
                        "technicien": edit_tech, "statut": statut_recalcule
                    }
                    try:
                        supabase.table("controles_beton").update(update_data).eq("id", row_selected["id"]).execute()
                        st.success(f"Mise à jour réussie ! Nouveau statut : {statut_recalcule}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur lors de la mise à jour : {e}")
        else:
            st.info(f"Aucun camion enregistré pour le {str_date_choisie}.")

# ---------------------------------------------------------
# --- ONGLET 3 : SUPPRIMER (Réservé Admin) ---
# ---------------------------------------------------------
with tab_supprimer:
    if not st.session_state["is_admin"]:
        st.warning("🔒 La suppression est réservée au responsable du laboratoire.")
        pwd_admin_input_del = st.text_input("Saisissez le code Administrateur :", type="password", key="login_admin_del")
        if st.button("Débloquer la suppression"):
            if pwd_admin_input_del == PASSWORD_ADMIN:
                st.session_state["is_admin"] = True
                st.success("Accès Administrateur débloqué !")
                st.rerun()
            else:
                st.error("❌ Code Administrateur incorrect.")
    else:
        st.info("🔓 Mode Administrateur Actif")
        if data_jour and len(data_jour) > 0:
            options_del = {
                f"BL : {row.get('num_bon_livraison')} | Heure : {row.get('heure_arrivee_chantier')} | Ouvrage : {row.get('ouvrage')}": row["id"]
                for row in data_jour
            }
            choix_del = st.selectbox("⚠️ Choisissez la ligne à supprimer :", list(options_del.keys()), key="select_del")
            id_to_delete = options_del[choix_del]

            if st.button("🚨 Confirmer la suppression définitive", type="primary"):
                try:
                    supabase.table("controles_beton").delete().eq("id", id_to_delete).execute()
                    st.success("Entrée supprimée avec succès !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur lors de la suppression : {e}")
        else:
            st.info(f"Aucun enregistrement à supprimer pour le {str_date_choisie}.")

# ---------------------------------------------------------
# --- ONGLET 4 : HISTORIQUE COMPLET ---
# ---------------------------------------------------------
with tab_historique:
    st.subheader("📚 Registre Général (Toutes les journées confondues)")
    if data_all and len(data_all) > 0:
        df_all = pd.DataFrame(data_all)
        colonnes_historique = [
            "num_bon_livraison", "ouvrage", "element_betonne", "volume_beton", 
            "classe_beton", "date_betonnage", "heure_fin_production_cab", 
            "heure_arrivee_chantier", "tbf", "ta", "affaissement", "prelevement", "technicien", "statut"
        ]
        df_hist_vis = df_all[[c for c in colonnes_historique if c in df_all.columns]]
        df_hist_vis.index = range(1, len(df_hist_vis) + 1)
        st.dataframe(df_hist_vis, use_container_width=True)
        
        # Téléchargement EXCEL Formaté de tout le registre
        excel_all_bytes = generer_excel_lpee(df_hist_vis, est_historique=True)
        st.download_button(
            label="📊 Télécharger l'historique complet en Excel (.xlsx)",
            data=excel_all_bytes,
            file_name="Registre_Complet_Beton_LGV.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------------------------------------------------------
# --- ONGLET 5 : RÉCAPITULATIF MENSUEL (STYLISÉ LPEE / TGCC) ---
# ---------------------------------------------------------
with tab_recap_mensuel:
    st.subheader("📊 Récapitulatif Mensuel par Date et Ouvrage / Élément Bétonné")
    if data_all and len(data_all) > 0:
        df_recap_raw = pd.DataFrame(data_all)
        
        # Conversion de la colonne date
        df_recap_raw["dt"] = pd.to_datetime(df_recap_raw["date_betonnage"], format="%d/%m/%Y", errors="coerce")
        df_valid = df_recap_raw.dropna(subset=["dt"]).copy()
        
        if not df_valid.empty:
            df_valid["mois_annee"] = df_valid["dt"].dt.strftime("%m/%Y")
            liste_mois = sorted(df_valid["mois_annee"].unique(), reverse=True)
            
            mois_selectionne = st.selectbox("📅 Sélectionner le mois :", liste_mois, key="select_mois_recap")
            
            # Filtrage du mois
            df_mois = df_valid[df_valid["mois_annee"] == mois_selectionne].copy()
            
            df_mois["ouvrage_clean"] = df_mois["ouvrage"].fillna("").astype(str).str.strip()
            df_mois["element_clean"] = df_mois["element_betonne"].fillna("").astype(str).str.strip()
            
            def combiner_ouvrage_element(row):
                o = row["ouvrage_clean"]
                e = row["element_clean"]
                if o and e:
                    return f"{o} - {e}"
                elif o:
                    return o
                elif e:
                    return e
                else:
                    return "Non spécifié"
            
            df_mois["ouvrage_element"] = df_mois.apply(combiner_ouvrage_element, axis=1)
            
            df_mois["affaissement"] = pd.to_numeric(df_mois["affaissement"], errors="coerce")
            df_mois["tbf"] = pd.to_numeric(df_mois["tbf"], errors="coerce")
            
            # Groupement par Date et Ouvrage/Élément avec décompte des contrôles
            df_recap = df_mois.groupby(["dt", "date_betonnage", "ouvrage_element"]).agg(
                nb_controles=("ouvrage_element", "count"),
                aff_min=("affaissement", "min"),
                aff_max=("affaissement", "max"),
                tbf_min=("tbf", "min"),
                tbf_max=("tbf", "max")
            ).reset_index().sort_values(["dt", "ouvrage_element"])
            
            df_recap_final = df_recap[["date_betonnage", "ouvrage_element", "nb_controles", "aff_min", "aff_max", "tbf_min", "tbf_max"]].rename(columns={
                "date_betonnage": "Date",
                "ouvrage_element": "Ouvrage / Élément Bétonné",
                "nb_controles": "Nombre de contrôles",
                "aff_min": "Affaissement Min (mm)",
                "aff_max": "Affaissement Max (mm)",
                "tbf_min": "TBF Min (°C)",
                "tbf_max": "TBF Max (°C)"
            })
            
            df_recap_final.index = range(1, len(df_recap_final) + 1)
            
            st.markdown(f"#### Synthèse du mois de **{mois_selectionne}** — Projet : **LGV CASA SUD** | Client : **TGCC**")
            st.dataframe(df_recap_final, use_container_width=True)
            
            # Génération du fichier Excel Stylisé LPEE
            excel_recap_bytes = generer_excel_recap_mensuel_lpee(df_recap_final, mois_selectionne)
            nom_fichier_recap = f"Recap_Mensuel_Beton_LPEE_{mois_selectionne.replace('/', '_')}.xlsx"
            
            st.download_button(
                label=f"📊 Télécharger le Récapitulatif Mensuel Officiel LPEE (.xlsx)",
                data=excel_recap_bytes,
                file_name=nom_fichier_recap,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("Aucune donnée avec une date valide pour générer le récapitulatif.")
    else:
        st.info("Aucune donnée enregistrée pour le moment.")

# =========================================================
# 6. TABLEAU & FICHIER DE LA JOURNÉE SÉLECTIONNÉE
# =========================================================
st.markdown("---")
st.subheader(f"📋 Fichier et Tableau de Suivi pour le : {str_date_choisie}")

if data_jour and len(data_jour) > 0:
    df_jour = pd.DataFrame(data_jour)
    
    # CALCULS DU RÉCAPITULATIF JOURNALIER
    df_jour["vol_num"] = pd.to_numeric(df_jour["volume_beton"], errors="coerce").fillna(0)
    vol_tot_jour = df_jour["vol_num"].sum()
    nb_camions_jour = len(df_jour)
    
    vol_tot_global = 0.0
    if data_all:
        df_g = pd.DataFrame(data_all)
        if "volume_beton" in df_g.columns:
            vol_tot_global = pd.to_numeric(df_g["volume_beton"], errors="coerce").fillna(0).sum()

    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("🏗️ Quantité Totale Béton du Jour", f"{vol_tot_jour:.2f} m³")
    with m2:
        st.metric("🚛 Nombre de Camions (Jour)", f"{nb_camions_jour}")
    with m3:
        st.metric("📈 Total Cumulé Global Chantier", f"{vol_tot_global:.2f} m³")

    st.markdown("---")

    colonnes_visibles = [
        "num_bon_livraison", "ouvrage", "element_betonne", 
        "volume_beton", "classe_beton", "heure_fin_production_cab", 
        "heure_arrivee_chantier", "tbf", "ta", "affaissement", "prelevement", "technicien", "statut"
    ]
    df_affichage = df_jour[[c for c in colonnes_visibles if c in df_jour.columns]]
    df_affichage.index = range(1, len(df_affichage) + 1)
    
    st.dataframe(df_affichage, use_container_width=True)

    # Téléchargement EXCEL Formaté du jour
    excel_jour_bytes = generer_excel_lpee(df_affichage, date_rapport=str_date_choisie)
    nom_excel_jour = f"Rapport_Betonnage_{date_choisie.strftime('%Y-%m-%d')}.xlsx"
    
    st.download_button(
        label=f"📊 Télécharger le rapport Excel stylisé LPEE ({nom_excel_jour})",
        data=excel_jour_bytes,
        file_name=nom_excel_jour,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.warning(f"Aucun suivi enregistré pour la journée du {str_date_choisie}. Utilisez le formulaire ci-dessus pour ajouter le premier camion !")
