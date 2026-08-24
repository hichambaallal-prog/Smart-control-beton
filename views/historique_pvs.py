import io
import re
from datetime import date, datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
import pandas as pd
import streamlit as st


# ==============================================================================
# 1. GESTION DES UTILISATEURS ET CONNEXION SUPABASE
# ==============================================================================
def connecter_utilisateur(supabase, nom_utilisateur, mot_de_passe):
    """
    Vérifie le nom d'utilisateur, le mot de passe et récupère le champ 'can_edit'
    depuis la table 'users' de Supabase.
    """
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
            st.session_state["user_logged"] = True
            st.session_state["user"] = user_info
            st.session_state["username"] = user_info.get("username")
            st.session_state["role"] = user_info.get("role")
            st.session_state["user_role"] = user_info.get("role")
            st.session_state["can_edit"] = bool(user_info.get("can_edit", False))
            return True
        else:
            st.error("Nom d'utilisateur ou mot de passe incorrect.")
            return False
    except Exception as e:
        st.error(f"Erreur lors de la connexion : {e}")
        return False


# =========================================================
# FONCTION UTILITAIRE : VÉRIFICATION DES DOUBLONS DU N° DE RÉCEPTION
# =========================================================
def verifier_doublon_num_reception(supabase, num_reception, current_beton_id=None):
    """
    Vérifie dans la table 'suivi_betonnage' si le num_reception existe déjà.
    Retourne True si le numéro est en doublon, sinon False.
    """
    if not num_reception or str(num_reception).strip() in ["", "-", "None", "NaN", "N/A"]:
        return False
        
    num_clean = str(num_reception).strip()
    try:
        # Recherche par num_reception
        res1 = (
            supabase.table("suivi_betonnage")
            .select("id, num_reception")
            .eq("num_reception", num_clean)
            .execute()
        )
        # Recherche par num_reception si applicable
        res2 = (
            supabase.table("suivi_betonnage")
            .select("id, num_reception")
            .eq("num_reception", num_clean)
            .execute()
        )
        
        matches = (res1.data or []) + (res2.data or [])
        
        for m in matches:
            # S'il existe un enregistrement ayant le même numéro mais un ID différent
            if current_beton_id is None or int(m.get("id")) != int(current_beton_id):
                return True
    except Exception as e:
        st.warning(f"Note lors de la vérification des doublons : {e}")
        
    return False


# =========================================================
# FONCTION UTILITAIRE : EXTRACTION SÉCURISÉE DU N° BL
# =========================================================
def extraire_num_bl(*sources):
    """Inspecte récursivement les sources pour extraire le N° de Bon de Livraison (BL)."""
    clefs_possibles = [
        "num_bl",
        "bl",
        "num_bon_livraison",
        "n_bl",
        "bon_livraison",
        "num_bl_p",
        "n_bon",
        "bon_de_livraison",
        "code_bl",
    ]

    for source in sources:
        if isinstance(source, dict):
            for key in clefs_possibles:
                val = source.get(key)
                if val is not None:
                    val_str = str(val).strip()
                    if val_str and val_str.upper() not in [
                        "N/A",
                        "NONE",
                        "NAN",
                        "-",
                        "",
                    ]:
                        return val_str

            for key, val in source.items():
                if "bl" in key.lower() or "bon" in key.lower():
                    if val is not None:
                        val_str = str(val).strip()
                        if val_str and val_str.upper() not in [
                            "N/A",
                            "NONE",
                            "NAN",
                            "-",
                            "",
                        ]:
                            return val_str

        elif isinstance(source, str):
            match = re.search(r"BL\s*:\s*([^\|]+)", source, re.IGNORECASE)
            if match:
                val_str = match.group(1).strip()
                if val_str and val_str.upper() not in [
                    "N/A",
                    "NONE",
                    "NAN",
                    "-",
                    "",
                ]:
                    return val_str

    return "-"


# =========================================================
# 2. GÉNÉRATION DU PROCÈS-VERBAL EXCEL (FORMAT EXACT LPEE)
# =========================================================
def generer_pv_excel(export_data, infos_header):
    """Génère un Procès-Verbal (PV) d'écrasement de béton répliquant le modèle LPEE."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PV Écrasement LPEE"

    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.page_margins = PageMargins(
        left=0.3, right=0.3, top=0.4, bottom=0.4, header=0.2, footer=0.2
    )

    font_bold = Font(name="Calibri", size=9, bold=True)
    font_bold_white = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_title_white = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=8.5)
    font_small = Font(name="Calibri", size=8)

    fill_header_dark = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    fill_header_table = PatternFill(
        start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
    )
    fill_section_label = PatternFill(
        start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
    )

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)

    thin_side = Side(border_style="thin", color="000000")
    border_cell = Border(
        left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
    )

    default_bl = extraire_num_bl(infos_header)

    def remplacer_na(valeur, fallback=None):
        val_str = str(valeur).strip() if valeur is not None else ""
        if val_str.upper() in ["N/A", "NONE", "NAN", "", "-"]:
            return fallback if fallback is not None else default_bl
        return valeur

    # ENTÊTE
    ws.merge_cells("A1:D1")
    ws["A1"] = "LPEE / CTR CSB"
    ws["A1"].font = font_bold_white
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:D3")
    ws["A2"] = "Laboratoire de Contrôle Externe"
    ws["A2"].font = font_bold_white
    ws["A2"].alignment = align_center

    for r in range(1, 4):
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = fill_header_dark

    ws["E1"] = "RE N° :"
    ws["E1"].font = font_bold
    
    # ----------- MODIFICATIONS COLONNES F, G et H -----------
    ws.merge_cells("F1:G1")
    ws["F1"] = remplacer_na(infos_header.get("re_num"), "25/260/LGV/ B/")
    ws["F1"].font = font_regular
    
    ws["H1"] = "BETON"
    ws["H1"].font = font_bold
    ws["H1"].alignment = align_center
    # --------------------------------------------------------

    ws["E2"] = "DOSSIER :"
    ws["E2"].font = font_bold
    ws.merge_cells("F2:H2")
    ws["F2"] = remplacer_na(infos_header.get("dossier"), "2025-260-05985-2025-0247")
    ws["F2"].font = font_regular

    ws["E3"] = "CLIENT :"
    ws["E3"].font = font_bold
    ws.merge_cells("F3:H3")
    ws["F3"] = remplacer_na(infos_header.get("client"), "TGCC")
    ws["F3"].font = font_bold

    for r in range(1, 4):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    # TITRE
    ws.merge_cells("A4:H4")
    ws["A4"] = "ESSAIS MECANIQUES SUR BETON HYDRAULIQUE"
    ws["A4"].font = font_title_white
    ws["A4"].alignment = align_center
    for c in range(1, 9):
        ws.cell(row=4, column=c).fill = fill_header_dark
        ws.cell(row=4, column=c).border = border_cell

    ws.merge_cells("A5:D5")
    ws["A5"] = "[X] COMPRESSION NF EN 12390-3 (2019)"
    ws["A5"].font = font_bold
    ws["A5"].alignment = align_center

    ws.merge_cells("E5:H5")
    ws["E5"] = "[ ] TRACTION PAR FENDAGE NF EN 12390-6 (2019)"
    ws["E5"].font = font_bold
    ws["E5"].alignment = align_center

    for c in range(1, 9):
        ws.cell(row=5, column=c).border = border_cell

    ws.merge_cells("A6:F6")
    ws["A6"] = "Presse : Marque: Controls"
    ws["A6"].font = font_bold
    ws["A6"].alignment = align_right

    ws.merge_cells("G6:H6")
    ws["G6"] = "Classe : A"
    ws["G6"].font = font_bold
    ws["G6"].alignment = align_center

    for c in range(1, 9):
        ws.cell(row=6, column=c).border = border_cell

    # FICHE TECHNIQUE
    ws["A7"] = "Date de\nprélèvement"
    ws["A7"].font = font_bold
    ws["A7"].alignment = align_center
    ws["B7"] = str(remplacer_na(infos_header.get("date_coulee"), "-"))
    ws["B7"].font = font_bold
    ws["B7"].alignment = align_center

    ws.merge_cells("C7:D7")
    ws["C7"] = "Lieu de\nprélèvement"
    ws["C7"].font = font_bold
    ws["C7"].alignment = align_center

    ws.merge_cells("E7:H7")
    ws["E7"] = remplacer_na(
        infos_header.get("lieu_prelevement", infos_header.get("ouvrage")), "-"
    )
    ws["E7"].font = font_regular
    ws["E7"].alignment = align_center

    ws["A8"] = "Chantier"
    ws["A8"].font = font_bold
    ws["A8"].alignment = align_center

    ws.merge_cells("B8:D8")
    ws["B8"] = remplacer_na(
        infos_header.get("chantier"),
        "Augmentation de la capacité ferroviaire entre Kenitra et Marrakech et au niveau du hub de Casablanca\nTravaux d'exécution de terrassement, ouvrages d'art et rétablissement de communication entre PK 5+450 et PK 10+000-GARE CASA SUD",
    )
    ws["B8"].font = font_small
    ws["B8"].alignment = align_center

    ws.merge_cells("E8:F8")
    ws["E8"] = "Type de béton"
    ws["E8"].font = font_bold
    ws["E8"].alignment = align_center

    ws.merge_cells("G8:H8")
    classe_beton_val = str(remplacer_na(infos_header.get("classe_beton"), "C35/45")).upper()
    ws["G8"] = classe_beton_val
    ws["G8"].font = font_bold
    ws["G8"].alignment = align_center

    centrale_saisie = remplacer_na(infos_header.get("centrale"), "Centrale à Béton")
    ws.merge_cells("A9:B9")
    ws["A9"] = centrale_saisie
    ws["A9"].font = font_bold
    ws["A9"].alignment = align_center

    ws["C9"] = "- Dimensions"
    ws["C9"].font = font_regular

    ws.merge_cells("D9:H9")
    ws["D9"] = remplacer_na(infos_header.get("forme"), "Cylindrique 150x300")
    ws["D9"].font = font_bold
    ws["D9"].alignment = align_center

    ws.merge_cells("A10:B10")
    ws["A10"] = "Affaissement au cône d'abrams NF EN 12350-2"
    ws["A10"].font = font_small
    ws["A10"].alignment = align_center

    ws["C10"] = str(remplacer_na(infos_header.get("affaissement"), "-"))
    ws["C10"].font = font_bold
    ws["C10"].alignment = align_center

    ws["D10"] = "- Mode confection"
    ws["D10"].font = font_regular
    ws["D10"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("E10:H10")
    ws["E10"] = "Par vibration NF EN 12390-2 (2019)"
    ws["E10"].font = font_bold
    ws["E10"].alignment = align_center

    ws.merge_cells("A11:B11")
    ws["A11"] = "Température °C"
    ws["A11"].font = font_regular
    ws["A11"].alignment = align_center

    ws["C11"] = str(remplacer_na(infos_header.get("temperature"), "-"))
    ws["C11"].font = font_bold
    ws["C11"].alignment = align_center

    ws["D11"] = "- Mode conservation"
    ws["D11"].font = font_regular
    ws["D11"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("E11:H11")
    ws["E11"] = "au laboratoire par immersion dans l'eau NF EN 12390-2 (2019) à 20°C ± 2°C"
    ws["E11"].font = font_bold
    ws["E11"].alignment = align_center

    tech_prelevement = remplacer_na(
        infos_header.get("technicien_prelevement") 
        or infos_header.get("preleve_par") 
        or infos_header.get("technicien"), 
        "Technicien LPEE"
    )
    ws.merge_cells("A12:C12")
    ws["A12"] = f"prélèvement effectué par {tech_prelevement}"
    ws["A12"].font = font_small
    ws["A12"].alignment = align_center

    ws.merge_cells("D12:E12")
    ws["D12"] = "N° de bon de livraison"
    ws["D12"].font = font_regular
    ws["D12"].alignment = align_center

    ws.merge_cells("F12:H12")
    ws["F12"] = str(default_bl)
    ws["F12"].font = font_bold
    ws["F12"].alignment = align_center

    for r in range(7, 13):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border_cell

    for r in [7, 8, 9, 10, 11, 12]:
        ws.row_dimensions[r].height = 25

    # TABLEAU DES RÉSULTATS
    row_idx = 14
    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)
    ws.cell(row=row_idx, column=1).value = "TABLEAU DES RÉSULTATS"
    ws.cell(row=row_idx, column=1).font = font_title_white
    ws.cell(row=row_idx, column=1).fill = fill_header_dark
    ws.cell(row=row_idx, column=1).alignment = align_center
    for c in range(1, 9):
        ws.cell(row=row_idx, column=c).border = border_cell
        
    row_idx += 1

    colonnes = [
        "Repère \néprouvette",
        "Age\n(jours)",
        "Date\nd'écrasement",
        "Poids\n(Kg)",
        "Surface\n(cm²)",
        "Charge de \nrupture\n(kN)",
        "Résistance\n(MPa)",
        "Type de\nrupture",
    ]
    for c, titre in enumerate(colonnes, 1):
        ws.cell(row=row_idx, column=c).value = titre
        ws.cell(row=row_idx, column=c).font = font_bold
        ws.cell(row=row_idx, column=c).fill = fill_header_table
        ws.cell(row=row_idx, column=c).alignment = align_center
        ws.cell(row=row_idx, column=c).border = border_cell

    ws.row_dimensions[row_idx].height = 40
    row_idx += 1

    total_resistance = 0
    nb_eprouvettes = 0

    if not export_data.empty:
        for idx, row in export_data.iterrows():
            ech_str = str(row.get("echeance", "28 jours"))
            ech_match = re.search(r"(\d+)", ech_str)
            age = int(ech_match.group(1)) if ech_match else 28

            surface = float(row.get("section", 176.71)) if pd.notna(row.get("section")) else 176.71
            force_kn = float(row.get("force_kn", 0)) if pd.notna(row.get("force_kn")) else 0
            resistance_mpa = (
                float(row.get("fc_mpa", 0))
                if pd.notna(row.get("fc_mpa"))
                else (force_kn / surface) * 10
            )
            masse = (
                float(row.get("masse", 12.8))
                if pd.notna(row.get("masse"))
                else "-"
            )

            ws.cell(row=row_idx, column=1).value = str(remplacer_na(row.get("repere_eprouvette"), "-"))
            ws.cell(row=row_idx, column=2).value = age
            ws.cell(row=row_idx, column=3).value = str(remplacer_na(row.get("date_ecrasement"), "-"))
            ws.cell(row=row_idx, column=4).value = (
                f"{masse:.1f}" if isinstance(masse, float) else masse
            )
            ws.cell(row=row_idx, column=5).value = f"{surface:.2f}"
            ws.cell(row=row_idx, column=6).Pour masquer définitivement ces colonnes dans votre application Streamlit, la méthode la plus fiable consiste à les supprimer (ou "dropper") directement du DataFrame `pandas` **avant** de l'afficher avec `st.dataframe()`.

D'après l'image `AdobeExpressPhotos_23d2a6d67841434d88eb6a13c2d8ffae_CopyEdited.jpg`, des colonnes comme `created_at`, `masse`, `reference_controle`, et `num_reception` (qui sont vides ou inutiles pour cette vue) apparaissent toujours.

Voici le code complet et corrigé pour cette page. J'ai ajouté une étape explicite qui filtre les colonnes à cacher.

### Le Code Complet

```python
import streamlit as st
import pandas as pd
# Importez votre client Supabase si ce n'est pas déjà fait dans un fichier principal
# from supabase import create_client, Client 

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Base de données globale", layout="wide")

st.title("📊 Base de données globale")

# --- ZONES DE RECHERCHE ---
col1, col2 = st.columns(2)
with col1:
    search_ref = st.text_input("🔍 Recherche par Réf. Contrôle", placeholder="Ex: REF-123-GARE CASA SUD")
with col2:
    search_date = st.date_input("🗓️ Recherche par Date de coulée", value=None)

# --- FONCTION DE RÉCUPÉRATION DES DONNÉES (À adapter avec votre code Supabase exact) ---
@st.cache_data(ttl=60) # Cache optionnel pour optimiser les requêtes
def load_data():
    # EXEMPLE SUPABASE:
    # response = st.session_state.supabase.table("votre_table_eprouvettes").select("*").execute()
    # df = pd.DataFrame(response.data)
    
    # Code simulé pour l'exemple (à remplacer par votre vraie requête)
    # df = pd.DataFrame(...) 
    pass 

# --- LOGIQUE D'AFFICHAGE ---
try:
    # 1. Charger les données
    # df = load_data() 
    
    # -------------------------------------------------------------------------
    # ASSUMONS QUE df EST VOTRE DATAFRAME RÉCUPÉRÉ DEPUIS SUPABASE.
    # Décommentez et utilisez votre vrai DataFrame ici.
    # -------------------------------------------------------------------------
    
    if not df.empty:
        
        # 2. Appliquer les filtres de recherche (si l'utilisateur a tapé quelque chose)
        if search_ref:
            df = df[df['ref_controle'].str.contains(search_ref, case=False, na=False)]
            
        if search_date:
            # Assurez-vous que la colonne date est bien au format datetime pour filtrer
            # df['date_coulee'] = pd.to_datetime(df['date_coulee']).dt.date
            df = df[df['date_coulee'] == search_date]

        # 3. 🔴 L'ÉTAPE CLÉ : LISTE DES COLONNES À MASQUER 🔴
        # Ajoutez ou retirez les noms exacts des colonnes de votre base de données Supabase
        colonnes_a_cacher = [
            'created_at', 
            'masse', 
            'reference_controle', 
            'num_reception',
            # 'force_kn',   # Décommentez si vous voulez aussi cacher celles-ci
            # 'fc_mpa', 
            # 'technicien', 
            # 'observations'
        ]
        
        # On vérifie que les colonnes existent bien dans le DataFrame avant de les supprimer pour éviter les erreurs
        colonnes_presentes_a_cacher = [col for col in colonnes_a_cacher if col in df.columns]
        
        # On crée un nouveau DataFrame nettoyé pour l'affichage
        df_affichage = df.drop(columns=colonnes_presentes_a_cacher)

        # 4. Affichage du DataFrame propre
        st.dataframe(
            df_affichage, 
            use_container_width=True, 
            hide_index=True # Cache la colonne des numéros (0, 1, 2...)
        )
        
    else:
        st.info("Aucune donnée trouvée dans la base de données.")

except Exception as e:
    st.error(f"Une erreur est survenue lors du chargement des données : {e}")
