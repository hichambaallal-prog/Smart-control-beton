import streamlit as st
import pandas as pd
from datetime import date, datetime
from supabase import create_client, Client

# ==============================================================================
# ⚙️ 1. CONFIGURATION DE LA PAGE
# ==============================================================================
st.set_page_config(
    page_title="LPEE CTR-CSB - LGV CASA SUD", 
    page_icon="🏗️", 
    layout="wide"
)

# ==============================================================================
# 🔐 2. GESTION DU MOT DE PASSE ET AUTHENTIFICATION
# ==============================================================================
MOT_DE_PASSE_ACCES = "lpee2026"  # 👈 Mot de passe d'accès

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# --- ÉCRAN DE CONNEXION ---
if not st.session_state["authenticated"]:
    col_g, col_c, col_d = st.columns([1, 2, 1])
    
    with col_c:
        url_image_al_boraq = "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/ONCF_Al_boraq.jpeg/1280px-ONCF_Al_boraq.jpeg"
        st.image(url_image_al_boraq, caption="Projet LGV CASA SUD - LPEE CTR-CSB", use_container_width=True)
        
        st.title("🔒 Connexion au Portail Laboratoire")
        st.markdown("##### **LPEE - CTR-CSB** | Projet : **LGV CASA SUD** | Client : **TGCC**")
        st.markdown("---")
        
        pwd_input = st.text_input("Veuillez saisir le mot de passe :", type="password")
        
        if st.button("Se connecter", type="primary", use_container_width=True):
            if pwd_input == MOT_DE_PASSE_ACCES:
                st.session_state["authenticated"] = True
                st.success("Accès autorisé !")
                st.rerun()
            else:
                st.error("❌ Mot de passe incorrect.")
                
    st.stop()

# ==============================================================================
# ⚙️ 3. CONNEXION SUPABASE
# ==============================================================================
URL = "https://yqijsvxyrdymcnqluipa.supabase.co"
CLE = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlxaWpzdnh5cmR5bWNucWx1aXBhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5NDIwMjIsImV4cCI6MjEwMTUxODAyMn0.xjYXfGqea7P8kK8df9ootEJywCz-zoOzt8LESNRo2i0"

try:
    supabase: Client = create_client(URL, CLE)
except Exception as e:
    st.error(f"Erreur Supabase : {e}")
    st.stop()

# ==============================================================================
# 📌 4. BARRE LATÉRALE DE NAVIGATION
# ==============================================================================
with st.sidebar:
    st.title("🏢 LPEE - CTR-CSB")
    st.caption("Projet : **LGV CASA SUD** | Client : **TGCC**")
    st.write("---")
    
    page = st.radio(
        "📌 Menu Principal",
        [
            "🏠 Accueil", 
            "🏗️ Suivi de Bétonnage", 
            "🪨 Essai à la Plaque"
        ]
    )
    
    st.write("---")
    if st.button("🚪 Déconnexion", type="secondary", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

# ==============================================================================
# 📄 5. CONTENU DES MODULES
# ==============================================================================

# ------------------------------------------------------------------------------
# PAGE 1 : ACCUEIL
# ------------------------------------------------------------------------------
if page == "🏠 Accueil":
    st.title("👋 Bienvenue sur le Portail de Contrôle Qualité")
    st.subheader("Laboratoire Public d'Essais et d'Études (LPEE)")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.markdown("""
        Ce portail rassemble la gestion et le suivi des essais de contrôle qualité sur site :
        
        * **🏗️ Suivi de Bétonnage :** Enregistrement des bons de livraison (BL), calculs des volumes journaliers et mensuels.
        * **🪨 Essai à la Plaque :** Saisie des enfoncements Z1 et Z2, calculs des modules EV1, EV2 et du rapport k.
        """)
    with col2:
        st.info("""
        **Rappels Projet :**
        * **Projet :** LGV CASA SUD
        * **Entreprise :** TGCC
        * **Centre :** CTR-CSB
        """)

# ------------------------------------------------------------------------------
# PAGE 2 : SUIVI DE BÉTONNAGE (Avec Récapitulatifs & Connexion Supabase)
# ------------------------------------------------------------------------------
elif page == "🏗️ Suivi de Bétonnage":
    st.title("🏗️ Suivi et Contrôle Qualité Béton")
    
    # --- A. Chargement des données Supabase ---
    try:
        resp_b = supabase.table("suivi_betonnage").select("*").execute()
        data_beton = resp_b.data or []
    except Exception:
        try:
            # Fallback si le nom de la table dans votre Supabase est 'beton'
            resp_b = supabase.table("beton").select("*").execute()
            data_beton = resp_b.data or []
        except Exception:
            data_beton = []

    df_beton = pd.DataFrame(data_beton)

    # Convertir la colonne date si elle existe
    if not df_beton.empty and "date_livraison" in df_beton.columns:
        df_beton["date_parsed"] = pd.to_datetime(df_beton["date_livraison"], errors="coerce")
    else:
        df_beton["date_parsed"] = pd.NaT

    # --- B. Section Récapitulatifs (Journalier & Mensuel) ---
    st.subheader("📊 Récapitulatifs des Coulages")
    
    col_filtre_date, col_m1, col_m2, col_m3 = st.columns([1.5, 1, 1, 1])
    
    with col_filtre_date:
        date_recap = st.date_input("📅 Sélectionner la date :", value=date(2026, 8, 1))
        str_date_recap = date_recap.strftime("%Y-%m-%d")

    # Calculs Journaliers
    if not df_beton.empty and "volume" in df_beton.columns:
        df_jour = df_beton[df_beton["date_parsed"].dt.strftime("%Y-%m-%d") == str_date_recap]
        vol_jour = df_jour["volume"].astype(float).sum() if not df_jour.empty else 0.0
        nb_bl_jour = len(df_jour)

        # Calculs Mensuels (basés sur le mois sélectionné)
        mois_sel = date_recap.month
        annee_sel = date_recap.year
        df_mois = df_beton[
            (df_beton["date_parsed"].dt.month == mois_sel) & 
            (df_beton["date_parsed"].dt.year == annee_sel)
        ]
        vol_mois = df_mois["volume"].astype(float).sum() if not df_mois.empty else 0.0
        nb_bl_mois = len(df_mois)
    else:
        vol_jour, nb_bl_jour, vol_mois, nb_bl_mois = 0.0, 0, 0.0, 0

    with col_m1:
        st.metric("Volume du Jour", f"{vol_jour:.1f} m³", delta=f"{nb_bl_jour} Bon(s) BL")
    with col_m2:
        st.metric(f"Volume Mensuel ({date_recap.strftime('%m/%Y')})", f"{vol_mois:.1f} m³", delta=f"{nb_bl_mois} Bon(s) BL")
    with col_m3:
        st.metric("Total Général Cumulé", f"{df_beton['volume'].astype(float).sum() if not df_beton.empty and 'volume' in df_beton.columns else 0.0:.1f} m³")

    st.markdown("---")

    # --- C. Formulaire de Saisie d'un nouveau bétonnage ---
    with st.form("form_beton"):
        st.subheader("📝 Nouvel enregistrement de bétonnage")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            date_b = st.date_input("Date de livraison", value=date_recap)
            bl_num = st.text_input("N° Bon de Livraison (BL)", value="BL-2026-001")
            classe_b = st.selectbox("Classe de béton", ["C20/25", "C25/30", "C30/37", "C35/45", "C40/50"])
            
        with c2:
            element = st.text_input("Élément d'ouvrage / PK", value="Semelle P1 - PK 14+200")
            volume_b = st.number_input("Volume (m³)", value=8.0, step=0.5, min_value=0.1)
            temp_b = st.number_input("Température (°C)", value=21.0, step=0.5)

        with c3:
            slump_b = st.number_input("Affaissement / Slump (cm)", value=16.0, step=0.5)
            prelev_b = st.selectbox("Prélèvement (NF EN 12390-2)", ["OUI - Conforme", "NON", "Sans objet"])
            nb_ep_b = st.number_input("Nombre d'éprouvettes", min_value=0, max_value=12, value=6)

        submit_b = st.form_submit_button("💾 Enregistrer dans Supabase", type="primary")
        
        if submit_b:
            nouvel_essai = {
                "date_livraison": date_b.strftime("%Y-%m-%d"),
                "bl_numero": bl_num,
                "classe_beton": classe_b,
                "element_ouvrage": element,
                "volume": float(volume_b),
                "temperature": float(temp_b),
                "slump": float(slump_b),
                "prelevement": prelev_b,
                "nb_eprouvettes": int(nb_ep_b)
            }
            try:
                supabase.table("suivi_betonnage").insert(nouvel_essai).execute()
                st.success("✅ Bétonnage enregistré avec succès dans la base de données !")
                st.rerun()
            except Exception as e:
                try:
                    # Tente d'insérer dans la table alternative 'beton'
                    supabase.table("beton").insert(nouvel_essai).execute()
                    st.success("✅ Bétonnage enregistré avec succès !")
                    st.rerun()
                except Exception as e2:
                    st.error(f"Erreur d'enregistrement Supabase : {e2}")

    # --- D. Historique et Données enregistrées ---
    st.markdown("---")
    st.subheader("📋 Historique complet des bétonnages")
    
    if not df_beton.empty:
        # Nettoyage de l'affichage des colonnes
        cols_to_show = [c for c in ["date_livraison", "bl_numero", "classe_beton", "element_ouvrage", "volume", "slump", "prelevement", "nb_eprouvettes"] if c in df_beton.columns]
        df_display = df_beton[cols_to_show] if cols_to_show else df_beton
        df_display.index = range(1, len(df_display) + 1)
        st.dataframe(df_display, use_container_width=True)
    else:
        st.info("Aucune donnée enregistrée pour le moment dans Supabase.")

# ------------------------------------------------------------------------------
# PAGE 3 : ESSAI À LA PLAQUE
# ------------------------------------------------------------------------------
elif page == "🪨 Essai à la Plaque":
    st.title("🪨 Contrôle de Portance - Essai à la Plaque (NF P94-117-1)")
    
    try:
        resp = supabase.table("essais_plaque").select("*").execute()
        data_all_plaque = resp.data or []
    except Exception:
        data_all_plaque = []

    date_choisie_p = st.date_input("📅 Date de l'essai :", value=date.today())
    str_date_p = date_choisie_p.strftime("%d/%m/%Y")

    with st.form("form_plaque"):
        st.markdown(f"### 📝 Saisie Essai à la Plaque ({str_date_p})")
        
        col_proj1, col_proj2 = st.columns(2)
        with col_proj1:
            projet = st.text_input("Projet", value="LGV CASA SUD", disabled=True)
        with col_proj2:
            client = st.text_input("Entreprise / Client", value="TGCC", disabled=True)
            
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            pk_emp = st.text_input("Emplacement / PK", value="PK 14+250 - Voie 1")
            couche_elem = st.selectbox("Couche / Support", ["PFT3 (Couche de Forme)", "PST (Arase)", "Couche d'Assise", "Remblai"])
            
        with c2:
            z1 = st.number_input("Enfoncement 1er chargement Z1 (mm)", value=1.50, step=0.01, min_value=0.01)
            z2 = st.number_input("Enfoncement 2ème chargement Z2 (mm)", value=0.50, step=0.01, min_value=0.01)
            
        with c3:
            ev1 = round(112.5 / (z1 * 2.0), 2) if z1 > 0 else 0.0
            ev2 = round(90.0 / (z2 * 2.0), 2) if z2 > 0 else 0.0
            rapport_calc = round(ev2 / ev1, 2) if ev1 > 0 else 0.0
            
            st.metric("EV1 (MPa)", value=ev1)
            st.metric("EV2 (MPa)", value=ev2)
            st.metric("Rapport k = EV2 / EV1", value=rapport_calc)
            
            is_conforme = (ev2 >= 50.0) and (rapport_calc <= 2.2)
            statut_auto = "✅ Conforme" if is_conforme else "⚠️ Non Conforme"
            st.info(f"Statut : **{statut_auto}**")

        obs_p = st.text_area("Observations", value="RAS - Sol bien compacté")
        
        submitted = st.form_submit_button("💾 Enregistrer l'essai à la plaque", type="primary")
        if submitted:
            row_p = {
                "date_essai": str_date_p,
                "projet": projet,
                "client": client,
                "pk_emplacement": pk_emp,
                "couche_element": couche_elem,
                "z1": float(z1),
                "z2": float(z2),
                "ev1": float(ev1),
                "ev2": float(ev2),
                "rapport_ev2_ev1": float(rapport_calc),
                "statut": statut_auto,
                "observations": obs_p
            }
            try:
                supabase.table("essais_plaque").insert(row_p).execute()
                st.success("✅ Essai enregistré avec succès !")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur d'enregistrement : {e}")

    st.markdown("---")
    st.subheader("📋 Historique des Essais à la Plaque")
    if data_all_plaque:
        df_p = pd.DataFrame(data_all_plaque)
        df_p.index = range(1, len(df_p) + 1)
        st.dataframe(df_p, use_container_width=True)
    else:
        st.info("Aucun essai enregistré.")
