import streamlit as st
import pandas as pd
from supabase import create_client, Client

# 1. Configuration de la page
st.set_page_config(page_title="Suivi Béton - LGV Casa Sud (LPEE)", layout="wide")

# =========================================================
# 2. MOTS DE PASSE (GÉNÉRAL ET ADMINISTRATEUR)
# =========================================================
PASSWORD_GENERAL = "lpee2026"          # Mot de passe pour les techniciens (Saisie & Consultation)
PASSWORD_ADMIN = "lpee@2026"     # ⚠️ VOTRE CODE SECRET (Pour Modifier & Supprimer)

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
    data = response.data
except Exception as e:
    st.error(f"Erreur de connexion à la base de données : {e}")
    data = []

# En-tête de l'application
st.title("🚧 Fiche de Suivi et Contrôle du Béton - LGV Casa Sud")
st.markdown("### Laboratoire Public d'Essais et d'Études (LPEE)")

# =========================================================
# 4. GESTION PAR ONGLETS (AJOUTER / MODIFIER / SUPPRIMER)
# =========================================================
tab_ajouter, tab_modifier, tab_supprimer = st.tabs([
    "➕ Nouveau rapport", 
    "✏️ Modifier un rapport (Admin)", 
    "❌ Supprimer un rapport (Admin)"
])

# ---------------------------------------------------------
# --- ONGLET 1 : NOUVEAU RAPPORT (Accessible à tous) ---
# ---------------------------------------------------------
with tab_ajouter:
    with st.form("form_controle_ajouter"):
        st.markdown("#### 1. Identification, Traçabilité & Conditions Chantier")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            num_betonnage = st.text_input("N° Bétonnage", value="25/260/IA/01", key="add_num_bet")
            projet = st.text_input("Projet", value="LGV Casa Sud", key="add_projet")
            ouvrage = st.text_input("Ouvrage", value="PRO745 OA1", key="add_ouvrage")
            element_betonne = st.text_input("Élément bétonné", value="Semelle C0", key="add_elem")
            entreprise = st.text_input("Entreprise", value="TGCC", key="add_ent")
            volume_beton = st.text_input("Volume béton", value="120 m³", key="add_vol")
            
        with c2:
            centrale_beton = st.text_input("Centrale béton", value="Centrale X", key="add_centrale")
            heure_malaxage = st.text_input("Heure malaxage", value="08:30", key="add_h_mal")
            num_bon_livraison = st.text_input("N° bon livraison", value="BL2548", key="add_bl")
            camion_toupie = st.text_input("Camion toupie", value="T12", key="add_toupie")
            classe_beton = st.text_input("Classe béton", value="C30/37", key="add_classe")
            date_betonnage = st.text_input("Date bétonnage", value="05/08/2026", key="add_date")
            
        with c3:
            meteo = st.selectbox("Météo", ["Soleil", "Nuageux", "Pluie", "Vent"], key="add_meteo")
            observations = st.text_area("Observations", value="RAS", key="add_obs")

        st.markdown("---")
        st.markdown("#### 3. Contrôle du Béton Frais (Mesures Chantier)")
        
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            heure_fin_prod = st.text_input("Heure fin prod. CAB", value="15:28", key="add_h_fin")
            heure_arrivee = st.text_input("Heure arrivée chantier", value="16:31", key="add_h_arr")
        with col_m2:
            tbf = st.number_input("TBF (°C)", value=32.0, key="add_tbf")
            ta = st.number_input("TA (°C) - Ambiante", value=28.9, key="add_ta")
        with col_m3:
            affaissement = st.number_input("Affaissement (mm)", value=170.0, key="add_aff")
            prelevement = st.selectbox("Prélèvement", ["OUI", "NON"], key="add_prelev")
        with col_m4:
            statut = st.selectbox("STATUT", ["✅ Conforme", "⚠️ Non Conforme"], key="add_statut")

        submit_add = st.form_submit_button("💾 Enregistrer le rapport complet")
        
        if submit_add:
            data_to_insert = {
                "num_betonnage": num_betonnage, "projet": projet, "ouvrage": ouvrage,
                "element_betonne": element_betonne, "entreprise": entreprise, "volume_beton": volume_beton,
                "centrale_beton": centrale_beton, "heure_malaxage": heure_malaxage,
                "num_bon_livraison": num_bon_livraison, "camion_toupie": camion_toupie,
                "classe_beton": classe_beton, "date_betonnage": date_betonnage,
                "meteo": meteo, "observations": observations,
                "heure_fin_production_cab": heure_fin_prod, "heure_arrivee_chantier": heure_arrivee,
                "tbf": tbf, "ta": ta, "affaissement": affaissement,
                "prelevement": prelevement, "statut": statut
            }
            try:
                supabase.table("controles_beton").insert(data_to_insert).execute()
                st.success("Rapport enregistré avec succès dans Supabase !")
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
        if data and len(data) > 0:
            options_edit = {
                f"ID #{row.get('id')} | N° Bétonnage : {row.get('num_betonnage')} | BL : {row.get('num_bon_livraison')} | Ouvrage : {row.get('ouvrage')}": row
                for row in data
            }
            choix = st.selectbox("📌 Choisissez la fiche à modifier :", list(options_edit.keys()), key="select_edit")
            row_selected = options_edit[choix]

            with st.form("form_controle_modifier"):
                st.markdown(f"#### Modification de la fiche ID #{row_selected.get('id')}")
                c1, c2, c3 = st.columns(3)
                
                with c1:
                    edit_num_bet = st.text_input("N° Bétonnage", value=str(row_selected.get('num_betonnage') or ''))
                    edit_projet = st.text_input("Projet", value=str(row_selected.get('projet') or ''))
                    edit_ouvrage = st.text_input("Ouvrage", value=str(row_selected.get('ouvrage') or ''))
                    edit_elem = st.text_input("Élément bétonné", value=str(row_selected.get('element_betonne') or ''))
                    edit_entreprise = st.text_input("Entreprise", value=str(row_selected.get('entreprise') or ''))
                    edit_vol = st.text_input("Volume béton", value=str(row_selected.get('volume_beton') or ''))
                    
                with c2:
                    edit_centrale = st.text_input("Centrale béton", value=str(row_selected.get('centrale_beton') or ''))
                    edit_heure_mal = st.text_input("Heure malaxage", value=str(row_selected.get('heure_malaxage') or ''))
                    edit_bl = st.text_input("N° bon livraison", value=str(row_selected.get('num_bon_livraison') or ''))
                    edit_toupie = st.text_input("Camion toupie", value=str(row_selected.get('camion_toupie') or ''))
                    edit_classe = st.text_input("Classe béton", value=str(row_selected.get('classe_beton') or ''))
                    edit_date = st.text_input("Date bétonnage", value=str(row_selected.get('date_betonnage') or ''))
                    
                with c3:
                    meteo_opts = ["Soleil", "Nuageux", "Pluie", "Vent"]
                    m_idx = meteo_opts.index(row_selected.get('meteo')) if row_selected.get('meteo') in meteo_opts else 0
                    edit_meteo = st.selectbox("Météo", meteo_opts, index=m_idx, key="edit_meteo_select")
                    edit_obs = st.text_area("Observations", value=str(row_selected.get('observations') or ''))

                st.markdown("---")
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                with col_m1:
                    edit_h_fin = st.text_input("Heure fin prod. CAB", value=str(row_selected.get('heure_fin_production_cab') or ''))
                    edit_h_arr = st.text_input("Heure arrivée chantier", value=str(row_selected.get('heure_arrivee_chantier') or ''))
                with col_m2:
                    edit_tbf = st.number_input("TBF (°C)", value=float(row_selected.get('tbf') or 0.0), key="edit_tbf_input")
                    edit_ta = st.number_input("TA (°C)", value=float(row_selected.get('ta') or 0.0), key="edit_ta_input")
                with col_m3:
                    edit_aff = st.number_input("Affaissement (mm)", value=float(row_selected.get('affaissement') or 0.0), key="edit_aff_input")
                    prelev_opts = ["OUI", "NON"]
                    p_idx = prelev_opts.index(row_selected.get('prelevement')) if row_selected.get('prelevement') in prelev_opts else 0
                    edit_prelev = st.selectbox("Prélèvement", prelev_opts, index=p_idx, key="edit_prelev_select")
                with col_m4:
                    statut_opts = ["✅ Conforme", "⚠️ Non Conforme"]
                    s_idx = statut_opts.index(row_selected.get('statut')) if row_selected.get('statut') in statut_opts else 0
                    edit_statut = st.selectbox("STATUT", statut_opts, index=s_idx, key="edit_statut_select")

                submit_update = st.form_submit_button("🔄 Mettre à jour cette fiche")

                if submit_update:
                    update_data = {
                        "num_betonnage": edit_num_bet, "projet": edit_projet, "ouvrage": edit_ouvrage,
                        "element_betonne": edit_elem, "entreprise": edit_entreprise, "volume_beton": edit_vol,
                        "centrale_beton": edit_centrale, "heure_malaxage": edit_heure_mal,
                        "num_bon_livraison": edit_bl, "camion_toupie": edit_toupie,
                        "classe_beton": edit_classe, "date_betonnage": edit_date,
                        "meteo": edit_meteo, "observations": edit_obs,
                        "heure_fin_production_cab": edit_h_fin, "heure_arrivee_chantier": edit_h_arr,
                        "tbf": edit_tbf, "ta": edit_ta, "affaissement": edit_aff,
                        "prelevement": edit_prelev, "statut": edit_statut
                    }
                    try:
                        supabase.table("controles_beton").update(update_data).eq("id", row_selected["id"]).execute()
                        st.success("Fiche mise à jour avec succès !")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur lors de la mise à jour : {e}")
        else:
            st.info("Aucun contrôle enregistré dans la base de données.")

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
        if data and len(data) > 0:
            options_del = {
                f"ID #{row.get('id')} | N° Bétonnage : {row.get('num_betonnage')} | BL : {row.get('num_bon_livraison')} | Ouvrage : {row.get('ouvrage')}": row["id"]
                for row in data
            }
            choix_del = st.selectbox("⚠️ Choisissez la fiche à supprimer définitivement :", list(options_del.keys()), key="select_del")
            id_to_delete = options_del[choix_del]

            if st.button("🚨 Confirmer la suppression définitive", type="primary"):
                try:
                    supabase.table("controles_beton").delete().eq("id", id_to_delete).execute()
                    st.success("Fiche supprimée avec succès !")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur lors de la suppression : {e}")
        else:
            st.info("Aucun contrôle enregistré à supprimer.")

# =========================================================
# 5. REGISTRE ET TABLEAU D'AFFICHAGE GÉNÉRAL
# =========================================================
st.markdown("---")
st.subheader("📋 Registre Général des Contrôles de Béton Frais")

if data and len(data) > 0:
    df = pd.DataFrame(data)
    colonnes_visibles = [
        "id", "num_betonnage", "ouvrage", "element_betonne", "centrale_beton", 
        "num_bon_livraison", "camion_toupie", "classe_beton", 
        "heure_fin_production_cab", "heure_arrivee_chantier", "tbf", "ta", "affaissement", "prelevement", "statut"
    ]
    df_affichage = df[[c for c in colonnes_visibles if c in df.columns]]
    st.dataframe(df_affichage, use_container_width=True)

    csv = df_affichage.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Télécharger le registre officiel (Format CSV / Excel)",
        data=csv,
        file_name="registre_suivi_beton_lgv_casasud.csv",
        mime="text/csv",
    )
else:
    st.info("Aucun contrôle enregistré dans la base de données pour le moment.")
