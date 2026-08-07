import streamlit as st
import pandas as pd
from datetime import datetime, date
from supabase import create_client, Client

# 1. Configuration de la page
st.set_page_config(page_title="Suivi Béton - LGV Casa Sud (LPEE)", layout="wide")

# =========================================================
# 2. VALEURS FIXES & MOTS DE PASSE
# =========================================================
DEFAULT_PROJET = "LGV Casa Sud"
DEFAULT_ENTREPRISE = "TGCC"
DEFAULT_CENTRALE = "TG PREFA"

PASSWORD_GENERAL = "lpee2026"          # Accès technicien
PASSWORD_ADMIN = "lpee_admin_2026"     # Code administrateur (Pour Modifier & Supprimer)

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
st.markdown("### Laboratoire Public d'Essais et d'Études (LPEE)")

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
# 5. GESTION PAR ONGLETS (JOURNÉE ACTUELLE)
# =========================================================
tab_ajouter, tab_modifier, tab_supprimer, tab_historique = st.tabs([
    "➕ Ajouter un camion (Saisie du Jour)", 
    "✏️ Modifier un camion (Admin)", 
    "❌ Supprimer un camion (Admin)",
    "📚 Historique & Tous les Fichiers"
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
            entreprise = st.text_input("Entreprise", value=DEFAULT_ENTREPRISE, disabled=True)
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
            heure_fin_prod = st.text_input("Heure fin prod. CAB", value="15:28", key="add_h_fin")
            heure_arrivee = st.text_input("Heure arrivée chantier", value="16:31", key="add_h_arr")
        with col_m2:
            tbf = st.number_input("TBF (°C)", value=32.0, step=0.1, format="%.1f", key="add_tbf")
            ta = st.number_input("TA (°C) - Ambiante", value=28.9, step=0.1, format="%.1f", key="add_ta")
        with col_m3:
            affaissement = st.number_input("Affaissement (mm)", value=170, step=1, format="%d", key="add_aff")
            meteo = st.selectbox("Météo", ["Soleil", "Nuageux", "Pluie", "Vent"], key="add_meteo")
        with col_m4:
            prelevement = st.selectbox("Prélèvement", ["OUI", "NON"], key="add_prelev")
            statut = st.selectbox("STATUT", ["✅ Conforme", "⚠️ Non Conforme"], key="add_statut")

        observations = st.text_area("Observations", value="RAS", key="add_obs")

        submit_add = st.form_submit_button("💾 Enregistrer le camion dans la journée")
        
        if submit_add:
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
                "statut": statut
            }
            try:
                supabase.table("controles_beton").insert(data_to_insert).execute()
                st.success(f"Camion BL : {num_bon_livraison} enregistré dans la fiche du {str_date_choisie} !")
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
                    edit_entreprise = st.text_input("Entreprise", value=str(row_selected.get('entreprise') or DEFAULT_ENTREPRISE))
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
                    edit_h_fin = st.text_input("Heure fin prod. CAB", value=str(row_selected.get('heure_fin_production_cab') or ''))
                    edit_h_arr = st.text_input("Heure arrivée chantier", value=str(row_selected.get('heure_arrivee_chantier') or ''))
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
                    statut_opts = ["✅ Conforme", "⚠️ Non Conforme"]
                    s_idx = statut_opts.index(row_selected.get('statut')) if row_selected.get('statut') in statut_opts else 0
                    edit_statut = st.selectbox("STATUT", statut_opts, index=s_idx, key="edit_statut_select")

                edit_obs = st.text_area("Observations", value=str(row_selected.get('observations') or ''))

                submit_update = st.form_submit_button("🔄 Mettre à jour cette entrée")

                if submit_update:
                    update_data = {
                        "projet": edit_projet, "entreprise": edit_entreprise, "centrale_beton": edit_centrale,
                        "ouvrage": edit_ouvrage, "element_betonne": edit_elem,
                        "volume_beton": edit_vol, "num_bon_livraison": edit_bl,
                        "classe_beton": edit_classe, "date_betonnage": str_date_choisie, "meteo": edit_meteo,
                        "observations": edit_obs, "heure_fin_production_cab": edit_h_fin,
                        "heure_arrivee_chantier": edit_h_arr, "tbf": round(edit_tbf, 1), "ta": round(edit_ta, 1),
                        "affaissement": int(edit_aff), "prelevement": edit_prelev, "statut": edit_statut
                    }
                    try:
                        supabase.table("controles_beton").update(update_data).eq("id", row_selected["id"]).execute()
                        st.success("Mise à jour réussie !")
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
            "heure_arrivee_chantier", "tbf", "ta", "affaissement", "prelevement", "statut"
        ]
        df_hist_vis = df_all[[c for c in colonnes_historique if c in df_all.columns]]
        # 🔹 Début de l'indexation à 1 au lieu de 0
        df_hist_vis.index = range(1, len(df_hist_vis) + 1)
        st.dataframe(df_hist_vis, use_container_width=True)
        
        csv_all = df_hist_vis.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Télécharger l'historique complet (Tout le chantier)",
            data=csv_all,
            file_name="registre_complet_beton_lgv.csv",
            mime="text/csv",
        )

# =========================================================
# 6. TABLEAU & FICHIER DE LA JOURNÉE SÉLECTIONNÉE
# =========================================================
st.markdown("---")
st.subheader(f"📋 Fichier et Tableau de Suivi pour le : {str_date_choisie}")

if data_jour and len(data_jour) > 0:
    df_jour = pd.DataFrame(data_jour)
    colonnes_visibles = [
        "num_bon_livraison", "ouvrage", "element_betonne", 
        "volume_beton", "classe_beton", "heure_fin_production_cab", 
        "heure_arrivee_chantier", "tbf", "ta", "affaissement", "prelevement", "statut"
    ]
    df_affichage = df_jour[[c for c in colonnes_visibles if c in df_jour.columns]]
    
    # 🔹 Début de l'indexation à 1 au lieu de 0
    df_affichage.index = range(1, len(df_affichage) + 1)
    
    st.dataframe(df_affichage, use_container_width=True)

    # Bouton de téléchargement du fichier SPÉCIFIQUE de la journée
    csv_jour = df_affichage.to_csv(index=False).encode('utf-8')
    nom_fichier_jour = f"Rapport_Betonnage_{date_choisie.strftime('%Y-%m-%d')}.csv"
    
    st.download_button(
        label=f"📥 Télécharger le fichier journalier ({nom_fichier_jour})",
        data=csv_jour,
        file_name=nom_fichier_jour,
        mime="text/csv",
    )
else:
    st.warning(f"Aucun suivi enregistré pour la journée du {str_date_choisie}. Utilisez le formulaire ci-dessus pour ajouter le premier camion !")
