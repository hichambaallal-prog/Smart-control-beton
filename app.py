# Le message d'erreur indique que la colonne 'client' est absente de la table 'essais_plaque'.
# Il faut supprimer 'client' du dictionnaire de données dans le code pour résoudre l'erreur.

# Voici le bloc corrigé de l'insertion dans la page "Essai à la Plaque" :

code_corrigé = """
    if st.button("💾 Enregistrer l'Essai à la Plaque", type="primary", key="btn_save_plaque"):
        row_p = {
            "date_essai": str_date_p,
            "technicien": technicien_p,
            # "client": client_p,  <-- SUPPRIMÉ car la colonne n'existe pas dans la table
            "localisation": localisation,
            "projet": projet_lgv,
            "type_plateforme": type_plateforme,
            "z1": float(z1),
            "z2": float(z2),
            "ev1": float(ev1),
            "ev2": float(ev2),
            "k": float(k_val),
            "observations": obs_p
        }
        try:
            supabase.table("essais_plaque").insert(row_p).execute()
            st.success("✅ Essai à la plaque enregistré avec succès !")
            st.rerun()
"""

# Je vais mettre à jour tout le fichier streamlit pour inclure cette correction.
print("Correction prête.")
