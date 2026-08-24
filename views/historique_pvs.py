# ... (début de votre code)
                renommage_colonnes = {
                    "affaissement": "affaissement_mm",
                    "temperature": "temp_beton_C",
                }

                colonnes_ordre = [
                    "id",
                    "betonnage_id",
                    "ref_controle",
                    "repere_eprouvette",
                    "num_bl",
                    "ouvrage",
                    "classe_beton",
                    "date_coulee",
                    "echeance",
                    "date_ecrasement",
                    "forme",
                    "section",
                    "force_kn",
                    "fc_mpa",
                    "technicien",
                    "observations",
                    "affaissement_mm",
                    "temp_beton_C"
                ]
                
                # Renommer les colonnes selon le dictionnaire défini plus haut
                df_all = df_all.rename(columns=renommage_colonnes)

                # Filtrer pour s'assurer que l'on n'appelle que des colonnes qui existent
                colonnes_presentes = [col for col in colonnes_ordre if col in df_all.columns]
                
                # Ajouter les éventuelles autres colonnes non listées à la fin
                autres_colonnes = [col for col in df_all.columns if col not in colonnes_presentes]
                toutes_colonnes = colonnes_presentes + autres_colonnes
                
                # Affichage du DataFrame complet
                st.dataframe(df_all[toutes_colonnes], use_container_width=True, hide_index=True)
                
                # Bouton pour télécharger la base de données globale
                excel_global = exporter_dataframe_excel(df_all[toutes_colonnes], "Historique_Global")
                st.download_button(
                    label="📊 Télécharger l'historique complet (Excel)",
                    data=excel_global,
                    file_name=f"Historique_Complet_Ecrasements_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_download_historique_global",
                )
        except Exception as e:
            st.error(f"❌ Erreur lors du chargement de l'historique global : {e}")

# ==============================================================================
# FIN DU SCRIPT
# ==============================================================================
