import os
import streamlit as st

st.title("🔍 Diagnostic des fichiers du serveur")

st.write("📂 **Dossier racine :**", os.listdir("."))

if os.path.exists("pages"):
    st.write("📁 **Contenu du dossier `pages` :**", os.listdir("pages"))

if os.path.exists("views"):
    st.write("📁 **Contenu du dossier `views` :**", os.listdir("views"))
