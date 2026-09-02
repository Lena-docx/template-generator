import streamlit as st
import pandas as pd

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page pour le compositeur")

@st.cache_data
def charger_donnees():
    # Chargement automatique du fichier CSV externe
    df = pd.read_csv("revues.csv")
    # Utilisation du nom de la revue comme index
    return df.set_index("Revue")

try:
    df_revues = charger_donnees()
    
    # Création du menu déroulant à partir de la base de données externe
    choix = st.selectbox("Choisir une revue :", ["-- Sélectionnez une revue --"] + list(df_revues.index))

    if choix != "-- Sélectionnez une revue --":
        # Extraction de la ligne de données correspondant à la revue sélectionnée
        instructions_revue = df_revues.loc[choix]
        
        st.markdown("---")
        # Boucle automatique sur toutes les colonnes/sections du fichier CSV
        for section, contenu in instructions_revue.items():
            # On affiche la section uniquement si elle contient du texte utile (ni vide, ni juste un "/")
            if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
                st.subheader(section)
                st.text(str(contenu).strip())

except FileNotFoundError:
    st.error("⚠️ Le fichier 'revues.csv' est introuvable. Veuillez le placer dans le même dossier que ce script.")
