import streamlit as st
import pandas as pd

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page pour le compositeur")

@st.cache_data
def charger_donnees_excel():
    # Lecture directe du fichier Excel d'origine
    df = pd.read_excel("revues.xlsx")
    # On définit la première colonne 'Revue' comme index
    return df.set_index("Revue")

try:
    df_revues = charger_donnees_excel()
    
    # Menu déroulant généré à partir des lignes du fichier Excel
    choix = st.selectbox("Choisir une revue :", ["-- Sélectionnez une revue --"] + list(df_revues.index))

    if choix != "-- Sélectionnez une revue --":
        instructions_revue = df_revues.loc[choix]
        
        st.markdown("---")
        # Parcours automatique de toutes les colonnes de la feuille Excel
        for section, contenu in instructions_revue.items():
            # Affichage si la cellule n'est pas vide et ne contient pas un simple "/"
            if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
                st.subheader(section)
                # Utilisation de st.text() pour préserver les retours à la ligne d'Excel
                st.text(str(contenu).strip())

except FileNotFoundError:
    st.error("⚠️ Le fichier 'revues.xlsx' est introuvable. Veuillez le nommer ainsi et le placer au même niveau que 'app.py' sur GitHub.")
