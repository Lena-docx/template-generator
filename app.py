import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page pour le compositeur")

@st.cache_data
def charger_donnees_excel():
    df = pd.read_excel("revues.xlsx")
    return df.set_index("Revue")

def generer_document_word(nom_revue, données_instructions):
    # Création d'un document Word en mémoire
    doc = Document()
    doc.add_heading(f"Instructions de mise en page — {nom_revue}", level=1)
    
    for section, contenu in données_instructions.items():
        if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
            doc.add_heading(section, level=2)
            doc.add_paragraph(str(contenu).strip())
            
    # Sauvegarde dans un flux de données binaire pour le téléchargement
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

try:
    df_revues = charger_donnees_excel()
    
    # Le menu déroulant de Streamlit permet nativement de chercher en tapant du texte !
    choix = st.selectbox(
        "Choisir ou chercher une revue :", 
        ["-- Sélectionnez une revue --"] + list(df_revues.index),
        help="Cliquez et tapez les premières lettres pour chercher une revue."
    )

    if choix != "-- Sélectionnez une revue --":
        instructions_revue = df_revues.loc[choix]
        
        # Génération du fichier Word correspondant à la revue sélectionnée
        fichier_word = generer_document_word(choix, instructions_revue)
        
        # Bouton de téléchargement placé en haut pour être facilement accessible
        st.download_button(
            label="📄 Télécharger au format Word (.docx)",
            data=fichier_word,
            file_name=f"Instructions_Mise_En_Page_{choix}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        st.markdown("---")
        
        # Affichage des sections sur la page web
        for section, contenu in instructions_revue.items():
            if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
                st.subheader(section)
                st.text(str(contenu).strip())

except FileNotFoundError:
    st.error("⚠️ Le fichier 'revues.xlsx' est introuvable. Veuillez le placer au même niveau que 'app.py' sur GitHub.")
