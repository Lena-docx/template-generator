import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page")

# Initialisation de la base de données en mémoire de session
if "df_revues" not in st.session_state:
    st.session_state.df_revues = None

def generer_document_word(nom_revue, données_instructions):
    doc = Document()
    doc.add_heading(f"Instructions de mise en page — {nom_revue}", level=1)
    
    for section, contenu in données_instructions.items():
        if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
            doc.add_heading(section, level=2)
            doc.add_paragraph(str(contenu).strip())
            
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Création des deux points d'entrée
tab_editeurs, tab_compositeurs = st.tabs(["✍️ Éditeurs", "🎼 Compositeurs"])

# ==========================================
# 1. POINT D'ENTRÉE : ÉDITEURS
# ==========================================
with tab_editeurs:
    st.header("Espace Éditeurs")
    st.write("Chargez ici le fichier Excel contenant les instructions de mise en page de vos revues.")
    
    fichier_charge = st.file_uploader(
        "Déposer le fichier Excel (.xlsx)", 
        type=["xlsx"],
        help="Le fichier doit contenir une colonne 'Revue' qui servira d'identifiant."
    )
    
    if fichier_charge is not None:
        try:
            # Lecture du fichier Excel chargé
            df = pd.read_excel(fichier_charge)
            
            if "Revue" in df.columns:
                # Sauvegarde dans la session Streamlit
                st.session_state.df_revues = df.set_index("Revue")
                st.success("✅ Fichier chargé avec succès ! Les données sont prêtes pour les compositeurs.")
            else:
                st.error("⚠️ Erreur : Le fichier Excel doit obligatoirement contenir une colonne nommée exactement 'Revue'.")
                
        except Exception as e:
            st.error(f"⚠️ Erreur lors de la lecture du fichier : {e}")

# ==========================================
# 2. POINT D'ENTRÉE : COMPOSITEURS
# ==========================================
with tab_compositeurs:
    st.header("Espace Compositeurs")
    
    # Vérification si des données ont été chargées par les éditeurs
    if st.session_state.df_revues is None:
        st.info("ℹ️ En attente du chargement des données. Veuillez demander à un Éditeur de déposer le fichier Excel dans l'onglet dédié.")
    else:
        df_revues = st.session_state.df_revues
        
        choix = st.selectbox(
            "Choisir ou chercher une revue :", 
            ["-- Sélectionnez une revue --"] + list(df_revues.index),
            help="Cliquez et tapez les premières lettres pour chercher une revue."
        )

        if choix != "-- Sélectionnez une revue --":
            instructions_revue = df_revues.loc[choix]
            
            # Génération du fichier Word
            fichier_word = generer_document_word(choix, instructions_revue)
            
            # Bouton de téléchargement
            st.download_button(
                label="📄 Télécharger au format Word (.docx)",
                data=fichier_word,
                file_name=f"Instructions_Mise_En_Page_{choix}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.markdown("---")
            
            # Affichage des sections (Utilisation de st.write pour une meilleure lisibilité)
            for section, contenu in instructions_revue.items():
                if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
                    st.subheader(section)
                    st.write(str(contenu).strip())
