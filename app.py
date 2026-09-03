import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import os

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page")

# INITIALISATION AUTOMATIQUE DU FICHIER PAR DÉFAUT
# On ne charge le fichier automatique QUE si la session n'a pas encore de données
if "df_revues" not in st.session_state:
    fichier_par_defaut = "revues.xlsx"
    if os.path.exists(fichier_par_defaut):
        try:
            df_init = pd.read_excel(fichier_par_defaut)
            if "Revue" in df_init.columns:
                st.session_state.df_revues = df_init.set_index("Revue")
            else:
                st.session_state.df_revues = None
        except Exception:
            st.session_state.df_revues = None
    else:
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

def exporter_excel(df):
    output = BytesIO()
    df.reset_index().to_excel(output, index=False)
    output.seek(0)
    return output

# Création des deux points d'entrée
tab_editeurs, tab_compositeurs = st.tabs(["✍️ Éditeurs", "🎼 Compositeurs"])

# ==========================================
# 1. POINT D'ENTRÉE : ÉDITEURS
# ==========================================
with tab_editeurs:
    st.header("Espace Éditeurs")
    
    st.subheader("1. Base de données actuelle")
    
    # Message d'état pour savoir d'où proviennent les données
    if st.session_state.df_revues is not None:
        st.info("📊 Une base de données est actuellement active (chargée automatiquement ou par dépôt).")
    else:
        st.warning("⚠️ Aucune base de données active. Le fichier 'revues.xlsx' est introuvable à la racine.")

    # Zone de remplacement optionnelle
    fichier_charge = st.file_uploader(
        "Remplacer la base de données par un nouveau fichier Excel (.xlsx)", 
        type=["xlsx"],
        key="uploader_excel"
    )
    
    # Si l'éditeur dépose un NOUVEAU fichier, il écrase l'ancien en mémoire
    if fichier_charge is not None:
        try:
            df_nouveau = pd.read_excel(fichier_charge)
            if "Revue" in df_nouveau.columns:
                st.session_state.df_revues = df_nouveau.set_index("Revue")
                st.success("✅ Nouveau fichier chargé avec succès et appliqué en mémoire !")
            else:
                st.error("⚠️ Erreur : Le fichier doit contenir une colonne nommée 'Revue'.")
        except Exception as e:
            st.error(f"⚠️ Erreur de lecture : {e}")

    # Zone d'édition (accessible uniquement si un fichier existe par défaut ou a été chargé)
    if st.session_state.df_revues is not None:
        st.markdown("---")
        st.subheader("2. Modifier le contenu")
        
        df_edition = st.session_state.df_revues
        liste_sections = list(df_edition.columns)
        
        mode_modification = st.radio(
            "Périmètre de la modification :",
            ["Pour une seule revue", "Pour l'ensemble des revues"],
            horizontal=True
        )
        
        # CAS 1 : MODIFICATION UNIQUE
        if mode_modification == "Pour une seule revue":
            revue_a_modifier = st.selectbox("Sélectionner la revue à éditer :", df_edition.index)
            
            with st.form("form_mono_revue"):
                st.write(f"✍️ Modifications pour la revue : **{revue_a_modifier}**")
                nouveaux_contenus = {}
                
                for section in liste_sections:
                    valeur_actuelle = str(df_edition.loc[revue_a_modifier, section])
                    if valeur_actuelle == "nan" or valeur_actuelle == "/":
                        valeur_actuelle = ""
                    
                    nouveaux_contenus[section] = st.text_area(f"Section : {section}", value=valeur_actuelle)
                
                soumettre = st.form_submit_button("Enregistrer les modifications de cette revue")
                if soumettre:
                    for section, nv_texte in nouveaux_contenus.items():
                        st.session_state.df_revues.loc[revue_a_modifier, section] = nv_texte if nv_texte.strip() != "" else "/"
                    st.success(f"💾 Les modifications pour '{revue_a_modifier}' ont été appliquées en mémoire !")
                    st.rerun()

        # CAS 2 : MODIFICATION GLOBALE
        else:
            section_a_modifier = st.selectbox("Sélectionner la section à harmoniser partout :", liste_sections)
            
            premiere_revue_nom = df_edition.index[0]
            valeur_premiere_revue = str(df_edition.iloc[0][section_a_modifier])
            
            if valeur_premiere_revue == "nan" or valeur_premiere_revue == "/":
                valeur_premiere_revue = ""
            
            with st.form("form_global_revue"):
                st.write(f"🚨 Vous allez écraser la section **{section_a_modifier}** pour **toutes** les revues.")
                st.caption(f"💡 Le champ ci-dessous a été pré-rempli avec le texte actuel de la première revue : *{premiere_revue_nom}*.")
                
                texte_global = st.text_area(
                    "Nouveau texte commun à appliquer partout :", 
                    value=valeur_premiere_revue
                )
                
                soumettre_global = st.form_submit_button("⚠️ Appliquer à TOUTES les revues")
                if soumettre_global:
                    st.session_state.df_revues[section_a_modifier] = texte_global if texte_global.strip() != "" else "/"
                    st.success(f"💾 La section '{section_a_modifier}' a été mise à jour pour l'intégralité des revues !")
                    st.rerun()

        # Étape C : Récupérer le fichier modifié
        st.markdown("---")
        st.subheader("3. Sauvegarder le fichier Excel mis à jour")
        st.write("Pour ne pas perdre vos modifications lors de la fermeture de l'application, téléchargez la nouvelle version du fichier Excel :")
        
        fichier_excel_modifie = exporter_excel(st.session_state.df_revues)
        st.download_button(
            label="🟢 Télécharger le fichier Excel mis à jour (.xlsx)",
            data=fichier_excel_modifie,
            file_name="revues_modifie.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==========================================
# 2. POINT D'ENTRÉE : COMPOSITEURS
# ==========================================
with tab_compositeurs:
    st.header("Espace Compositeurs")
    
    if st.session_state.df_revues is None:
        st.info("ℹ️ En attente du chargement ou de la configuration des données par un Éditeur.")
    else:
        df_revues = st.session_state.df_revues
        
        choix = st.selectbox(
            "Choisir ou chercher une revue :", 
            ["-- Sélectionnez une revue --"] + list(df_revues.index),
            key="select_compositeur"
        )

        if choix != "-- Sélectionnez une revue --":
            instructions_revue = df_revues.loc[choix]
            
            fichier_word = generer_document_word(choix, instructions_revue)
            st.download_button(
                label="📄 Télécharger au format Word (.docx)",
                data=fichier_word,
                file_name=f"Instructions_Mise_En_Page_{choix}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.markdown("---")
            
            for section, contenu in instructions_revue.items():
                if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
                    st.subheader(section)
                    st.write(str(contenu).strip())
