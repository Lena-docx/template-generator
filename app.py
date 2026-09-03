import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import os

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page")

# Nom du fichier unique utilisé par l'application
FICHIER_EXCEL = "revues.xlsx"

# INITIALISATION DES DONNÉES
if "df_revues" not in st.session_state:
    if os.path.exists(FICHIER_EXCEL):
        try:
            df_init = pd.read_excel(FICHIER_EXCEL)
            if "Revue" in df_init.columns:
                st.session_state.df_revues = df_init.set_index("Revue")
            else:
                st.session_state.df_revues = None
        except Exception:
            st.session_state.df_revues = None
    else:
        st.session_state.df_revues = None

def sauvegarder_sur_disque(df):
    """Sauvegarde le DataFrame directement dans le fichier Excel local du serveur"""
    try:
        # On réinitialise l'index pour réintégrer la colonne 'Revue' dans le fichier Excel
        df.reset_index().to_excel(FICHIER_EXCEL, index=False)
        return True
    except Exception as e:
        st.error(f"❌ Erreur technique lors de l'écriture sur le disque : {e}")
        return False

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
    
    st.subheader("1. Gestion de la base de données")
    
    if st.session_state.df_revues is not None:
        st.success("📊 La base de données 'revues.xlsx' est active et connectée.")
    else:
        st.warning(f"⚠️ Le fichier '{FICHIER_EXCEL}' n'existe pas encore. Vous devez charger un premier fichier pour l'initialiser.")

    # Zone pour charger un NOUVEAU fichier complet (écrase l'ancien sur le disque)
    fichier_charge = st.file_uploader(
        "Remplacer complètement la base de données actuelle par un nouveau fichier Excel", 
        type=["xlsx"],
        key="uploader_excel"
    )
    
    if fichier_charge is not None:
        try:
            df_nouveau = pd.read_excel(fichier_charge)
            if "Revue" in df_nouveau.columns:
                new_df = df_nouveau.set_index("Revue")
                # Sauvegarde immédiate sur le disque dur
                if sauvegarder_sur_disque(new_df):
                    st.session_state.df_revues = new_df
                    st.success("✅ Nouveau fichier enregistré sur le disque et appliqué avec succès !")
                    st.rerun()
            else:
                st.error("⚠️ Erreur : Le fichier doit contenir une colonne nommée 'Revue'.")
        except Exception as e:
            st.error(f"⚠️ Erreur de lecture : {e}")

    # Zone d'édition en direct
    if st.session_state.df_revues is not None:
        st.markdown("---")
        st.subheader("2. Modifier et Sauvegarder en direct")
        
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
                
                soumettre = st.form_submit_button("💾 Enregistrer et appliquer définitivement")
                if soumettre:
                    # Mise à jour en mémoire
                    for section, nv_texte in nouveaux_contenus.items():
                        st.session_state.df_revues.loc[revue_a_modifier, section] = nv_texte if nv_texte.strip() != "" else "/"
                    
                    # Sauvegarde automatique dans le fichier Excel physique
                    if sauvegarder_sur_disque(st.session_state.df_revues):
                        st.toast(f"Fichier Excel mis à jour pour {revue_a_modifier} !", icon="💾")
                        st.rerun()

        # CAS 2 : MODIFICATION GLOBALE
        else:
            section_a_modifier = st.selectbox("Sélectionner la section à harmoniser partout :", liste_sections)
            
            premiere_revue_nom = df_edition.index
            valeur_premiere_revue = str(df_edition.iloc[section_a_modifier])
            
            if valeur_premiere_revue == "nan" or valeur_premiere_revue == "/":
                valeur_premiere_revue = ""
            
            with st.form("form_global_revue"):
                st.write(f"🚨 Vous allez écraser la section **{section_a_modifier}** pour **toutes** les revues.")
                st.caption(f"💡 Champ pré-rempli avec le texte actuel de la première revue : *{premiere_revue_nom}*.")
                
                texte_global = st.text_area(
                    "Nouveau texte commun à appliquer partout :", 
                    value=valeur_premiere_revue
                )
                
                soumettre_global = st.form_submit_button("⚠️ Écraser et Sauvegarder partout")
                if soumettre_global:
                    # Mise à jour en mémoire
                    st.session_state.df_revues[section_a_modifier] = texte_global if texte_global.strip() != "" else "/"
                    
                    # Sauvegarde automatique dans le fichier Excel physique
                    if sauvegarder_sur_disque(st.session_state.df_revues):
                        st.toast("Fichier Excel global mis à jour !", icon="💾")
                        st.rerun()

# ==========================================
# 2. POINT D'ENTRÉE : COMPOSITEURS
# ==========================================
with tab_compositeurs:
    st.header("Espace Compositeurs")
    
    if st.session_state.df_revues is None:
        st.info("ℹ️ En attente de l'initialisation du fichier 'revues.xlsx' par un Éditeur.")
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
