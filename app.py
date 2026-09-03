import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import sqlite3

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page")

MOT_DE_PASSE_EDITEUR = "Editeur2026"  # 🔐 Modifiez ce mot de passe selon vos besoins
DB_NOM = "revues.db"

# CONNEXION ET CRÉATION AUTOMATIQUE DE LA BASE DE DONNÉES LOCALES
def initialiser_sqlite():
    """Crée la table SQLite locale au premier démarrage si elle n'existe pas"""
    conn = sqlite3.connect(DB_NOM)
    cursor = conn.cursor()
    # On crée une structure flexible qui stockera la revue et ses sections
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS instructions (
        revue TEXT PRIMARY KEY,
        donnees_json TEXT
    );
    """)
    conn.commit()
    conn.close()

initialiser_sqlite()

def charger_donnees_sqlite():
    """Lit les données de SQLite et reconstruit un DataFrame Pandas propre"""
    conn = sqlite3.connect(DB_NOM)
    try:
        # Récupération de toutes les lignes
        df_sql = pd.read_sql_query("SELECT * FROM instructions", conn)
        conn.close()
        
        if df_sql.empty:
            return None
            
        # Reconstitution du dictionnaire JSON en colonnes pour Pandas
        import json
        liste_dictionnaires = []
        for _, row in df_sql.iterrows():
            dict_revue = json.loads(row["donnees_json"])
            dict_revue["Revue"] = row["revue"]
            liste_dictionnaires.append(dict_revue)
            
        df_final = pd.DataFrame(liste_dictionnaires)
        return df_final.set_index("Revue")
    except Exception:
        conn.close()
        return None

def sauvegarder_revue_sqlite(nom_revue, dictionnaire_sections):
    """Sauvegarde ou met à jour une revue et ses consignes en format JSON structuré"""
    import json
    conn = sqlite3.connect(DB_NOM)
    cursor = conn.cursor()
    json_consignes = json.dumps(dictionnaire_sections)
    
    cursor.execute("""
        INSERT INTO instructions (revue, donnees_json) 
        VALUES (?, ?)
        ON CONFLICT(revue) DO UPDATE SET donnees_json = excluded.donnees_json;
    """, (nom_revue, json_consignes))
    
    conn.commit()
    conn.close()

def supprimer_revue_sqlite(nom_revue):
    """Supprime définitivement une revue de la base de données"""
    conn = sqlite3.connect(DB_NOM)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM instructions WHERE revue = ?;", (nom_revue,))
    conn.commit()
    conn.close()

def generer_document_word(nom_revue, données_instructions):
    doc = Document()
    doc.add_heading(f"Instructions de mise en page — {nom_revue}", level=1)
    
    for section, contenu in données_instructions.items():
        if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
            doc.add_heading(str(section).capitalize(), level=2)
            doc.add_paragraph(str(contenu).strip())
            
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Chargement immédiat des données SQLite
st.session_state.df_revues = charger_donnees_sqlite()

# Création des deux points d'entrée
tab_editeurs, tab_compositeurs = st.tabs(["✍️ Éditeurs", "🎼 Compositeurs"])

# ==========================================
# 1. POINT D'ENTRÉE : ÉDITEURS
# ==========================================
with tab_editeurs:
    st.header("Espace Éditeurs")
    
    if "authentifie" not in st.session_state:
        st.session_state.authentifie = False
        
    if not st.session_state.authentifie:
        with st.form("form_auth"):
            mdp_saisi = st.text_input("Veuillez saisir le mot de passe pour accéder à cet espace :", type="password")
            valider_auth = st.form_submit_button("Se connecter")
            if valider_auth:
                if mdp_saisi == MOT_DE_PASSE_EDITEUR:
                    st.session_state.authentifie = True
                    st.rerun()
                else:
                    st.error("🔑 Mot de passe incorrect.")

    else:
        # Bouton de déconnexion discret si l'utilisateur est authentifié
        if st.button("🔒 Se déconnecter de l'espace Éditeur"):
            st.session_state.authentifie = False
            st.rerun()
            
        st.markdown("---")

        if st.session_state.df_revues is not None:
            df_edition = st.session_state.df_revues
            liste_sections = list(df_edition.columns)
            
            # --- SECTION 1 : MODIFIER ET SAUVEGARDER EN DIRECT ---
            st.subheader("1. Modifier et Sauvegarder en direct")
            
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
                        # Préparation du dictionnaire nettoyé des consignes de cette revue
                        dict_sauvegarde = {}
                        for section, nv_texte in nouveaux_contenus.items():
                            dict_sauvegarde[section] = nv_texte if nv_texte.strip() != "" else "/"
                        
                        # Sauvegarde immédiate dans la base SQLite locale de Streamlit
                        sauvegarder_revue_sqlite(revue_a_modifier, dict_sauvegarde)
                        st.toast(f"Base locale mise à jour pour {revue_a_modifier} !", icon="💾")
                        st.rerun()

            # CAS 2 : MODIFICATION GLOBALE (Écraser une section pour toutes les revues)
            else:
                section_a_modifier = st.selectbox("Sélectionner la section à harmoniser partout :", liste_sections)
                premiere_revue_nom = df_edition.index if len(df_edition.index) > 0 else "Aucune"
                valeur_premiere_revue = str(df_edition.iloc[section_a_modifier]) if len(df_edition.index) > 0 else ""
                
                if valeur_premiere_revue == "nan" or valeur_premiere_revue == "/":
                    valeur_premiere_revue = ""
                
                with st.form("form_global_revue"):
                    st.write(f"🚨 Vous allez écraser la section **{section_a_modifier}** pour **toutes** les revues.")
                    st.caption(f"💡 Champ pré-rempli avec le texte actuel de la première revue : *{premiere_revue_nom}*.")
                    
                    texte_global = st.text_area("Nouveau texte commun à appliquer partout :", value=valeur_premiere_revue)
                    soumettre_global = st.form_submit_button("⚠️ Écraser et Sauvegarder sur tout le catalogue")
                    
                    if soumettre_global:
                        texte_global_propre = texte_global if texte_global.strip() != "" else "/"
                        
                        # On applique la modification globale sur chaque revue présente dans la base
                        for nom_revue in df_edition.index:
                            dict_actuel = df_edition.loc[nom_revue].to_dict()
                            dict_actuel[section_a_modifier] = texte_global_propre
                            sauvegarder_revue_sqlite(nom_revue, dict_actuel)
                            
                        st.toast("Mise à jour globale enregistrée localement !", icon="💾")
                        st.rerun()

            # --- SECTION 2 : STRUCTURER LA BASE DE DONNÉES (AJOUT & SUPPRESSION) ---
            st.markdown("---")
            st.subheader("2. Structurer la base de données")
            
            # Bloc A : AJOUTS (Côte à côte)
            st.markdown("##### ➕ Ajouts")
            col_revue, col_section = st.columns(2)
            
            with col_revue:
                with st.form("form_ajouter_revue"):
                    st.write("**Ajouter une nouvelle revue**")
                    nouvelle_revue_nom = st.text_input("Nom de la nouvelle revue :")
                    soumettre_nouvelle_revue = st.form_submit_button("Créer la revue")
                    if soumettre_nouvelle_revue and nouvelle_revue_nom.strip() != "":
                        nom_propre = nouvelle_revue_nom.strip()
                        if nom_propre in st.session_state.df_revues.index:
                            st.error("⚠️ Cette revue existe déjà.")
                        else:
                            # Création d'une revue avec des consignes vides
                            dict_vide = {col: "/" for col in liste_sections}
                            sauvegarder_revue_sqlite(nom_propre, dict_vide)
                            st.success(f"Revue '{nom_propre}' créée localement !")
                            st.rerun()
                                
            with col_section:
                with st.form("form_ajouter_section"):
                    st.write("**Ajouter une nouvelle section**")
                    nouvelle_section_nom = st.text_input("Nom de la nouvelle section :")
                    soumettre_nouvelle_section = st.form_submit_button("Créer la section")
                    if soumettre_nouvelle_section and nouvelle_section_nom.strip() != "":
                        sec_propre = nouvelle_section_nom.strip()
                        if sec_propre in st.session_state.df_revues.columns:
                            st.error("⚠️ Cette section existe déjà.")
                        else:
                            # Ajouter une colonne consiste à modifier le JSON de chaque revue
                            for nom_revue in df_edition.index:
                                dict_actuel = df_edition.loc[nom_revue].to_dict()
                                dict_actuel[sec_propre] = "/"
                                sauvegarder_revue_sqlite(nom_revue, dict_actuel)
                            st.success(f"Section '{sec_propre}' ajoutée partout !")
                            st.rerun()

            # Bloc B : SUPPRESSIONS (Irréversibles)
            st.markdown("##### 🗑️ Suppressions (Irréversible)")
            col_del_revue, col_del_section = st.columns(2)
            
            with col_del_revue:
                with st.form("form_supprimer_revue"):
                    st.write("**Supprimer une revue**")
                    revue_a_supprimer = st.selectbox("Revue à détruire :", ["-- Sélectionner --"] + list(df_edition.index))
                    soumettre_del_revue = st.form_submit_button("💥 Supprimer la revue")
                    if soumettre_del_revue and revue_a_supprimer != "-- Sélectionner --":
                        supprimer_revue_sqlite(revue_a_supprimer)
                        st.success(f"La revue '{revue_a_supprimer}' a été supprimée.")
                        st.rerun()
                            
            with col_del_section:
                with st.form("form_supprimer_section"):
                    st.write("**Supprimer une section complète**")
                    section_a_supprimer = st.selectbox("Section à détruire :", ["-- Sélectionner --"] + liste_sections)
                    soumettre_del_section = st.form_submit_button("💥 Supprimer la section")
                    if soumettre_del_section and section_a_supprimer != "-- Sélectionner --":
                        # Pour supprimer une colonne, on la retire du dictionnaire JSON de chaque ligne
                        for nom_revue in df_edition.index:
                            dict_actuel = df_edition.loc[nom_revue].to_dict()
                            if section_a_supprimer in dict_actuel:
                                del dict_actuel[section_a_supprimer]
                            sauvegarder_revue_sqlite(nom_revue, dict_actuel)
                        st.success(f"La section '{section_a_supprimer}' a été retirée pour tout le monde.")
                        st.rerun()

        # --- SECTION 3 : IMPORT INITIAL DE TOUT LE FICHIER EXCEL ---
        st.markdown("---")
        st.subheader("3. Remplissage ou Remplacement global via Excel")
        st.write("Déposez votre fichier Excel d'origine pour pré-remplir instantanément la base locale de l'application.")
        
        fichier_charge = st.file_uploader(
            "Déposer le fichier Excel complet (.xlsx)", 
            type=["xlsx"],
            key="uploader_excel"
        )
        
        if fichier_charge is not None:
            try:
                df_nouveau = pd.read_excel(fichier_charge)
                if "Revue" in df_nouveau.columns:
                    # On vide l'ancienne table SQLite pour repartir sur une base propre
                    conn_clear = sqlite3.connect(DB_NOM)
                    conn_clear.execute("DELETE FROM instructions;")
                    conn_clear.commit()
                    conn_clear.close()
                    
                    # On boucle sur le fichier Excel pour tout enregistrer au format JSON
                    for _, row in df_nouveau.iterrows():
                        nom_revue = str(row["Revue"]).strip()
                        dict_revue = {}
                        for col in df_nouveau.columns:
                            if col != "Revue":
                                dict_revue[str(col).strip()] = str(row[col]).strip() if pd.notna(row[col]) else "/"
                        
                        sauvegarder_revue_sqlite(nom_revue, dict_revue)
                        
                    st.success("✅ Félicitations ! Votre application est entièrement configurée avec vos données Excel d'origine.")
                    st.rerun()
                else:
                    st.error("⚠️ Erreur : Le fichier Excel doit contenir une colonne nommée exactement 'Revue'.")
            except Exception as e:
                st.error(f"⚠️ Erreur de traitement du fichier Excel : {e}")

# ==========================================
# 2. POINT D'ENTRÉE : COMPOSITEURS
# ==========================================
with tab_compositeurs:
    st.header("Espace Compositeurs")
    
    if st.session_state.df_revues is None:
        st.info("ℹ️ L'application est prête. Veuillez vous connecter à l'espace Éditeur pour y charger votre fichier Excel d'origine.")
    else:
        df_revues = st.session_state.df_revues
        
        choix = st.selectbox(
            "Choisir ou chercher une revue :", 
            ["-- Sélectionnez une revue --"] + list(df_revues.index),
            key="select_compositeur"
        )

        if choix != "-- Sélectionnez une revue --":
            instructions_revue = df_revues.loc[choix]
            
            # Génération à la volée du document Word
            fichier_word = generer_document_word(choix, instructions_revue)
            st.download_button(
                label="📄 Télécharger au format Word (.docx)",
                data=fichier_word,
                file_name=f"Instructions_Mise_En_Page_{choix}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.markdown("---")
            
            # Affichage à l'écran
            for section, contenu in instructions_revue.items():
                if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
                    st.subheader(str(section))
                    st.write(str(contenu).strip())

