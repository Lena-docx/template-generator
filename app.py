import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
from sqlalchemy import text  # 🔐 Requis pour la sécurité de SQLAlchemy 2.0

# Configuration de la page
st.set_page_config(page_title="Instructions Compositeur", layout="centered")
st.title("Instructions de mise en page")

MOT_DE_PASSE_EDITEUR = "Editeur2026"  # 🔐 Modifiez ce mot de passe selon vos besoins

# CONNEXION NATIVE POSTGRESQL (Utilise le fichier .streamlit/secrets.toml)
try:
    conn = st.connection("postgresql", type="sql")
except Exception as e:
    st.error(f"❌ Impossible de se connecter à Supabase. Vérifiez vos Secrets. Erreur : {e}")
    st.stop()

def charger_donnees_supabase():
    """Récupère toutes les données de Supabase et reconstruit le DataFrame"""
    try:
        df = conn.query("SELECT * FROM instructions_revues;", ttl="0m")
        if df.empty:
            return None
        # On retire l'id technique SQL et on met la Revue en index comme avant
        if "id" in df.columns:
            df = df.drop(columns=["id"])
        return df.set_index("revue")
    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture sur Supabase : {e}")
        return None

def executer_requete_sql(requete, parametres=None):
    """Exécute une commande d'écriture (INSERT, UPDATE, DELETE, ALTER) sécurisée pour SQLAlchemy 2.0"""
    try:
        with conn.session as session:
            # L'expression textuelle est explicitement déclarée avec text() pour éviter les erreurs de compilation
            session.execute(text(requete), parametres)
            session.commit()
        return True
    except Exception as e:
        st.error(f"❌ Erreur de modification SQL : {e}")
        return False

def generer_document_word(nom_revue, données_instructions):
    doc = Document()
    doc.add_heading(f"Instructions de mise en page — {nom_revue}", level=1)
    
    for section, contenu in données_instructions.items():
        # Transformation du nom de colonne SQL en titre lisible (ex: open_access -> Open access)
        titre_propre = section.replace("_", " ").capitalize()
        if pd.notna(contenu) and str(contenu).strip() not in ["", "/"]:
            doc.add_heading(titre_propre, level=2)
            doc.add_paragraph(str(contenu).strip())
            
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Chargement immédiat des données globales depuis le Cloud
df_revues_db = charger_donnees_supabase()
if df_revues_db is not None:
    st.session_state.df_revues = df_revues_db
else:
    st.session_state.df_revues = None

# Création des deux points d'entrée
tab_editeurs, tab_compositeurs = st.tabs(["✍️ Éditeurs", "🎼 Compositeurs"])

# ==========================================
# 1. POINT D'ENTRÉE : ÉDITEURS
# ==========================================
with tab_editeurs:
    st.header("Espace Éditeurs")
    
    # Sécurisation par mot de passe
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
                        
                        # Affichage d'un nom de section plus lisible pour l'humain
                        nom_label = section.replace("_", " ").capitalize()
                        nouveaux_contenus[section] = st.text_area(f"Section : {nom_label}", value=valeur_actuelle)
                    
                    soumettre = st.form_submit_button("💾 Enregistrer et appliquer sur le Cloud Supabase")
                    if soumettre:
                        succes_global = True
                        # Mise à jour colonne par colonne dans PostgreSQL
                        for section, nv_texte in nouveaux_contenus.items():
                            texte_propre = nv_texte if nv_texte.strip() != "" else "/"
                            
                            requete = f"UPDATE instructions_revues SET {section} = :texte WHERE revue = :nom_revue;"
                            if not executer_requete_sql(requete, {"texte": texte_propre, "nom_revue": revue_a_modifier}):
                                succes_global = False
                        
                        if succes_global:
                            st.toast(f"Supabase mis à jour pour {revue_a_modifier} !", icon="☁️")
                            st.rerun()

            # CAS 2 : MODIFICATION GLOBALE (Écraser une section pour toutes les revues)
            else:
                section_a_modifier = st.selectbox("Sélectionner la section à harmoniser partout :", liste_sections)
                premiere_revue_nom = df_edition.index if len(df_edition.index) > 0 else "Aucune"
                valeur_premiere_revue = str(df_edition.iloc[section_a_modifier]) if len(df_edition.index) > 0 else ""
                
                if valeur_premiere_revue == "nan" or valeur_premiere_revue == "/":
                    valeur_premiere_revue = ""
                
                nom_label_global = section_a_modifier.replace("_", " ").capitalize()
                
                with st.form("form_global_revue"):
                    st.write(f"🚨 Vous allez écraser la section **{nom_label_global}** pour **toutes** les revues.")
                    st.caption(f"💡 Champ pré-rempli avec le texte actuel de la première revue : *{premiere_revue_nom}*.")
                    
                    texte_global = st.text_area("Nouveau texte commun à appliquer partout :", value=valeur_premiere_revue)
                    soumettre_global = st.form_submit_button("⚠️ Écraser et Sauvegarder sur tout le Cloud")
                    
                    if soumettre_global:
                        texte_global_propre = texte_global if texte_global.strip() != "" else "/"
                        
                        requete_globale = f"UPDATE instructions_revues SET {section_a_modifier} = :texte;"
                        if executer_requete_sql(requete_globale, {"texte": texte_global_propre}):
                            st.toast("Mise à jour globale réussie sur Supabase !", icon="☁️")
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
                            st.error("⚠️ Cette revue existe déjà dans Supabase.")
                        else:
                            requete_add = "INSERT INTO instructions_revues (revue) VALUES (:nom_revue);"
                            if executer_requete_sql(requete_add, {"nom_revue": nom_propre}):
                                st.success(f"Revue '{nom_propre}' créée sur le Cloud !")
                                st.rerun()
                                
            with col_section:
                with st.form("form_ajouter_section"):
                    st.write("**Ajouter une nouvelle section**")
                    nouvelle_section_nom = st.text_input("Nom de la nouvelle section (ex: 'Format PDF') :")
                    soumettre_nouvelle_section = st.form_submit_button("Créer la section")
                    if soumettre_nouvelle_section and nouvelle_section_nom.strip() != "":
                        sec_sql = nouvelle_section_nom.strip().lower().replace(" ", "_").replace("-", "_")
                        if sec_sql in st.session_state.df_revues.columns:
                            st.error("⚠️ Cette section existe déjà.")
                        else:
                            requete_alter = f"ALTER TABLE instructions_revues ADD COLUMN {sec_sql} TEXT DEFAULT '/';"
                            if executer_requete_sql(requete_alter):
                                st.success(f"Section '{nouvelle_section_nom}' ajoutée à toute la base !")
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
                        requete_del_rev = "DELETE FROM instructions_revues WHERE revue = :nom_revue;"
                        if executer_requete_sql(requete_del_rev, {"nom_revue": revue_a_supprimer}):
                            st.success(f"La revue '{revue_a_supprimer}' a été supprimée de Supabase.")
                            st.rerun()
                            
            with col_del_section:
                with st.form("form_supprimer_section"):
                    st.write("**Supprimer une section complète**")
                    section_a_supprimer = st.selectbox("Section à détruire :", ["-- Sélectionner --"] + liste_sections)
                    soumettre_del_section = st.form_submit_button("💥 Supprimer la section")
                    if soumettre_del_section and section_a_supprimer != "-- Sélectionner --":
                        requete_del_sec = f"ALTER TABLE instructions_revues DROP COLUMN {section_a_supprimer};"
                        if executer_requete_sql(requete_del_sec):
                            st.success(f"La section '{section_a_supprimer}' a été définitivement retirée.")
                            st.rerun()

        # --- SECTION 3 : IMPORT INITIAL DE TOUT LE FICHIER EXCEL ---
        st.markdown("---")
        st.subheader("3. Remplissage initial ou Remplacement de masse")
        st.write("Utilisez cette section pour charger votre fichier Excel initial et remplir Supabase d'un seul coup.")
        
        fichier_charge = st.file_uploader(
            "Déposer le fichier Excel complet pour peupler le Cloud", 
            type=["xlsx"],
            key="uploader_excel"
        )
        
        if fichier_charge is not None:
            try:
                df_nouveau = pd.read_excel(fichier_charge)
                if "Revue" in df_nouveau.columns:
                    # Vidage de la table via l'appel textuel sécurisé
                    executer_requete_sql("TRUNCATE TABLE instructions_revues;")
                    
                    for _, row in df_nouveau.iterrows():
                        nom_revue = str(row["Revue"]).strip()
                        executer_requete_sql("INSERT INTO instructions_revues (revue) VALUES (:nom_revue);", {"nom_revue": nom_revue})
                        
                        for col in df_nouveau.columns:
                            if col != "Revue":
                                col_sql = col.strip().lower().replace(" ", "_").replace("-", "_").replace(",", "_").replace("(", "_").replace(")", "_")
                                valeur = str(row[col]).strip() if pd.notna(row[col]) else "/"
                                
                                try:
                                    executer_requete_sql(f"UPDATE instructions_revues SET {col_sql} = :val WHERE revue = :nom;", {"val": valeur, "nom": nom_revue})
                                except Exception:
                                    pass
                    st.success("✅ Félicitations ! Votre base PostgreSQL de Supabase est désormais entièrement peuplée.")
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
        st.info("ℹ️ L'application est connectée au Cloud Supabase, mais aucune revue n'a encore été créée par l'Éditeur.")
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
                    titre_affiche = section.replace("_", " ").capitalize()
                    st.subheader(titre_affiche)
                    st.write(str(contenu).strip())
