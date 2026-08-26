import streamlit as st
import io
import json
import os

# 1. CONFIGURATION DE LA PAGE
st.set_page_config(page_title="Centre de Ressources Éditoriales", page_icon="📚", layout="wide")

FICHIER_SAUVEGARDE = "revues_config.json"

# Liste complète de vos revues extraites de votre image précédente
LISTE_ACRONYMES = [
    "cagri", "geotech", "jbio", "tpe", "bsgf", "limn", "nss", "parasite", 
    "pmed", "radiopro", "aacus", "alr", "mfreview", "ocl", "rees", "stet", 
    "kmae", "mattech", "meca", "metal", "emsci", "ijmqe", "jeos", "rdne", 
    "sbuild", "smdo", "swsc", "ject", "sicotj", "vcm", "sands", "epn", 
    "photon", "npvcafe", "npvelsa", "npvequi", "esaim-cocv", "esaim-m2an", 
    "esaim-ps", "mmnp", "rairo-ro", "rairo-ita", "medsci", "jomos", "ppsy"
]

DONNEES_PAR_DEFAUT = {}
for acro in LISTE_ACRONYMES:
    style_cit = "APA" if acro in ["pmed", "nss", "ppsy", "medsci"] else "IEEE"
    is_twocol = True if acro in ["jeos", "meca", "photon", "rairo-ro"] else False
    
    DONNEES_PAR_DEFAUT[acro] = {
        "police": "Times New Roman" if style_cit == "APA" else "Arial",
        "taille_titre": "16" if style_cit == "APA" else "18",
        "couleur": "#d62728" if style_cit == "APA" else "#1f77b4",
        "marges": "2.5cm",
        "header": f"{acro.upper()}",
        "id_line_format": "Article Number" if style_cit == "APA" else "Pagination",
        "style_citation": style_cit,
        "open_access": True,
        "deux_colonnes": is_twocol,
        "sections_numerotees": not is_twocol
    }

def charger_donnees():
    if os.path.exists(FICHIER_SAUVEGARDE):
        try:
            with open(FICHIER_SAUVEGARDE, "r", encoding="utf-8") as f: 
                return json.load(f)
        except Exception: 
            return DONNEES_PAR_DEFAUT
    return DONNEES_PAR_DEFAUT

def sauvegarder_donnees(donnees):
    with open(FICHIER_SAUVEGARDE, "w", encoding="utf-8") as f: 
        json.dump(donnees, f, ensure_ascii=False, indent=4)

if "revues" not in st.session_state:
    st.session_state.revues = charger_donnees()

# 2. GENERATEUR DE FICHIER WORD (.DOC) REPRENANT EXACTEMENT VOTRE DOCUMENT SOURCE
def generer_instructions_word_html(nom_revue, config, options_article):
    """Génère un document Word (.doc) basé sur un format HTML parfaitement interprété par MS Word."""
    
    # Construction dynamique des exemples selon la configuration de la revue
    if config["id_line_format"] == "Pagination":
        ex_idline = f"{config['header']}, 22(1), 78-82, 2026"
        ex_running = f"S. Mathy et al.: {config['header']}, 22(1), 78-82, 2026"
    else:
        ex_idline = f"{config['header']}, 22, 62, 2026"
        ex_running = f"S. Mathy et al.: {config['header']}, 22, 62, 2026"

    if config["style_citation"] == "IEEE":
        ex_citation = f'G. Liu, K. Y. Lee, and H. F. Jordan, "TDM and TWDM de Bruijn networks and shufflenets for optical communications," IEEE Trans. Comp., vol. 46, pp. 695-701, June 1997.'
    else:
        ex_citation = f'Weinstein, J. (2009). "The market in Plato\'s Republic." Classical Philology, 104(4), 439-458.'

    html_content = f"""
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://w3.org">
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; font-size: 10.5pt; line-height: 1.4; }}
            h1 {{ font-size: 16pt; color: {config['couleur']}; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 20px; }}
            h2 {{ font-size: 12pt; font-weight: bold; margin-top: 15px; margin-bottom: 5px; }}
            .table-style {{ border-collapse: collapse; width: 100%; margin-top: 10px; }}
            .table-style td, .table-style th {{ border: 1px solid #000; padding: 6px; text-align: left; font-size: 10pt; }}
            .table-style th {{ background-color: #f2f2f2; }}
            .example {{ background-color: #f9f9f9; border-left: 3px solid #ccc; padding: 5px 10px; margin: 5px 0; font-style: italic; }}
            .highlight {{ color: #d62728; font-weight: bold; }}
        </style>
    </head>
    <body>
        <div style="font-size: 9pt; text-align: right; margin-bottom: 20px; color: #555;">
            General Copy Editing procedure<br>Last update: 26/08/2026
        </div>

        <h1>General Copy Editing Procedure — {nom_revue.upper()}</h1>

        <h2>IDLine</h2>
        {"<p><b>If there are article numbers:</b><br>[Journal shortened name], Volume, Article number ([Year])</p><div class='example'>Example: " + ex_idline + "</div>" if config["id_line_format"] == "Article Number" else "<p><b>If there is a pagination:</b><br>[Journal shortened name], Volume(Issue), Page numbers([Year])</p><div class='example'>Example: " + ex_idline + "</div>"}

        <h2>Font Copyright</h2>
        <p><b>If the copyright is to the journal:</b><br>The Author(s), Published by EDP Sciences, Year<br><div class='example'>Example: The Author(s), Published by EDP Sciences, 2026</div></p>
        <p><b>If the copyright is to the authors:</b><br>[Name of the authors], Published by EDP Sciences, [Year]<br><div class='example'>Example: J.M. Bertho and M. Bourguignon, Published by EDP Sciences, 2026</div></p>
        <p>• <i>If there is one author:</i> [Initial + Surname], Published by EDP Sciences, [Year]<br>
        • <i>If there are two authors:</i> [Initial + Surname and Initial + Surname], Published by EDP Sciences, [Year]<br>
        • <i>If there are three or more authors:</i> [Initial + Surname et al.], Published by EDP Sciences, [Year]</p>

        <h2>Open Access</h2>
        <p><b>If the journal is in Open Access:</b><br>There must always be :<br>
        - At the top of the page : the <b>Open Access</b> logo<br>
        - At the bottom of the page : the mention <i>"This is an Open Access article distributed under the terms of the Creative Commons Attribution License (https://creativecommons.org), which permits unrestricted use, distribution, and reproduction in any medium, provided the original work is properly cited."</i></p>

        <h2>Special Issue</h2>
        <p>If the article belongs to a Special Issue or a Topical Issue, the name of the Special Issue should appear above the banner, like so:<br>
        <div class='example'>Special Issue/Topical Issue - [Name of the Special Issue/Topical Issue]<br>Guest Editors: [Names of the Guest Editors]</div>
        The font should be the same as that of the main text.</p>

        <h2>Article type</h2>
        <p>The article type is displayed on the banner, on the left. It must be written identically to what is in SAGA. Please respect the uppercase/lowercase letters.</p>

        <h2>Title</h2>
        <p>No capital in the title (only at the beginning of the first word of the title, for proper nouns, name of species).</p>

        <h2>Translated title</h2>
        <p>If there is a translated title: No capital in the title (only at the beginning of the first word of the title, for proper nouns, name of species). The translated title is set in bold at the beginning of the translated abstract.</p>

        <h2>List of authors</h2>
        <p>Full first name + full last name. The names are separated by commas, except for the last one, which is preceded by <b>and</b>. There is no full stop at the end of the list of authors.<br>
        <div class='example'>Example: Yi-Ping Wang1, Shi-Chuang Jiang1,2,* and Dong Sun1</div></p>

        <h2>List of affiliations</h2>
        <p>The list must be numbered if there is more than one affiliation. There must be at least one address for each author. The city and country must be included in the address. <b>No full stop</b> at the end of the addresses and after acronyms (example: USA, PR China, UK, PO Box).<br>
        <div class='example'>Example: 1 School of Optoelectronic and Communication Engineering, Xiamen University of Technology, Xiamen 361024, PR China</div></p>

        <h2>Corresponding author</h2>
        <p>Put the symbol <b>*</b> after the name of the corresponding author. Add this footnote:<br>
        <div class='example'>* Corresponding author: [e-mail of the corresponding author]</div></p>

        <h2>Equal contribution</h2>
        <p>When two authors contributed equally to the manuscript, add the symbol <b>&star;</b> and the footnote: <i>these authors contributed equally</i>.</p>

        <h2>Abstract</h2>
        <p>The abstract is preceded by <b>Abstract</b> in bold.</p>

        <h2>Keywords</h2>
        <p><b>Keywords:</b> keyword 1 / keyword 2 / keyword 3<br>
        <b>Keywords:</b> is in bold with an uppercase letter at the start. The keywords are all lowercase.</p>

        <h2>Translated abstract</h2>
        <p>The translated abstract is preceded by <b>Résumé</b> (or translated title) in bold.</p>

        <h2>Translated keywords</h2>
        <p><b>Mots-clés :</b> mot-clé 1 / mot-clé 2<br>
        <b>Mots-clés :</b> is in bold with an uppercase letter at the start. The keywords are all lowercase.</p>

        <h2>Running title</h2>
        <p>Centered and at the top of every page.<br>
        {"<b>If there are article numbers:</b> [Initial + surname of the first author] et al.: [Journal shortened name], Volume, Article number ([Year])<br><div class='example'>Example: " + ex_running + "</div>" if config["id_line_format"] == "Article Number" else "<b>If there is a pagination:</b> [Initial + surname of the first author] et al.: [Journal shortened name], Volume(Issue), Page numbers([Year])<br><div class='example'>Example: " + ex_running + "</div>"}</p>

        <h2>Sections</h2>
        {"<p><b>If the sections are numbered:</b> The sections are numbered in Arabic numerals.<br><div class='example'>Example:<br>1 Introduction<br>2 Theory</div></p>" if config["sections_numerotees"] else "<p><b>If the sections are not numbered:</b> The sections are not numbered.</p>"}

        <h2>Main text</h2>
        <p>The main text of the article should be typeset in font size 10.</p>

        <h2>Columns</h2>
        <p>The main text should be typeset into <b>{'two columns' if config['deux_colonnes'] else 'one column'}</b>.</p>

        <h2>Latin terms</h2>
        <p>Latin terms must appear in <b>roman font</b> (e.g., i.e., cf., et al., in situ, versus, vs., ab initio, a priori, via, in vitro, ad hoc, ...); unbreakable spaces within.</p>

        <h2>Abbreviations</h2>
        <p>1. The standard rules for abbreviation must be followed.<br>
        2. When the authors use an abbreviation for the first time, they need to present both the spelled-out version and the short form.<br>
        3. In the spelled-out version, capital letter for the initial letter of each word is only needed for the names of organization or person.</p>

        <h2>Symbol, units, equations, functions, and numbers</h2>
        <p>All measurements, data and symbols (variables) should be given using international norms (ISO) and should always be written in <i>italic</i>. SI units should be used: the unit "litre" should be abbreviated as "L" (also mL, etc.), minutes as min, degrees as °C or K. All units should be typeset in roman. <b>There must have a space between a number and its unit.</b><br>
        Equations that are referred to in the text should be numbered with the number on the right-hand side and should be numbered sequentially throughout the text (i.e., (1), (2), (3)). There is no punctuation at the end.</p>

        <h2>Figures</h2>
        <p>All figures must be cited within the text. Figures should be numbered sequentially as Figure 1, Figure 2, etc. They are referred to in the text as Figure 1, Figure 2, (Fig. 1) etc. Captions should be placed <b>below</b> the figure. The caption is mandatory. There must be a full-stop at the end.</p>

        <h2>Tables</h2>
# 4. INTERFACE GRAPHIQUE STREAMLIT
st.title("📚 Centre de Ressources Éditoriales")
st.caption("Génération automatisée des paquets d'instructions et des maquettes visuelles.")

onglet_typesetter, onglet_editeur = st.tabs(["🚀 Espace Téléchargement (Typesetters)", "⚙️ Configuration des Chartes (Éditeurs)"])

with onglet_typesetter:
    st.header("Paquet de Publication par Revue")
    st.write("Sélectionnez la revue et configurez l'article pour exporter ou consulter vos documents.")
    
    # Menu déroulant affichant l'intégralité de vos acronymes (de cagri à ppsy)
    revue_choisie = st.selectbox("Sélectionner la revue :", list(st.session_state.revues.keys()))
    
    if revue_choisie:
        cfg = st.session_state.revues[revue_choisie]
        
        st.subheader("⚙️ Métadonnées de l'article à traiter")
        c_opt1, c_opt2 = st.columns(2)
        with c_opt1:
            is_special = st.checkbox("L'article fait partie d'un Numéro Spécial")
            equal_contrib = st.checkbox("Les auteurs ont une contribution égale (★)")
            is_s2o = st.checkbox("La revue utilise le modèle S2O sur cet article")
        with c_opt2:
            has_funding = st.toggle("Présence de financements spécifiques", value=False)
            has_conflict = st.toggle("Présence de conflits d'intérêts", value=False)
            
        funding_text = "This research was funded by institutional grants." if has_funding else "This research did not receive any specific funding."
        conflict_text = "The authors declare the following conflicts..." if has_conflict else "The authors declare that they have no competing interests to report."
        
        data_phrasing = st.selectbox("Data Availability Statement :", [
            "The research data available on request from the authors",
            "The research data associated with this article are included within the article",
            "This article has no associated data generated and/or analyzed",
            "The research data associated with this article are available in [Name of public data repository], under the reference [DOI or other data identifier]",
            "Data associated with this article cannot be disclosed due to legal/ethical/other reason."
        ])
        
        options_article = {
            "is_special": is_special, "equal_contrib": equal_contrib, "is_s2o": is_s2o,
            "funding_text": funding_text, "conflict_text": conflict_text, "data_phrasing": data_phrasing
        }
        
        # Panneau de prévisualisation directe à l'écran
        st.write("---")
        st.subheader("👀 Visualisation directe des règles (sans téléchargement)")
        
        expander_rules = st.expander("📌 Cliquez ici pour déplier les consignes de cette revue", expanded=True)
        with expander_rules:
            col_preview1, col_preview2 = st.columns(2)
            with col_preview1:
                st.markdown("**Charte Graphique :**")
                st.markdown(f"- **Police demandée :** {cfg['police']}")
                st.markdown(f"- **Taille des titres :** {cfg['taille_titre']} pt")
                st.markdown(f"- **Marges :** {cfg['marges']}")
                st.markdown(f"- **Mise en page :** {'Deux colonnes' if cfg['deux_colonnes'] else 'Une seule colonne standard'}")
            with col_preview2:
                st.markdown("**Règles Éditoriales :**")
                st.markdown(f"- **Style bibliographique :** `{cfg['style_citation']}`")
                st.markdown(f"- **Ligne d'identification :** {cfg['id_line_format']}")
                st.markdown(f"- **Mention de copyright :** {'Open Access requis' if cfg['open_access'] else 'Standard'}")
            
            st.warning("⚠️ **Rappels Typographiques Source :** Les termes latins (*et al.*, *in situ*) doivent rester en Romain (pas d'italique). Espace insécable obligatoire avant les unités de mesure (ex: 5 L, 25 °C).")

        st.write("---")
        st.subheader("📥 Documents à exporter")
        
        col_btn1, col_btn2 = st.columns(2)
        
        # Téléchargement du .doc HTML encapsulé (sans blocage Word)
        data_docx = generer_instructions_word_html(revue_choisie, cfg, options_article)
        col_btn1.download_button(
            label="📄 Télécharger la Liste d'Instructions (.doc / Word)",
            data=data_docx,
            file_name=f"instructions_typesetter_{revue_choisie.lower().replace(' ', '_')}.doc",
            mime="application/msword"
        )
        
        data_tex = generer_visuel_latex(revue_choisie, cfg, options_article)
        col_btn2.download_button(
            label="🛠️ Télécharger l'Exemple de Code Visuel (.tex)",
            data=data_tex,
            file_name=f"exemple_visuel_{revue_choisie.lower().replace(' ', '_')}.tex",
            mime="text/plain"
        )
with onglet_editeur:
    st.header("Édition des chartes graphiques par Revue")
    action = st.radio("Action :", ["Modifier une revue existante", "Créer une nouvelle revue"])
    
    if action == "Modifier une revue existante":
        revue_a_modifier = st.selectbox("Sélectionner la revue à éditer :", list(st.session_state.revues.keys()))
        cfg_actuelle = st.session_state.revues[revue_a_modifier]
        
        with st.form("form_edit"):
            police = st.text_input("Police", value=cfg_actuelle["police"])
            taille = st.text_input("Taille des titres", value=cfg_actuelle["taille_titre"])
            couleur = st.color_picker("Couleur de la revue", value=cfg_actuelle["couleur"])
            marges = st.text_input("Marges (ex: 2.5cm)", value=cfg_actuelle["marges"])
            header = st.text_input("Nom abrégé du journal (Header/IDLine)", value=cfg_actuelle["header"])
            id_line_format = st.selectbox("Format de l'IDLine", ["Pagination", "Article Number"], index=0 if cfg_actuelle["id_line_format"] == "Pagination" else 1)
            style_citation = st.selectbox("Style de Citation bibliographique", ["IEEE", "APA"], index=0 if cfg_actuelle["style_citation"] == "IEEE" else 1)
            open_access = st.checkbox("Revue obligatoirement en Open Access", value=cfg_actuelle["open_access"])
            deux_colonnes = st.checkbox("Mise en page sur deux colonnes", value=cfg_actuelle["deux_colonnes"])
            sections_numerotees = st.checkbox("Numérotation des sections en chiffres arabes", value=cfg_actuelle["sections_numerotees"])
            
            if st.form_submit_button("Enregistrer les règles éditoriales"):
                st.session_state.revues[revue_a_modifier] = {
                    "police": police, "taille_titre": taille, "couleur": couleur, "marges": marges, "header": header,
                    "id_line_format": id_line_format, "style_citation": style_citation, "open_access": open_access,
                    "deux_colonnes": deux_colonnes, "sections_numerotees": sections_numerotees
                }
                sauvegarder_donnees(st.session_state.revues)
                st.success("Charte graphique mise à jour sur le disque.")
                st.rerun()

    elif action == "Créer une nouvelle revue":
        with st.form("form_add"):
            nom_nouvel = st.text_input("Nom de la nouvelle revue")
            police = st.text_input("Police", value="Arial")
            taille = st.text_input("Taille Titres", value="16")
            couleur = st.color_picker("Couleur", value="#000000")
            marges = st.text_input("Marges", value="2.5cm")
            header = st.text_input("Nom abrégé du journal", value="J. Short Name")
            id_line_format = st.selectbox("Format IDLine", ["Pagination", "Article Number"])
            style_citation = st.selectbox("Style Citation", ["IEEE", "APA"])
            open_access = st.checkbox("Open Access")
            deux_colonnes = st.checkbox("Deux colonnes")
            sections_numerotees = st.checkbox("Sections numérotées")
            
            if st.form_submit_button("Créer et Sauvegarder la revue") and nom_nouvel:
                st.session_state.revues[nom_nouvel] = {
                    "police": police, "taille_titre": taille, "couleur": couleur, "marges": marges, "header": header,
                    "id_line_format": id_line_format, "style_citation": style_citation, "open_access": open_access,
                    "deux_colonnes": deux_colonnes, "sections_numerotees": sections_numerotees
                }
                sauvegarder_donnees(st.session_state.revues)
                st.success(f"Revue '{nom_nouvel}' ajoutée à la base.")
                st.rerun()
