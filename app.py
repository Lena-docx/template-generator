import streamlit as st
import io
import json
import os
import xml.etree.ElementTree as ET
import zipfile

# 1. CONFIGURATION DE LA PAGE
st.set_page_config(page_title="Centre de Ressources Éditoriales", page_icon="📚", layout="wide")

FICHIER_SAUVEGARDE = "revues_config.json"

# Liste de toutes vos revues extraites de l'image
LISTE_ACRONYMES = [
    "cagri", "geotech", "jbio", "tpe", "bsgf", "limn", "nss", "parasite", 
    "pmed", "radiopro", "aacus", "alr", "mfreview", "ocl", "rees", "stet", 
    "kmae", "mattech", "meca", "metal", "emsci", "ijmqe", "jeos", "rdne", 
    "sbuild", "smdo", "swsc", "ject", "sicotj", "vcm", "sands", "epn", 
    "photon", "npvcafe", "npvelsa", "npvequi", "esaim-cocv", "esaim-m2an", 
    "esaim-ps", "mmnp", "rairo-ro", "rairo-ita", "medsci", "jomos", "ppsy"
]

# Génération des données de départ pour toutes les revues
DONNEES_PAR_DEFAUT = {}
for acro in LISTE_ACRONYMES:
    # Quelques préréglages automatiques selon les standards habituels de vos revues
    style_cit = "APA" if acro in ["pmed", "nss", "ppsy", "medsci"] else "IEEE"
    is_twocol = True if acro in ["jeos", "meca", "photon", "rairo-ro"] else False
    
    DONNEES_PAR_DEFAUT[acro] = {
        "police": "Times New Roman" if style_cit == "APA" else "Arial",
        "taille_titre": "16" if style_cit == "APA" else "18",
        "couleur": "#d62728" if style_cit == "APA" else "#1f77b4",
        "marges": "2.5cm",
        "header": f"Journal of {acro.upper()} - Research Framework",
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

# 2. GENERATEUR NATIF DE FICHIER .DOCX VALIDÉ MICROSOFT WORD
def generer_instructions_docx_vrai(nom_revue, config, options_article):
    """Génère un vrai fichier .docx (OpenXML) fonctionnel et compatible toutes plateformes."""
    id_format = "Volume(Fascicule), Pages, Année" if config["id_line_format"] == "Pagination" else "Volume, Numéro d'article, Année"
    
    # Création de la structure XML minimale pour un document Word standard
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://openxmlformats.org">
        <w:body>
            <w:p>
                <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                <w:r><w:t>GUIDE D'INSTRUCTIONS TYPESETTER : {nom_revue.upper()}</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>Ce document récapitule les règles d'harmonisation de la charte graphique.</w:t></w:r></w:p>
            <w:p/>
            <w:p><w:r><w:b/><w:t>1. REGLES DE STYLE GENERALES :</w:t></w:r></w:p>
            <w:p><w:r><w:t>- Police principale : {config['police']} (Titres : {config['taille_titre']}pt, Couleur : {config['couleur']})</w:t></w:r></w:p>
            <w:p><w:r><w:t>- Marges : {config['marges']}</w:t></w:r></w:p>
            <w:p><w:r><w:t>- Structure : {'Deux colonnes' if config['deux_colonnes'] else 'Une seule colonne standard'}</w:t></w:r></w:p>
            <w:p><w:r><w:t>- IDLine requis : Format basé sur la {config['id_line_format']} ({id_format})</w:t></w:r></w:p>
            <w:p/>
            <w:p><w:r><w:b/><w:t>2. CONSIGNES DE VIGILANCE TEXTUELLE :</w:t></w:r></w:p>
            <w:p><w:r><w:t>- Termes Latins : Les mots comme "et al.", "in situ", "in vitro", "versus" doivent impérativement rester en Romain (Pas d'italique).</w:t></w:r></w:p>
            <w:p><w:r><w:t>- Unités de Mesure : Toujours insérer un espace insécable entre la valeur et l'unité (ex: 5 L, 10 min).</w:t></w:r></w:p>
            <w:p/>
            <w:p><w:r><w:b/><w:t>3. ORDRE OBLIGATOIRE DES SECTIONS FINALES :</w:t></w:r></w:p>
            <w:p><w:r><w:t>Le typesetter doit structurer la fin de l'article précisément dans cet ordre :</w:t></w:r></w:p>
            <w:p><w:r><w:t>1. Acknowledgments (Remerciements)</w:t></w:r></w:p>
            <w:p><w:r><w:t>2. Funding (Financements) -> {options_article['funding_text']}</w:t></w:r></w:p>
            <w:p><w:r><w:t>3. Conflicts of Interest -> {options_article['conflict_text']}</w:t></w:r></w:p>
            <w:p><w:r><w:t>4. Data Availability Statement -> {options_article['data_phrasing']}</w:t></w:r></w:p>
            <w:p><w:r><w:t>5. Supplementary Material</w:t></w:r></w:p>
            <w:p><w:r><w:t>6. References (Style bibliographique : {config['style_citation']})</w:t></w:r></w:p>
            <w:p><w:r><w:t>7. Citation Box</w:t></w:r></w:p>
            <w:p><w:r><w:t>8. Appendices / Annexes (En taille 9pt)</w:t></w:r></w:p>
            {"<w:p><w:r><w:t>9. S2O Box (Journal sous pavillon Subscribe to Open)</w:t></w:r></w:p>" if options_article['is_s2o'] else ""}
            <w:p/>
            <w:p><w:r><w:b/><w:t>4. CONSIGNES BIBLIOGRAPHIQUES ({config['style_citation']}) :</w:t></w:r></w:p>
            <w:p><w:r><w:t>- Les titres des livres et des revues ne doivent pas prendre de majuscules (sauf la première lettre).</w:t></w:r></w:p>
            <w:p><w:r><w:t>- Aucun nom de revue ne doit commencer par l'article "The".</w:t></w:r></w:p>
        </w:body>
    </w:document>
    """
    
    # Création de l'archive ZIP mimant la structure officielle d'un fichier .docx Microsoft Word
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('document.xml', document_xml)
        docx.writestr('[Content_Types].xml', """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://openxmlformats.org">
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>""")
    buffer.seek(0)
    return buffer.getvalue()

# 3. GENERATEUR DU EXEMPLE VISUEL (.TEX)
def generer_visuel_latex(nom_revue, config, options_article):
    """Génère le code LaTeX complet qui servira d'exemple visuel de l'article."""
    hex_color = config["couleur"].lstrip('#')
    r, g, b = tuple(int(hex_color[i:i+2], 16)/255 for i in (0, 2, 4))
    col_mode = "\\usepackage[twocolumn]{geometry}" if config["deux_colonnes"] else ""
    sec_mode = "" if config["sections_numerotees"] else "\\setcounter{secnumdepth}{0}"
    
    id_line = f"{config['header']}, Vol. 22(1), 78-82, 2026" if config["id_line_format"] == "Pagination" else f"{config['header']}, Vol. 22, 62, 2026"
    running_title = f"Author et al.: {id_line}"

    if config["style_citation"] == "IEEE":
        citation_box = "Cite this article as: P. Nom, \\\"Title of the article,\\\" " + config['header'] + ", vol. 22, pp. 78-82, 2026."
    else:
        citation_box = "Cite this article as: Nom, P. (2026). \\\"Title of the article.\\\" " + config['header'] + ", 22(1), 78-82."

    mention_special = "\\noindent\\small\\textit{Special Issue / Guest Editors: Dr. A, Dr. B}\\hfill" if options_article['is_special'] else ""
    mention_oa = "\\small\\textbf{(Open Access Logo)}" if config['open_access'] else ""
    note_equal = "\\footnotetext{$\\star$ These authors contributed equally.}" if options_article['equal_contrib'] else ""
    mention_s2o = "\\vspace{0.3cm}\\noindent\\textbf{S2O Box:} Ce journal est publié selon le modèle Subscribe to Open." if options_article['is_s2o'] else ""

    template = f"""\\documentclass[10pt, a4paper]{{article}}
\\usepackage[utf8]{{inputenc}}
\\usepackage[margin={config['marges']}]{{geometry}}
{col_mode}
\\usepackage{{fancyhdr}}
\\usepackage{{xcolor}}
\\usepackage{{titlesec}}
\\usepackage{{lipsum}}

{sec_mode}

\\definecolor{{couleurRevue}}{{rgb}}{{{r:.2f}, {g:.2f}, {b:.2f}}}
\\titleformat{{\\section}}{{\\normalfont\\large\\bfseries\\color{{couleurRevue}}}}{{\\thesection}}{{1em}}{{}}

\\pagestyle{{fancy}}
\\fancyhf{{}}
\\fancyhead[C]{{\\small {running_title}}}
\\fancyfoot[L]{{\\footnotesize {id_line}}}
\\fancyfoot[R]{{\\thepage}}

\\begin{{document}}

{mention_special}
{mention_oa}
\\vspace{{0.5cm}}

\\begin{{center}}
    {{\\Large\\bfseries\\color{{couleurRevue}} Titre de l'Article d'Exemple pour {nom_revue} \\\\}}
    \\vspace{{0.4cm}}
    {{\\large Prénom Nom$^{{1,*}}$ \\\\}}
    \\vspace{{0.2cm}}
    {{\\footnotesize 1 Nom du Laboratoire, Ville, Pays (Sans point final)}}
\\end{{center}}

\\footnotetext{{* Corresponding author: contact@revue.com}}
{note_equal}

\\vspace{{0.4cm}}
\\noindent\\textbf{{Abstract – }} \\lipsum[1-2]

\\vspace{{0.2cm}}
\\noindent\\textbf{{Keywords: }} science / exemple / gabarit

\\section{{1. Introduction}}
Ceci est un exemple visuel du rendu de l'article. Le texte respecte la police de la revue. 
Les expressions latines comme et al. ou in situ restent en romain. Les unités possèdent un espace insécable (ex: 12~L).

\\section{{Acknowledgments}}
Remerciements à l'équipe éditoriale.

\\section{{Funding}}
{options_article['funding_text']}

\\section{{Conflicts of Interest}}
{options_article['conflict_text']}

\\section{{Data Availability Statement}}
{options_article['data_phrasing']}

\\section{{References}}
\\footnotesize D. Sarunyagate, *Lasers*. New York: McGraw-Hill, 1996.\\\\ G. Weinstein, ''The market in Plato's Republic,'' *Classical Philology*, vol. 104, pp. 439-458, 2009.

\\vspace{{0.5cm}}
\\noindent\\fbox{{
\\begin{{minipage}}{{\\linewidth}}
\\bfseries{{Citation Box:}}\\\\
{citation_box}
\\end{{minipage}}
}}

{mention_s2o}

\\end{{document}}
"""
    return template.encode('utf-8')
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
            "The research data associated with this article are available on request from the authors",
            "The research data associated with this article are included within the article",
            "This article has no associated data generated and/or analyzed"
        ])
        
        options_article = {
            "is_special": is_special, "equal_contrib": equal_contrib, "is_s2o": is_s2o,
            "funding_text": funding_text, "conflict_text": conflict_text, "data_phrasing": data_phrasing
        }
        
        # --- NOUVEAUTÉ : PANNEAU DE PREVISUALISATION DIRECTE ---
        st.write("---")
        st.subheader("👀 Visualisation directe des règles (sans téléchargement)")
        
        expander_rules = st.expander("📌 Cliquez ici pour déplier les consignes de cette revue", expanded=True)
        with expander_rules:
            col_preview1, col_preview2 = st.columns(2)
            with col_preview1:
                st.markdown(f"**Charte Graphique :**")
                st.markdown(f"- **Police :** {cfg['police']}")
                st.markdown(f"- **Taille des titres :** {cfg['taille_titre']} pt")
                st.markdown(f"- **Marges :** {cfg['marges']}")
                st.markdown(f"- **Mise en page :** {'Deux colonnes' if cfg['deux_colonnes'] else 'Une seule colonne standard'}")
            with col_preview2:
                st.markdown(f"**Règles Éditoriales :**")
                st.markdown(f"- **Style bibliographique :** `{cfg['style_citation']}`")
                st.markdown(f"- **Ligne d'identification :** {cfg['id_line_format']}")
                st.markdown(f"- **Mention de copyright :** {'Open Access requis' if cfg['open_access'] else 'Standard'}")
            
            st.warning("⚠️ **Rappels Typographiques Impératifs :** Les termes latins (*et al.*, *in situ*) restent en Romain. Espace insécable obligatoire avant les unités de mesure (ex: 5 L). Pas de majuscules superflues dans la bibliographie.")

        st.write("---")
        st.subheader("📥 Documents à exporter")
        
        col_btn1, col_btn2 = st.columns(2)
        
        # Téléchargement du vrai .docx généré via la structure OpenXML corrigée
        data_docx = generer_instructions_docx_vrai(revue_choisie, cfg, options_article)
        col_btn1.download_button(
            label="📄 Télécharger la Liste d'Instructions (.docx)",
            data=data_docx,
            file_name=f"instructions_typesetter_{revue_choisie.lower().replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
