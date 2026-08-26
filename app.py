import streamlit as st
import io
import json
import os

# 1. CONFIGURATION DE LA PAGE
st.set_page_config(page_title="Centre de Ressources Éditoriales", page_icon="📚", layout="wide")

FICHIER_SAUVEGARDE = "revues_config.json"

# Liste complète de vos 45 revues
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
# 2. GENERATEUR DE GUIDE COMPATIBLE MICROSOFT WORD
def generer_instructions_word_html(nom_revue, config, options_article):
    """Génère un document Word (.doc) via une structure de dictionnaire pour éviter les coupures de texte."""
    ex_idline = f"{config['header']}, 22(1), 78-82, 2026" if config["id_line_format"] == "Pagination" else f"{config['header']}, 22, 62, 2026"
    ex_running = f"S. Mathy et al.: {config['header']}, 22(1), 78-82, 2026" if config["id_line_format"] == "Pagination" else f"S. Mathy et al.: {config['header']}, 22, 62, 2026"
    ex_citation = 'G. Liu, ... "TDM networks...", IEEE Trans. Comp., 1997.' if config["style_citation"] == "IEEE" else 'Weinstein, J. (2009). "The market..." Classical Philology.'

    # Entête HTML propre et stylisée pour Word
    html = f"""<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://w3.org">
    <head><meta charset="utf-8"><style>
        body {{ font-family: Arial, sans-serif; font-size: 10.5pt; line-height: 1.4; }}
        h1 {{ font-size: 16pt; color: {config['couleur']}; border-bottom: 2px solid {config['couleur']}; padding-bottom: 5px; }}
        h2 {{ font-size: 12pt; font-weight: bold; color: #333; margin-top: 15px; margin-bottom: 5px; }}
        .example {{ background-color: #f4f4f4; border-left: 4px solid #ccc; padding: 5px 10px; margin: 5px 0; font-style: italic; }}
    </style></head><body>
    <div style="text-align: right; font-size: 9pt; color: #666;">General Copy Editing procedure<br>Last update: 26/08/2026</div>
    <h1>General Copy Editing Procedure — {nom_revue.upper()}</h1>"""

    # Liste structurée des rubriques issues de votre document de référence
    rubriques = [
        ("IDLine", f"Format requis basé sur la <b>{config['id_line_format']}</b>.", f"Example: {ex_idline}"),
        ("Font Copyright", "If copyright to journal: <i>The Author(s), Published by EDP Sciences, 2026</i>.<br>If copyright to authors: <i>[Names], Published by EDP Sciences, 2026</i>.<br>1 author: [I. Surname] / 2 authors: [I. Surname and I. Surname] / 3+ authors: [I. Surname et al.]", "Example: J.M. Bertho and M. Bourguignon, Published by EDP Sciences, 2026"),
        ("Open Access", f"<b>Statut de la revue : {'Activé' if config['open_access'] else 'Désactivé'}</b>.<br>Si actif : Logo Open Access obligatoire en haut de page. Mention légale Creative Commons CC-BY 4.0 obligatoire en pied de page.", "distributed under the terms of the Creative Commons Attribution License..."),
        ("Special Issue", "Si l'article appartient à un numéro spécial, faire figurer au-dessus du bandeau :<br><i>Special Issue/Topical Issue - [Nom du numéro]<br>Guest Editors: [Noms]</i>", "The font should be the same as that of the main text."),
        ("Article type", "Affiché sur le bandeau à gauche. Doit correspondre strictement au type saisi dans SAGA (respecter la casse).", ""),
        ("Title / Translated title", "Pas de majuscules dans le titre (uniquement au premier mot, aux noms propres et espèces). Le titre traduit se place en gras au début du résumé traduit.", ""),
        ("List of authors", "Format : Prénom complet + Nom complet. Séparés par des virgules, sauf le dernier précédé de 'and'. Pas de point final.", "Example: Yi-Ping Wang1, Shi-Chuang Jiang1,2,* and Dong Sun1"),
        ("List of affiliations", "Numérotée si plusieurs affiliations. Au moins une adresse par auteur avec ville et pays. Pas de point final à la fin de la ligne ou des acronymes (USA, UK).", "Example: 1 School of Optoelectronic Engineering, Xiamen, PR China"),
        ("Corresponding author / Equal contribution", "Ajouter l'astérisque (*) après le nom et la note de bas de page avec l'e-mail.<br>Si contribution égale, ajouter le symbole (★) et la mention 'these authors contributed equally'.", "Example footnote: * Corresponding author: contact@email.com"),
        ("Abstract / Keywords", "Précédé de <b>Abstract</b> en gras. Mots-clés introduits par <b>Keywords:</b> en gras, le reste entièrement en minuscules, séparés par des slashes (/).", "Keywords: optics / laser / computing"),
        ("Translated abstract / keywords", "Précédé de <b>Résumé</b> en gras. Mots-clés introduits par <b>Mots-clés :</b> en gras, séparés par des slashes (/).", "Mots-clés : optique / laser"),
        ("Running title", f"Centré en haut de chaque page. Format généré pour cette revue :", f"Example: {ex_running}"),
        ("Sections", f"<b>Mise en règle :</b> {'Numbered in Arabic numerals (e.g., 1 Introduction)' if config['sections_numerotees'] else 'The sections are not numbered.'}", ""),
        ("Main text / Columns", f"Texte principal réglé en taille 10 pt. Structure configurée sur <b>{'deux colonnes' if config['deux_colonnes'] else 'une seule colonne standard'}</b>.", ""),
        ("Latin terms", "Doivent impérativement apparaître en police <b>Romaine</b> (pas d'italique) : <i>et al., in situ, in vitro, versus, vs., cf., i.e.</i>. Utiliser des espaces insécables à l'intérieur.", ""),
        ("Abbreviations", "Suivre les règles standards. Présenter la forme développée + la forme abrégée lors de la première apparition. Majuscule uniquement pour les noms de personnes ou d'organisations.", ""),
        ("Symbols, units & equations", "Variables et symboles en <i>italique</i> (normes ISO). Unités en romain (L, mL, min, °C, K) avec un espace obligatoire après le nombre (ex: 5 L). Équations numérotées séquentiellement à droite (1), sans ponctuation finale.", ""),
        ("Figures & Tables", "Figures : numérotées (Figure 1, Fig. 1), légende obligatoire <b>en dessous</b> se finissant par un point.<br>Tables : numérotées (Table 1, Tab. 1), légende obligatoire <b>au-dessus</b> se finissant par un point. Lignes verticales interdites.", ""),
        ("Final sections", f"Placées juste après la conclusion et avant les références. Typographie réglée en taille 9 pt.<br>Ordre obligatoire : Acknowledgments / Funding (<i>{options_article['funding_text']}</i>) / Conflicts of interest (<i>{options_article['conflict_text']}</i>) / Data availability statement (<i>{options_article['data_phrasing']}</i>) / Author contribution statement.", ""),
        ("Supplementary material", "Placé après les déclarations finales et avant les références.", "The supplementary material of this article is available at..."),
        ("References / Citation box", f"Toutes les références doivent être numérotées et correspondre au texte. Pas de majuscules aux titres de livres/revues. Aucun nom de revue ne commence par 'The'.<br>Format de la boîte de citation ({config['style_citation']}) :", f"Template: {ex_citation}"),
        ("Appendices & S2O Box", f"Appendices en taille 9 pt placés après les références.<br><b>S2O Box :</b> {'À insérer tout à la fin de l\'article (Modèle Subscribe to Open actif)' if options_article['is_s2o'] else 'Non applicable sur cette revue.'}", "")
    ]

    # Construction automatique du corps HTML
    for titre, texte, exemple in rubriques:
        html += f"<h2>{titre}</h2><p>{texte}</p>"
        if exemple:
            html += f"<div class='example'>{exemple}</div>"

    html += "</body></html>"
    return html.encode('utf-8')
# 3. GENERATEUR DU EXEMPLE VISUEL (.TEX)
def generer_visuel_latex(nom_revue, config, options_article):
    """Génère le code LaTeX d'illustration visuelle."""
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
    {{\\footnotesize 1 Nom du Laboratoire, Ville, Pays}}
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
\\footnotesize 1. D. Sarunyagate, Lasers. New York: McGraw-Hill, 1996.

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
        
        # --- MISE À JOUR : PANNEAU DE PREVISUALISATION INTÉGRAL ---
        st.write("---")
        st.subheader("👀 Guide de Copy-Editing direct (Conforme au fichier Word)")
        
        expander_rules = st.expander(f"📌 Consulter la Charte Officielle — {revue_choisie.upper()}", expanded=True)
        with expander_rules:
            # Éléments variables calculés
            ex_idline = f"{cfg['header']}, 22(1), 78-82, 2026" if cfg["id_line_format"] == "Pagination" else f"{cfg['header']}, 22, 62, 2026"
            ex_running = f"S. Mathy et al.: {cfg['header']}, 22(1), 78-82, 2026" if cfg["id_line_format"] == "Pagination" else f"S. Mathy et al.: {cfg['header']}, 22, 62, 2026"
            ex_citation = 'G. Liu, ... "TDM networks...", IEEE Trans. Comp., 1997.' if cfg["style_citation"] == "IEEE" else 'Weinstein, J. (2009). "The market..." Classical Philology.'

            # Affichage structuré
            st.markdown(f"### 📑 IDLine & En-tête")
            st.markdown(f"- **Format d'identification :** Basé sur la `{cfg['id_line_format']}`")
            st.info(f"**Exemple IDLine d'en-tête :** {ex_idline}")
            st.info(f"**Exemple Running Title (Haut de page centré) :** {ex_running}")

            st.markdown("### ✉️ Mentions Légales & Copyright")
            st.markdown(f"- **Style de Citation Bibliographique :** `{cfg['style_citation']}`")
            st.markdown(f"- **Mise en page générale :** `{'Deux colonnes' if cfg['deux_colonnes'] else 'Une seule colonne standard'}`")
            st.markdown(f"- **Numérotation :** `{'Sections numérotées (1 Introduction)' if cfg['sections_numerotees'] else 'Sections non numérotées'}`")
            if cfg['open_access']:
                st.success("🔓 **Revue en Open Access :** Logo obligatoire en haut de page. Mention légale Creative Commons CC-BY 4.0 obligatoire en pied de page.")
            
            st.markdown("### 🖋️ Structure du Manuscrit (Ordre Strict des Sections Finales)")
            st.markdown(f"Le typesetter doit obligatoirement respecter l'alignement en **Taille 9 pt** des sections dans l'ordre suivant :")
            st.code(f"1. Acknowledgments\n2. Funding -> {funding_text}\n3. Conflicts of interest -> {conflict_text}\n4. Data availability statement -> {data_phrasing}\n5. Author contribution statement\n6. Supplementary material\n7. References (Style {cfg['style_citation']})\n8. Citation Box\n9. Appendices / Annexes (9pt)" + ("\n10. S2O Box (Modèle Subscribe to Open Actif)" if is_s2o else ""), language="text")

            st.markdown("### ⚠️ Rappels Typographiques Majeurs (Directives Sources)")
            st.warning("• **Termes Latins :** Toujours en police **Romaine** (Pas d'italique pour *et al.*, *in situ*, *in vitro*, *versus*, *cf.*).\n\n• **Unités ISO :** Espace insécable obligatoire entre le nombre et l'unité (ex: 5 L, 24 °C, 10 min).\n\n• **Bibliographie :** Pas de majuscules superflues aux titres d'ouvrages. Aucun nom de revue ne doit démarrer par 'The'. Conserver le numéro DOI.")

        st.write("---")
        st.subheader("📥 Documents à exporter")
        
        col_btn1, col_btn2 = st.columns(2)
        
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
