import streamlit as st
import io
import json
import os

# 1. CONFIGURATION DE LA PAGE
st.set_page_config(page_title="Générateur de Templates & Instructions", page_icon="📚", layout="wide")

FICHIER_SAUVEGARDE = "revues_config.json"

DONNEES_PAR_DEFAUT = {
    "Revue Alpha (IEEE)": {
        "police": "Arial", "taille_titre": "18", "couleur": "#1f77b4", "marges": "2.0cm",
        "header": "J. Eur. Opt. Society-Rapid Publ.", "id_line_format": "Pagination",
        "style_citation": "IEEE", "open_access": True, "deux_colonnes": True, "sections_numerotees": True
    },
    "Revue Beta (APA)": {
        "police": "Times New Roman", "taille_titre": "16", "couleur": "#d62728", "marges": "2.5cm",
        "header": "Annals of Modern Physics", "id_line_format": "Article Number",
        "style_citation": "APA", "open_access": False, "deux_colonnes": False, "sections_numerotees": False
    }
}

def charger_donnees():
    if os.path.exists(FICHIER_SAUVEGARDE):
        try:
            with open(FICHIER_SAUVEGARDE, "r", encoding="utf-8") as f: return json.load(f)
        except Exception: return DONNEES_PAR_DEFAUT
    return DONNEES_PAR_DEFAUT

def sauvegarder_donnees(donnees):
    with open(FICHIER_SAUVEGARDE, "w", encoding="utf-8") as f: json.dump(donnees, f, ensure_ascii=False, indent=4)

if "revues" not in st.session_state:
    st.session_state.revues = charger_donnees()

# 2. GENERATEUR DE LA LISTE D'INSTRUCTIONS (.DOCX / RTF)
def generer_instructions_docx(nom_revue, config, options_article):
    """Génère le guide d'instructions textuelles pour le typesetter au format Word."""
    id_format = "Volume(Fascicule), Pages, Année" if config["id_line_format"] == "Pagination" else "Volume, Numéro d'article, Année"
    
    texte_rtf = f"""{{\\rtf1\\ansi\\deff0
{{\\fonttbl{{\\f0\\fnil\\fcharset0 Arial;}}}}
{{\\colortbl ;\\red{int(config['couleur'][1:3],16)}\\green{int(config['couleur'][3:5],16)}\\blue{int(config['couleur'][5:7],16)};}}
\\paperw11906\\paperh16838\\margl1440\\margr1440\\margt1440\\margb1440
\\f0\\fs28\\b\\cf1 GUIDE D'INSTRUCTIONS POUR LE TYPESETTER - {nom_revue.upper()}\\b0\\cf0\\fs20\\par
\\line
Ce document récapitule les règles strictes de mise en page à appliquer pour la revue \\b {nom_revue}\\b0.\\par
\\line
\\b 1. RÈGLES DE STYLE GÉNÉRALES\\b0\\par
- \\b Police principale :\\b0 {config['police']} (Titres : {config['taille_titre']}pt, Couleur : {config['couleur']})\\par
- \\b Marges :\\b0 {config['marges']}\\par
- \\b Structure :\\b0 {'Deux colonnes (Mise en page symétrique)' if config['deux_colonnes'] else 'Une seule colonne standard'}\\par
- \\b Numérotation :\\b0 {'Sections numérotées en chiffres arabes' if config['sections_numerotees'] else 'Sections non numérotées'}\\par
- \\b IDLine requis :\\b0 Format basé sur la \\b {config['id_line_format']}\\b0 ({id_format})\\par
\\line
\\b 2. CONSIGNES DE VIGILANCE TEXTUELLE\\b0\\par
- \\b Termes Latins :\\b0 Les mots comme "et al.", "in situ", "in vitro", "versus" doivent impérativement rester en \\b Romain\\b0 (Pas d'italique) et sans coupure de mot en fin de ligne.\\par
- \\b Unités de Mesure :\\b0 Toujours insérer un espace insécable entre la valeur et l'unité (ex: 5 L, 10 min, 25 °C).\\par
\\line
\\b 3. ORDRE OBLIGATOIRE DES SECTIONS FINALES\\b0\\par
Le typesetter doit structurer la fin de l'article précisément dans cet ordre :\\par
1. Acknowledgments (Remerciements)\\par
2. Funding (Financements) -> \\i {options_article['funding_text']}\\i0\\par
3. Conflicts of Interest (Conflits d'intérêt) -> \\i {options_article['conflict_text']}\\i0\\par
4. Data Availability Statement -> \\i {options_article['data_phrasing']}\\i0\\par
5. Supplementary Material (Si applicable)\\par
6. References (Bibliographie au style \\b {config['style_citation']}\\b0)\\par
7. Citation Box (Encadré de citation officiel)\\par
8. Appendices / Annexes (En taille 9pt)\\par
{'- 9. S2O Box (À placer tout à la fin car le journal est en Subscribe to Open)' if options_article['is_s2o'] else ''}\\par
\\line
\\b 4. CONSIGNES BIBLIOGRAPHIQUES ({config['style_citation']})\\b0\\par
- Les titres des livres et des revues ne doivent \\b pas\\b0 prendre de majuscules (sauf la première lettre).\\par
- Aucun nom de revue ne doit commencer par l'article "The".\\par
- Conserver précieusement le numéro DOI si celui-ci est fourni.\\par
}}"""
    return texte_rtf.encode('utf-8')

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
        citation_box = "Cite this article as: P. Nom, \"Title of the article,\" " + config['header'] + ", vol. 22, pp. 78-82, 2026."
    else:
        citation_box = "Cite this article as: Nom, P. (2026). \"Title of the article.\" " + config['header'] + ", 22(1), 78-82."

    # Gestion propre des conditions sans imbriquer de f-strings complexes
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
\\noindent\\textbf{{Abstract – }} \\lipsum[1]

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


# 4. INTERFACE GRAPHIQUE
st.title("📚 Centre de Ressources Éditoriales")
st.caption("Génération automatisée des paquets d'instructions et des maquettes visuelles.")

onglet_typesetter, onglet_editeur = st.tabs(["🚀 Espace Téléchargement (Typesetters)", "⚙️ Configuration des Chartes (Éditeurs)"])

with onglet_typesetter:
    st.header("Paquet de Publication par Revue")
    st.write("Sélectionnez la revue et configurez l'article pour exporter vos documents d'accompagnement.")
    
    revue_choisie = st.selectbox("Sélectionner la revue :", list(st.session_state.revues.keys()))
    
    if revue_choisie:
        cfg = st.session_state.revues[revue_choisie]
        
        # Configuration rapide de l'article en cours
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
"funding_text": funding_text, "conflict_text": conflict_text, "data_phrasing": data_phrasing}st.write("---")st.subheader("📥 Documents à exporter")col_btn1, col_btn2 = st.columns(2)# EXPORT 1 : INSTRUCTIONS HUMAINES (.DOCX)data_docx = generer_instructions_docx(revue_choisie, cfg, options_article)col_btn1.download_button(label="📄 Télécharger la Liste d'Instructions (.docx)",data=data_docx,file_name=f"instructions_typesetter_{revue_choisie.lower().replace(' ', '_')}.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")# EXPORT 2 : GABARIT VISUEL (.TEX)data_tex = generer_visuel_latex(revue_choisie, cfg, options_article)col_btn2.download_button(label="🛠️ Télécharger l'Exemple de Code Visuel (.tex)",data=data_tex,file_name=f"exemple_visuel_{revue_choisie.lower().replace(' ', '_')}.tex",mime="text/plain")st.caption("💡 Note pour le Typesetter : Pour visualiser le rendu final sous forme de PDF, téléchargez le fichier .tex ci-dessus et compilez-le dans votre logiciel de traitement LaTeX habituel.")(L'onglet éditeur reste inchangé pour la gestion des revues sur le disque)with onglet_editeur:st.header("Édition des chartes graphiques par Revue")action = st.radio("Action :", ["Modifier une revue existante", "Créer une nouvelle revue"])if action == "Modifier une revue existante":revue_a_modifier = st.selectbox("Sélectionner la revue à éditer :", list(st.session_state.revues.keys()))cfg_actuelle = st.session_state.revues[revue_a_modifier]with st.form("form_edit"):police = st.text_input("Police", value=cfg_actuelle["police"])taille = st.text_input("Taille des titres", value=cfg_actuelle["taille_titre"])couleur = st.color_picker("Couleur de la revue", value=cfg_actuelle["couleur"])marges = st.text_input("Marges (ex: 2.5cm)", value=cfg_actuelle["marges"])header = st.text_input("Nom abrégé du journal (Header/IDLine)", value=cfg_actuelle["header"])id_line_format = st.selectbox("Format de l'IDLine", ["Pagination", "Article Number"], index=0 if cfg_actuelle["id_line_format"] == "Pagination" else 1)style_citation = st.selectbox("Style de Citation bibliographique", ["IEEE", "APA"], index=0 if cfg_actuelle["style_citation"] == "IEEE" else 1)open_access = st.checkbox("Revue obligatoirement en Open Access", value=cfg_actuelle["open_access"])deux_colonnes = st.checkbox("Mise en page sur deux colonnes", value=cfg_actuelle["deux_colonnes"])sections_numerotees = st.checkbox("Numérotation des sections en chiffres arabes", value=cfg_actuelle["sections_numerotees"])if st.form_submit_button("Enregistrer les règles éditoriales"):st.session_state.revues[revue_a_modifier] = {"police": police, "taille_titre": taille, "couleur": couleur, "marges": marges, "header": header,"id_line_format": id_line_format, "style_citation": style_citation, "open_access": open_access,"deux_colonnes": deux_colonnes, "sections_numerotees": sections_numerotees}sauvegarder_donnees(st.session_state.revues)st.success("Charte graphique mise à jour sur le disque.")st.rerun()
