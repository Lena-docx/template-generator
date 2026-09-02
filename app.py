import streamlit as st

# Configuration de la page
st.set_page_config(page_title="Instructions de mise en page", layout="centered")
st.title("Instructions de mise en page pour le compositeur")

# Blocs de textes partagés pour éviter les répétitions
idline_vol_art = "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2026"
idline_vol_page = "[Journal shortened name], Volume(Issue), Page numbers([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22(1), 78-82, 2026"

text_copyright = """© [Name of the authors], Published by EDP Sciences, [Year] 
Example: J.M. Bertho and M. Bourguignon, Published by EDP Sciences, 2026
•	If there is one author : © [Initial + Surname], Published by EDP Sciences, [Year] 
•	If there are two authors: © [Initial + Surname and Initial + Surname], Published by EDP Sciences, [Year] 
•	If there are three or more authors: © [Initial + Surname et al.], Published by EDP Sciences, [Year]"""

text_open_access = """There must always be :
•	At the top of the page : the “Open Access” logo
•	At the bottom of the page : the mention “This is an Open Access article distributed under the terms of the Creative Commons Attribution License (https://creativecommons.org), which permits unrestricted use, distribution, and reproduction in any medium, provided the original work is properly cited. ”"""

# Dictionnaire optimisé des revues
revues = {
    "cagri": {"type": "art", "oa": True},
    "geotech": {"type": "art", "oa": False},
    "jbio": {"type": "page", "oa": False},
    "tpe": {"type": "art", "oa": False},
    "bsgf": {"type": "art", "oa": True},
    "limn": {"type": "art", "oa": True},
    "nss": {"type": "page", "oa": True},
    "parasite": {"type": "art", "oa": True},
    "pmed": {"type": "page", "oa": False},
    "radiopro": {"type": "page", "oa": True},
    "aacus": {"type": "art", "oa": True},
    "alr": {"type": "art", "oa": False, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2027"},
    "mfreview": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2028"},
    "ocl": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2029"},
    "rees": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2030"},
    "stet": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2031"},
    "kmae": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2032"},
    "mattech": {"type": "art", "oa": False, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2033"},
    "meca": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2034"},
    "metal": {"type": "art", "oa": False, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2035"},
    "emsci": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2036"},
    "ijmqe": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2037"},
    "jeos": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2038"},
    "rdne": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2039"},
    "sbuild": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2040"},
    "smdo": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2041"},
    "swsc": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2042"},
    "ject": {"type": "page", "oa": True},
    "sicotj": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2042"},
    "vcm": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2043"},
    "sands": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2044"},
    "esaim-cocv": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2044"},
    "esaim-m2an": {"type": "page", "oa": True},
    "esaim-ps": {"type": "page", "oa": True},
    "mmnp": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2044"},
    "rairo-ro": {"type": "page", "oa": True},
    "rairo-ita": {"type": "art", "oa": True, "custom_id": "[Journal shortened name], Volume, Article number ([Year])\nExample: J. Eur. Opt. Society-Rapid Publ., 22, 62, 2044"},
    "medsci": {"type": "page", "oa": False},
    "jomos": {"type": "art", "oa": True},
    "ppsy": {"type": "page", "oa": False}
}

# Menu déroulant
liste_revues = ["-- Sélectionnez une revue --"] + list(revues.keys())
choix = st.selectbox("Choisir une revue :", liste_revues)

# Affichage conditionnel des blocs
if choix != "-- Sélectionnez une revue --":
    config = revues[choix]
    
    # Choix de l'IDLine
    if "custom_id" in config:
        idline_text = config["custom_id"]
    else:
        idline_text = idline_vol_art if config["type"] == "art" else idline_vol_page

    # Bloc IDLine
    st.subheader("IDLine")
    st.text(idline_text)

    # Bloc Copyright
    st.subheader("Copyright")
    st.text(text_copyright)

    # Bloc Open Access
    st.subheader("Open Access")
    if config["oa"]:
        st.text(text_open_access)
    else:
        st.text("/")
