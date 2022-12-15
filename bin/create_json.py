##########################################################################################
# Filename:     create_json.py
# For:          Follow_Up_Creation_Tool
# Author:       Spyros Acheimastos (acheimastos@althom.eu)
# Date:         15/12/2022
##########################################################################################

import json

msn_R10_dict = {
    "all": [
        "0835",
        "0926",
        "0974",
        "1017",
        "1094",
        "1185",
        "1195",
        "1199",
        "1207",
        "1238",
        "1241",
        "1250",
        "1293",
        "1408",
        "1607",
        "1670",
        "1713",
        "1988",
        "1994",
        "2005",
        "2060",
        "2342",
        "2724",
        "2737",
        "2903",
        "2912",
        "3191",
        "3267",
        "3334",
        "3369",
        "3504",
        "3669",
        "3708",
        "3749",
        "5126",
        "5133"
    ],

    "all_A320": [
        "2724",
        "2737"
    ],

    "new": [
        "0926",
        "1094",
        "1185",
        "1670",
        "2060",
        "3669",
        "5133"
    ],

    "rev": [
        "1199",
        "1207",
        "1713"
    ]
}


authors_dict = {
    "IPC": [
        "D.Tsoukalas",
        "E.Koutelou",
        "G.Tassopoulos",
        "G.Ziogas",
        "P.Rezou",
        "S.Acheimastos",
        "S.Lagousi",
        "V.Menegatos",
        "V.Paliouras",
        "I.Moustakis",
        "Arty"
    ],

    "SRM": [
        "A.Geranios",
        "G.Frantzis",
        "P.Gianniotis"
    ],

    "ILLU": [
        "A.Apostolopoulou",
        "E.Makka",
        "E.Politi",
        "G.Tsakalozos",
        "G.Tzanakis",
        "A.Katsigialou"
    ]
}

def create_json_msns():
    """Create JSON with MSNs of R10 in case the file has been lost or as a template"""
    with open('INPUT_MSNs_sample.json', 'w') as file:
        json.dump(msn_R10_dict, file, indent=4)

def create_json_authors():
    """Create JSON with Authors of R10 in case the file has been lost or as a template"""
    with open('INPUT_AUTHORS_sample.json', 'w') as file:
        json.dump(authors_dict, file, indent=4)