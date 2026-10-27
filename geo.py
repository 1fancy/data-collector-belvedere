#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Déduction du PAYS d'un client.

Règle (demandée par l'utilisateur) : on se fie D'ABORD à la VILLE, pas à
l'adresse entière — car l'adresse contient parfois un mot trompeur
(ex. « GRP DOHA ELKHAIR » est un quartier de Sala Al Jadida, au Maroc, PAS Doha
au Qatar). On ne scanne l'adresse qu'en dernier recours, si la ville est vide.
"""

import re

# Ville (ou mot-clé de ville) -> pays étranger.
FOREIGN_CITY = {
    "ALLEMAGNE": "Allemagne", "DEUTSCHLAND": "Allemagne", "NEUSS": "Allemagne",
    "FRANKFURT": "Allemagne", "FRANCFORT": "Allemagne",
    "FRANCE": "France", "PALAISEAU": "France", "MALMAISON": "France",
    "VILLEURBANNE": "France", "VAULX": "France", "MERIGNAC": "France",
    "SANNOIS": "France", "VELIZY": "France", "VERSAILLES": "France",
    "COURBEVOIE": "France", "THONON": "France", "TOULOUSE": "France",
    "VESAILLES": "France", "VELIZY-VILLACOUBLAY": "France", "RUEIL": "France",
    "ORGE": "France", "VERNET": "France",
    "BELGIQUE": "Belgique", "BELGIE": "Belgique", "KORTRIJK": "Belgique",
    "COURTRAI": "Belgique", "ANS": "Belgique",
    "PAYS-BAS": "Pays-Bas", "NEDERLAND": "Pays-Bas", "AMSTERDAM": "Pays-Bas",
    "ALPHEN": "Pays-Bas",
    "QATAR": "Qatar", "DOHA": "Qatar",
    "KOWEIT": "Koweït", "KUWAIT": "Koweït", "SALMIYA": "Koweït",
    "NORWAY": "Norvège", "NORVEGE": "Norvège", "OSLO": "Norvège",
    "CANADA": "Canada", "REPENTIGNY": "Canada", "OTTAWA": "Canada",
    "ARABIE SAOUDITE": "Arabie Saoudite", "SAOUDITE": "Arabie Saoudite",
    "DJEDDAH": "Arabie Saoudite", "JEDDAH": "Arabie Saoudite",
    "CROATIE": "Croatie", "ZAGREB": "Croatie", "CROATIA": "Croatie",
    "MONTIGNY": "France", "BAGNOLET": "France", "MONTPELLIER": "France",
    "MASSY": "France", "EPINAY": "France", "LYON": "France", "MARSEILLE": "France",
    "LILLE": "France", "NANTES": "France", "BORDEAUX": "France", "STRASBOURG": "France",
    "NICE": "France", "RENNES": "France", "JUIN LE CANNIER": "France",
    "MARCHAL JUIN": "France", "SEINE": "France",
    "BARCELONE": "Espagne", "BARCELONA": "Espagne", "MADRID": "Espagne",
    "ESPAGNE": "Espagne", "VALENCIA": "Espagne", "VALENCE": "Espagne",
    "SAN DIEGO": "États-Unis", "ETAS UNIS": "États-Unis", "ETATS UNIS": "États-Unis",
    "USA": "États-Unis", "CALIFORNIA": "États-Unis", "NEW YORK": "États-Unis",
    "MALMO": "Suède", "SUEDE": "Suède", "STOCKHOLM": "Suède",
    "HAMBURG": "Allemagne", "HAMBOURG": "Allemagne", "BERLIN": "Allemagne",
    "MUNICH": "Allemagne", "DEUTSCHLAND": "Allemagne", "RAVENSBURG": "Allemagne",
    "BRUXELLES": "Belgique", "BRUSSELS": "Belgique", "LIEGE": "Belgique",
    "ANVERS": "Belgique", "GAND": "Belgique",
    "LONDON": "Royaume-Uni", "LONDRES": "Royaume-Uni", "MANCHESTER": "Royaume-Uni",
    "GENEVE": "Suisse", "GENEVA": "Suisse", "ZURICH": "Suisse", "SUISSE": "Suisse",
    "MONTREAL": "Canada", "QUEBEC": "Canada",
    "DAKAR": "Sénégal", "SENEGAL": "Sénégal", "SENAGAL": "Sénégal",
    "ROTTERDAM": "Pays-Bas", "LA HAYE": "Pays-Bas", "UTRECHT": "Pays-Bas",
    "DUBAI": "Émirats arabes unis", "ABU DHABI": "Émirats arabes unis",
    "EMIRATS": "Émirats arabes unis", "SHARJAH": "Émirats arabes unis",
    "RIYAD": "Arabie Saoudite", "RIYADH": "Arabie Saoudite", "DAMMAM": "Arabie Saoudite",
    "VIGO": "Espagne", "CASTELLON": "Espagne", "ALMASSORA": "Espagne",
    "LEVALLOIS": "France", "PARIS": "France", "THEUVILLE": "France",
    "ROSAY": "France", "CANNET": "France", "CANNIER": "France",
    "ABU DHABI": "Émirats arabes unis", "ABOU DHABI": "Émirats arabes unis",
    "AJMAN": "Émirats arabes unis",
    "KREFELD": "Allemagne", "DUISBURG": "Allemagne", "DUSSELDORF": "Allemagne",
    "KOLN": "Allemagne", "COLOGNE": "Allemagne", "STUTTGART": "Allemagne",
    "BONN": "Allemagne", "ESSEN": "Allemagne", "DORTMUND": "Allemagne",
    "ANTWERPEN": "Belgique", "GENT": "Belgique", "CHARLEROI": "Belgique",
    "TORONTO": "Canada", "LAVAL": "Canada", "GATINEAU": "Canada",
    "WASHINGTON": "États-Unis", "CHICAGO": "États-Unis", "HOUSTON": "États-Unis",
    "MIAMI": "États-Unis", "BOSTON": "États-Unis", "SEATTLE": "États-Unis",
}

# Marqueurs FORTS du Maroc : s'ils apparaissent, le pays est le Maroc.
MOROCCO_MARKERS = (
    "SALA AL JADIDA", "SALA EL JADIDA", "SALA ALJADIDA", "SALA JADIDA",
    "SALE", "SALÉ", "RABAT", "CASABLANCA", "CASA ", "MARRAKECH", "TANGER",
    "FES", "FÈS", "MEKNES", "MEKNÈS", "AGADIR", "KENITRA", "TEMARA",
    "TÉMARA", "MOHAMMEDIA", "OUJDA", "TETOUAN", "TÉTOUAN", "SOUISSI",
    "SOUSSI", "HARHOURA", "MAROC", "MOROCCO", "BOUZNIKA", "SKHIRAT",
)

# Pays acceptés tels quels (quand la ville EST déjà un nom de pays).
KNOWN_COUNTRIES = {
    "MAROC": "Maroc", "FRANCE": "France", "ALLEMAGNE": "Allemagne",
    "BELGIQUE": "Belgique", "PAYS-BAS": "Pays-Bas", "QATAR": "Qatar",
    "KOWEIT": "Koweït", "KOWEÏT": "Koweït", "NORVEGE": "Norvège",
    "NORVÈGE": "Norvège", "CANADA": "Canada", "CROATIE": "Croatie",
    "ARABIE SAOUDITE": "Arabie Saoudite", "ESPAGNE": "Espagne",
    "ITALIE": "Italie", "SUISSE": "Suisse",
}


# Villes connues (marocaines surtout) -> libellé d'affichage propre.
# L'ordre compte peu : on choisit la correspondance la PLUS À DROITE dans
# l'adresse (la ville est en général en fin de ligne).
CITY_DISPLAY = {
    "SALA AL JADIDA": "Sala Al Jadida", "SALA EL JADIDA": "Sala Al Jadida",
    "SALA ALJADIDA": "Sala Al Jadida", "SALA JADIDA": "Sala Al Jadida",
    "SALA AL JDIDA": "Sala Al Jadida", "SALE": "Salé", "SALÉ": "Salé",
    "SALI": "Salé",
    "RABAT": "Rabat", "CASABLANCA": "Casablanca", "CASA": "Casablanca",
    "MARRAKECH": "Marrakech", "TANGER": "Tanger", "FES": "Fès", "FÈS": "Fès",
    "MEKNES": "Meknès", "MEKNÈS": "Meknès", "AGADIR": "Agadir",
    "KENITRA": "Kénitra", "KÉNITRA": "Kénitra", "TEMARA": "Témara",
    "TÉMARA": "Témara", "MOHAMMEDIA": "Mohammedia", "OUJDA": "Oujda",
    "TETOUAN": "Tétouan", "TÉTOUAN": "Tétouan", "HARHOURA": "Harhoura",
    "BOUZNIKA": "Bouznika", "SKHIRAT": "Skhirat", "KHEMISSET": "Khémisset",
    "ROMMANI": "Rommani", "OUEZZANE": "Ouezzane", "LAAYOUN": "Laâyoune",
    "LAAYOUNE": "Laâyoune", "EL MARSA": "El Marsa", "NADOR": "Nador",
    "BENSLIMANE": "Benslimane", "SETTAT": "Settat", "BERRECHID": "Berrechid",
    "SAFI": "Safi", "ELJADIDA": "El Jadida", "EL JADIDA": "El Jadida",
    "DAKHLA": "Dakhla", "GUELMIM": "Guelmim", "ESSAOUIRA": "Essaouira",
    "BENI MELLAL": "Béni Mellal", "KHOURIBGA": "Khouribga",
    "TEROUAL": "Teroual", "AIN AOUDA": "Aïn Aouda", "SIDI SLIMANE": "Sidi Slimane",
    "MOHMMEDIA": "Mohammedia", "TAOUNATE": "Taounate", "LAAYAYDA": "Laâyayda",
    "KHENIFRA": "Khénifra", "MOULAY BOUAZZA": "Moulay Bouazza",
    "BEL KSIRI": "Bel Ksiri", "OUISLANE": "Ouislane", "GUERCIF": "Guercif",
    "BERKINE": "Berkine", "OUALIDIA": "Oualidia", "AZROU": "Azrou",
    "IFRANE": "Ifrane", "ERRACHIDIA": "Errachidia", "TAZA": "Taza",
    "LARACHE": "Larache", "KSAR EL KEBIR": "Ksar El Kébir", "INEZGANE": "Inezgane",
    "BERRADA": "Berrada", "YOUSSOUFIA": "Youssoufia", "SIDI KACEM": "Sidi Kacem",
    "SIDI BENNOUR": "Sidi Bennour", "OULED TEIMA": "Ouled Teima",
    "AIT MELLOUL": "Aït Melloul", "BOUSKOURA": "Bouskoura", "DEROUA": "Deroua",
    "SIDI RAHAL": "Sidi Rahal", "MARTIL": "Martil", "FNIDEQ": "Fnideq",
    "MDIQ": "M'diq", "CHEFCHAOUEN": "Chefchaouen", "AL HOCEIMA": "Al Hoceïma",
    "RABT": "Rabat", "TIFLET": "Tiflet", "GOULMIMA": "Goulmima",
    "SIDI ALLAL BAHRAOUI": "Sidi Allal Bahraoui", "KHMISSET": "Khémisset",
    "OULMES": "Oulmès", "ZAGORA": "Zagora", "TIZNIT": "Tiznit",
    "BENSLIMANE": "Benslimane", "BOUKNADEL": "Bouknadel", "AIN ATIQ": "Aïn Atiq",
    # villes étrangères fréquentes
    "NEUSS": "Neuss", "FRANKFURT": "Francfort", "VILLEURBANNE": "Villeurbanne",
    "PALAISEAU": "Palaiseau", "OSLO": "Oslo", "AMSTERDAM": "Amsterdam",
    "DOHA": "Doha", "ZAGREB": "Zagreb", "DJEDDAH": "Djeddah",
    "DAKAR": "Dakar", "RAVENSBURG": "Ravensburg", "KOWEIT": "Koweït City",
    "VIGO": "Vigo", "CASTELLON": "Castellón", "LEVALLOIS": "Levallois-Perret",
    "DUBAI": "Dubaï", "ALMASSORA": "Almassora",
    "ABU DHABI": "Abu Dhabi", "ABOU DHABI": "Abu Dhabi", "AJMAN": "Ajman",
    "KREFELD": "Krefeld", "DUISBURG": "Duisburg", "PARIS": "Paris",
    "OUARZAZATE": "Ouarzazate", "THEUVILLE": "Theuville", "ROSAY": "Rosay",
    "CANNET": "Le Cannet", "CANNIER": "Le Cannet", "DUSSELDORF": "Düsseldorf",
    "COLOGNE": "Cologne", "KOLN": "Cologne", "STUTTGART": "Stuttgart",
    "BONN": "Bonn", "ESSEN": "Essen", "DORTMUND": "Dortmund",
    "ANTWERPEN": "Anvers", "GENT": "Gand", "CHARLEROI": "Charleroi",
    "TORONTO": "Toronto", "LAVAL": "Laval", "GATINEAU": "Gatineau",
    "WASHINGTON": "Washington", "CHICAGO": "Chicago", "HOUSTON": "Houston",
    "MIAMI": "Miami", "BOSTON": "Boston", "SEATTLE": "Seattle",
    "SALMIYA": "Salmiya", "MONTIGNY": "Montigny-le-Bretonneux",
    "BAGNOLET": "Bagnolet", "MONTPELLIER": "Montpellier", "MASSY": "Massy",
    "EPINAY": "Épinay-sur-Seine", "BARCELONE": "Barcelone", "BARCELONA": "Barcelone",
    "SAN DIEGO": "San Diego", "MALMO": "Malmö", "HAMBURG": "Hambourg",
    "HAMBOURG": "Hambourg", "BRUXELLES": "Bruxelles", "LIEGE": "Liège",
    "MADRID": "Madrid", "VALENCIA": "Valence", "LONDON": "Londres",
    "LONDRES": "Londres", "BRUSSELS": "Bruxelles", "GENEVE": "Genève",
    "LYON": "Lyon", "MARSEILLE": "Marseille", "TOULOUSE": "Toulouse",
    "LILLE": "Lille", "NANTES": "Nantes", "BORDEAUX": "Bordeaux",
    "STRASBOURG": "Strasbourg", "NICE": "Nice", "RENNES": "Rennes",
    "COURBEVOIE": "Courbevoie", "VELIZY": "Vélizy", "VERSAILLES": "Versailles",
    "REPENTIGNY": "Repentigny", "OTTAWA": "Ottawa", "MONTREAL": "Montréal",
    "ALPHEN": "Alphen", "KORTRIJK": "Courtrai", "ANS": "Ans",
    "THONON": "Thonon-les-Bains", "MERIGNAC": "Mérignac", "SANNOIS": "Sannois",
    "VAULX": "Vaulx-en-Velin", "RUEIL": "Rueil-Malmaison",
}


def ville_from_addr(adresse):
    """Déduit la ville depuis l'adresse : on prend la ville CONNUE la plus à
    droite (la ville est généralement en fin d'adresse). Renvoie '' si rien.
    MAIS : « Sala Al Jadida » (3 mots) l'emporte sur « Salé » si les deux
    apparaissent (l'un contient l'autre spatialement)."""
    up = " " + (adresse or "").upper() + " "
    up = up.replace("É", "E").replace("È", "E")  # tolérant aux accents
    # « Sala Al Jadida » prime sur « Salé » si les deux mots-clés sont présents
    # (SALE est souvent juste la préfecture après l'adresse de Sala Al Jadida).
    if "SALA" in up and ("JADIDA" in up or "JDIDA" in up):
        return "Sala Al Jadida"
    best, pos = "", -1
    for kw, disp in CITY_DISPLAY.items():
        k = kw.replace("É", "E").replace("È", "E")
        p = up.rfind(" " + k)
        if p < 0:
            # tolère la ville collée (ex. « 1086HSSAIN SALA ... »)
            p = up.rfind(k)
            if p < 0:
                continue
        # exiger une frontière de mot après
        end = p + (1 if up[p] == " " else 0) + len(k)
        if end <= len(up) and (end == len(up) or not up[end].isalpha()):
            if p > pos:
                best, pos = disp, p
    return best


def _has_marker(up, markers):
    return any(m in up for m in markers)


def _scan_foreign(up):
    """Cherche un mot-clé de ville étrangère dans la chaîne majuscule `up`."""
    best, pos = "", -1
    for kw, country in FOREIGN_CITY.items():
        p = up.rfind(kw)
        if p > pos and re.search(r"(?<![A-Z])" + re.escape(kw), up):
            best, pos = country, p
    return best


def pays_from(ville="", adresse=""):
    """Déduit le pays. PRIORITÉ À LA VILLE, repli sur l'adresse.

    1. Ville « X (Pays) »          -> Pays.
    2. Ville = marqueur Maroc      -> Maroc.
    3. Ville = nom de pays connu   -> ce pays.
    4. Ville = ville étrangère     -> son pays.
    5. Ville vide -> on scanne l'adresse (marqueur Maroc, puis ville étrangère).
    6. Défaut : Maroc.
    """
    v = (ville or "").strip()
    vup = v.upper()

    if v:
        # 1. « Ville (Pays) »
        m = re.search(r"\(([^)]+)\)", v)
        if m:
            inside = m.group(1).strip()
            iu = inside.upper()
            if iu in KNOWN_COUNTRIES:
                return KNOWN_COUNTRIES[iu]
            if _has_marker(iu, MOROCCO_MARKERS):
                return "Maroc"
            return inside
        # 2. marqueur Maroc dans la ville
        if _has_marker(vup, MOROCCO_MARKERS):
            return "Maroc"
        # 3. la ville EST un pays connu
        if vup in KNOWN_COUNTRIES:
            return KNOWN_COUNTRIES[vup]
        # 4. ville étrangère connue
        f = _scan_foreign(" " + vup + " ")
        if f:
            return f
        # ville renseignée mais inconnue -> on considère le Maroc par défaut
        return "Maroc"

    # 5. pas de ville : on examine l'adresse
    aup = " " + (adresse or "").upper() + " "
    if _has_marker(aup, MOROCCO_MARKERS):
        return "Maroc"
    f = _scan_foreign(aup)
    if f:
        return f
    # 6. défaut
    return "Maroc"
