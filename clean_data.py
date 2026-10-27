"""Nettoyage PROFOND EN PLACE : récupère un maximum de données propres.
  1) Nom pollué (« NOM / Num : 0612... / CD587170 ») -> extrait tel+cin, nom propre.
  2) Ville manquante -> déduite de l'adresse (villes MA + étrangères + CP).
  3) Adresse = en fait une CIN -> déplace vers cin, vide l'adresse.

    python clean_data.py           # aperçu
    python clean_data.py --apply
"""
import re
import sys
import store
import geo

_CIN = re.compile(r'\b([A-Z]{1,3}\d{4,8}[A-Z]?|\d{2}[A-Z]{1,2}\d{3,6})\b')
_TEL = re.compile(r'\b(0\d(?:[\s.\-]?\d\d){4})\b')
_CIN_LABEL = re.compile(r'\b(?:CD|CIN|CNI)[\s:]*([A-Z]{0,3}\d{4,8})\b', re.I)

# Villes étrangères fréquentes (complément de geo pour l'international).
_FOREIGN = {
    'malaga': 'Malaga', 'madrid': 'Madrid', 'barcelone': 'Barcelone', 'barcelona': 'Barcelone',
    'colombes': 'Colombes', 'stains': 'Stains', 'paris': 'Paris', 'lyon': 'Lyon', 'marseille': 'Marseille',
    'montpellier': 'Montpellier', 'argenteuil': 'Argenteuil', 'sartrouville': 'Sartrouville',
    'suresnes': 'Suresnes', 'rouen': 'Rouen', 'orleans': 'Orléans', 'longueuil': 'Longueuil',
    'geneve': 'Genève', 'genève': 'Genève', 'bruxelles': 'Bruxelles', 'anvers': 'Anvers',
    'gorinchem': 'Gorinchem', 'malmo': 'Malmö', 'terrassa': 'Terrassa', 'castellon': 'Castellón',
    'abu dhabi': 'Abu Dhabi', 'dubai': 'Dubaï', 'dubaï': 'Dubaï',
}


def _clean_name(name):
    """Retourne (nom_propre, cin_trouvee, tel_trouve)."""
    cin = tel = None
    m = _CIN_LABEL.search(name)
    if m:
        cin = (m.group(0).split()[0].upper().replace(':', '') if False else None)
        cin = re.sub(r'[\s:]', '', m.group(0)).upper()
    mt = _TEL.search(name)
    if mt:
        tel = re.sub(r'[\s.\-]', '', mt.group(1))
    # nom propre = segments avant « Num/Tèl », sans code CIN
    cut = re.split(r'(?i)\s*/?\s*(?:num|t[èe]l|cd|cin)\s*[:.]', name)[0]
    cut = _CIN.sub('', cut)
    cut = re.sub(r'\s*/\s*$', '', cut).strip(' /')
    return cut, cin, tel


def _find_ville(adr):
    if not adr:
        return None
    v = geo.ville_from_addr(adr) if hasattr(geo, 'ville_from_addr') else None
    if v:
        return v
    low = adr.lower()
    for k, lbl in _FOREIGN.items():
        if k in low:
            return lbl
    # CP français « 92700 COLOMBES FRANCE » -> mot(s) après le CP
    m = re.search(r'\b\d{5}\s+([A-ZÀ-Ÿ][A-Za-zÀ-ÿ\- ]{2,30}?)(?:\s+(?:FRANCE|ESPAGNE|BELGIQUE))?\s*$', adr.upper())
    if m:
        return m.group(1).title().strip()
    return None


def plan():
    cs = store.load_clients() or []
    out = []
    for c in cs:
        patch = {}
        notes = []
        name = (c.get('nom_prenom') or '').strip()

        # 1) nom pollué
        if re.search(r'(?i)num\s*[:.]|t[èe]l\s*[:.]|\b(?:CD|CIN)\s*\d', name):
            clean, cin, tel = _clean_name(name)
            if clean and clean != name:
                patch['nom_prenom'] = clean
                notes.append(f'nom -> {clean!r}')
            if cin and not c.get('cin'):
                patch['cin'] = [cin]
                patch['cin_str'] = cin
                notes.append(f'cin={cin}')
            if tel and not (c.get('tel') or '').strip():
                patch['tel'] = tel
                notes.append(f'tel={tel}')

        # 2) adresse = en fait une CIN
        adr = (patch.get('nom_prenom') and c.get('adresse1') or c.get('adresse1') or '').strip()
        if adr and _CIN.fullmatch(adr.strip()):
            if not c.get('cin'):
                patch['cin'] = [adr.strip().upper()]
                patch['cin_str'] = adr.strip().upper()
                notes.append(f'adresse était une CIN -> cin={adr}')
            patch['adresse1'] = ''
            notes.append('adresse vidée (c\'était une CIN)')

        # 3) ville manquante
        cur_ville = (c.get('ville') or '').strip()
        use_adr = patch.get('adresse1', c.get('adresse1')) or ''
        if not cur_ville and use_adr:
            v = _find_ville(use_adr)
            if v:
                patch['ville'] = v
                notes.append(f'ville={v}')
                if not (c.get('pays') or '').strip():
                    patch['pays'] = geo.pays_from(v, use_adr) if hasattr(geo, 'pays_from') else 'Maroc'

        if patch:
            out.append((c.get('id'), name, patch, notes))
    return out


def main():
    apply = '--apply' in sys.argv
    rows = plan()
    print(f"{len(rows)} fiches à nettoyer ({'APPLIQUE' if apply else 'APERCU'}):\n")
    for cid, name, patch, notes in rows:
        print(f"  [{cid}] {name[:38]}")
        for n in notes:
            print(f"       - {n}")
    if apply:
        for cid, name, patch, notes in rows:
            store.update_client(cid, patch)
        print(f"\n{len(rows)} fiches nettoyées dans la base active.")
    else:
        print("\n(aperçu — relancer avec --apply)")


if __name__ == '__main__':
    main()
