"""Corrige EN PLACE les adresses étrangères marquées à tort pays=Maroc.
Le nom du pays est explicitement dans l'adresse -> on le prend comme source de
vérité (pattern simple, pas d'IA). Tente aussi de récupérer la ville étrangère.

    python fix_foreign_geo.py          # aperçu
    python fix_foreign_geo.py --apply  # écrit via store.update_client
"""
import re
import sys
import store

# nom pays (dans l'adresse) -> libellé normalisé
COUNTRIES = {
    'canada': 'Canada', 'france': 'France', 'belgique': 'Belgique', 'belgium': 'Belgique',
    'allemagne': 'Allemagne', 'germany': 'Allemagne', 'deutschland': 'Allemagne',
    'espagne': 'Espagne', 'spain': 'Espagne', 'españa': 'Espagne',
    'italie': 'Italie', 'italy': 'Italie', 'pays-bas': 'Pays-Bas', 'netherlands': 'Pays-Bas',
    'suisse': 'Suisse', 'switzerland': 'Suisse', 'usa': 'États-Unis',
    'united states': 'États-Unis', 'etats-unis': 'États-Unis',
    'uk': 'Royaume-Uni', 'angleterre': 'Royaume-Uni', 'royaume-uni': 'Royaume-Uni',
    'qatar': 'Qatar', 'emirats arabes unis': 'Émirats arabes unis', 'emirats': 'Émirats arabes unis',
    'dubai': 'Émirats arabes unis', 'abu dhabi': 'Émirats arabes unis',
    'norvege': 'Norvège', 'norway': 'Norvège', 'suede': 'Suède', 'sweden': 'Suède',
    'portugal': 'Portugal',
}
# villes étrangères connues pour récupération (sous-ensemble utile)
FOREIGN_CITIES = [
    'Abu Dhabi', 'Dubaï', 'Dubai', 'Montpellier', 'Colombes', 'Argenteuil',
    'Montigny Le Bretonneux', 'Theuville', 'Malmö', 'Malmo', 'Malaga', 'Castellon',
    'Terrassa', 'Barcelone', 'Orleans', 'Orléans', 'Genève', 'Geneve', 'Paris',
    'Longueuil', 'Montréal', 'Montreal', 'Toronto',
]
_CA_POSTAL = re.compile(r'\b([A-Z]\d[A-Z])\s?(\d[A-Z]\d)\b')


def _country_in(text):
    low = text.lower()
    for k, v in COUNTRIES.items():
        if re.search(r'\b' + re.escape(k) + r'\b', low):
            return v
    if _CA_POSTAL.search(text.upper()):
        return 'Canada'
    return None


def _city_in(text):
    low = text.lower()
    for city in FOREIGN_CITIES:
        if city.lower() in low:
            return city
    # Canada : "<VILLE> QC A1A1A1 CANADA" -> le mot avant le code province.
    m = re.search(r'\b([A-ZÀ-Ÿ][A-Za-zÀ-ÿ\-\' ]+?)\s+(?:QC|ON|BC|AB|QUEBEC|ONTARIO)\b', text.upper())
    if m:
        return m.group(1).title().strip()
    # France/Espagne : "<CP 5 chiffres> <VILLE> <PAYS>" -> mot(s) après le code postal.
    m = re.search(r'\b\d{4,5}\s+([A-ZÀ-Ÿ][A-Za-zÀ-ÿ\-\' ]{2,30}?)\s+(?:FRANCE|ESPAGNE|SUEDE|SUISSE|BELGIQUE)\b', text.upper())
    if m:
        return m.group(1).title().strip()
    return None


def plan():
    cs = store.load_clients() or []
    out = []
    for c in cs:
        adr = ' '.join(x for x in (c.get('adresse1'), c.get('adresse2'), c.get('ville')) if x)
        if not adr.strip():
            continue
        pays = (c.get('pays') or '').strip()
        hit = _country_in(adr)
        if not hit:
            continue
        cur_ville = (c.get('ville') or '').strip()
        patch = {}
        # 1) pays erroné -> corriger
        if pays.lower() != hit.lower():
            patch['pays'] = hit
        # 2) ville manquante (ou marocaine à tort) alors que pays étranger
        if not cur_ville or cur_ville.lower() in ('maroc', 'sala al jadida', 'salé', 'sale'):
            city = _city_in(adr)
            if city:
                patch['ville'] = city
        if patch:
            out.append((c.get('id'), c.get('nom_prenom'), pays, patch, adr[:55]))
    return out


def main():
    apply = '--apply' in sys.argv
    rows = plan()
    print(f"{len(rows)} adresses étrangères à corriger ({'APPLIQUE' if apply else 'APERCU'}):\n")
    for cid, nom, oldpays, patch, adr in rows:
        bits = []
        if 'pays' in patch:
            bits.append(f"pays {oldpays or '(vide)'} -> {patch['pays']}")
        if 'ville' in patch:
            bits.append(f"ville -> {patch['ville']}")
        print(f"  [{cid}] {nom}: " + " | ".join(bits))
        print(f"        {adr}")
    if apply:
        for cid, nom, oldpays, patch, adr in rows:
            store.update_client(cid, patch)
        print(f"\n{len(rows)} corrigés dans la base active.")
    else:
        print("\n(aperçu — relancer avec --apply)")


if __name__ == '__main__':
    main()
