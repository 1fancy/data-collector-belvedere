"""Classe piece_type (cin|passeport|carte_sejour|autre) sur chaque fiche
Belvédère, d'après le motif de la pièce. Une CIN marocaine (1-2 lettres +
3-7 chiffres) reste 'cin' même si le client vit à l'étranger.

    python classify_piece.py          # aperçu
    python classify_piece.py --apply
"""
import re
import sys
import store


def classify(raw):
    # Co-propriété : plusieurs pièces séparées par « / » -> on classe la 1re.
    raw = re.split(r'\s*/\s*', (raw or '').strip())[0]
    v = re.sub(r'\s', '', raw.upper())
    if v == '':
        return None
    if v in ('PASSEPORT', 'PASSPORT'):
        return 'passeport'
    if re.match(r'^\d{2}\.\d{2}\.\d{2}-\d{3}\.\d{2}$', v):   # registre belge
        return 'carte_sejour'
    if re.match(r'^[A-Z]{1,2}\d{3,7}$', v):                  # CIN marocaine
        return 'cin'
    if re.match(r'^[A-Z]{1,3}\d{3,7}[A-Z]$', v):             # ...lettre finale (golfe)
        return 'passeport'
    if re.match(r'^\d', v):                                  # commence par un chiffre
        return 'passeport'
    if re.match(r'^[A-Z0-9]{6,12}$', v) and re.search(r'\d', v) and re.search(r'[A-Z]', v) \
            and not re.match(r'^[A-Z]{1,2}\d+$', v):         # mélange atypique
        return 'passeport'
    if re.match(r'^\d{6,12}$', v):
        return 'passeport'
    return 'autre'


def main():
    apply = '--apply' in sys.argv
    cs = store.load_clients() or []
    types = {}
    changes = []
    for c in cs:
        cin = c.get('cin_str') or (c.get('cin') and c['cin'][0]) or ''
        t = classify(cin)
        if not t:
            continue
        types[t] = types.get(t, 0) + 1
        if c.get('piece_type') != t:
            changes.append((c.get('id'), c.get('nom_prenom'), cin, c.get('piece_type'), t))
    print('Répartition:', types)
    print('Changements:', len(changes))
    for cid, nom, cin, old, new in changes:
        if new != 'cin':
            print(f"  [{cid}] {(nom or '')[:28]:28} [{cin}] {old} -> {new}")
    if apply:
        for cid, nom, cin, old, new in changes:
            store.update_client(cid, {'piece_type': new})
        print(f"\n{len(changes)} types appliqués.")
    else:
        print("\n(aperçu — --apply)")


if __name__ == '__main__':
    main()
