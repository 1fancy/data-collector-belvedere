"""Enrichissement EN PLACE (pas de rebuild) : récupère CIN/téléphone coincés
dans le champ nom, et marque les non-clients (DE SISTEMENT, placeholders).

Usage:
    python enrich_names.py           # aperçu seulement (dry-run)
    python enrich_names.py --apply   # écrit dans la base active via store.update_client
"""
import re
import sys
import store

# CIN marocaine : 1-2 lettres + 4-8 chiffres (ex A633876, AB814972, 18HD12087 variante)
_CIN_RE = re.compile(r'\b([A-Z]{1,2}\d{4,8}|\d{2}[A-Z]{1,2}\d{3,6})\b')
_TEL_RE = re.compile(r'\b(0\d[\s.\-]?\d{2}[\s.\-]?\d{2}[\s.\-]?\d{2}[\s.\-]?\d{2})\b')
# "CD 587170" / "CD587170" collé après Tèl — on garde le préfixe lettre (CD587170)
_CIN_LABELED = re.compile(r'\b(CD|CIN|CNI)[\s:]*([A-Z]{0,2}\d{4,8})\b', re.I)

_NONCLIENT = re.compile(r'(?i)^\s*(de\s*sist[ée]?ment|desistement|r[ée]servation|rec\s*re\s*servation|test)\s*$')


def _clean_tel(t):
    return re.sub(r'[\s.\-]', '', t)


def plan():
    clients = store.load_clients() or []
    changes = []
    for c in clients:
        cid = c.get("id")
        name = (c.get("nom_prenom") or "").strip()
        if not name:
            continue
        patch = {}
        notes = []

        # 1) non-client / placeholder
        if _NONCLIENT.match(name):
            patch["_noise"] = True
            patch["exclude_from_stats"] = True
            notes.append("marque NON-CLIENT")
            changes.append((cid, name, patch, notes))
            continue

        has_cin = bool(c.get("cin"))
        has_tel = bool((c.get("tel") or "").strip())

        # 2) CIN coincé dans le nom (seulement si le nom contient un séparateur / ou :)
        if "/" in name or ":" in name:
            # labeled first (CD 587170)
            m = _CIN_LABELED.search(name)
            # "CD 587170" -> "CD587170" (préfixe collé au numéro)
            cin_found = (m.group(1).upper() + m.group(2)) if m else None
            if not cin_found:
                m2 = _CIN_RE.search(name.upper())
                cin_found = m2.group(1) if m2 else None
            if cin_found and not has_cin:
                patch.setdefault("cin", [cin_found])
                patch["cin_str"] = cin_found
                notes.append(f"CIN={cin_found} depuis nom")

            mt = _TEL_RE.search(name)
            if mt and not has_tel:
                patch["tel"] = _clean_tel(mt.group(1))
                notes.append(f"tel={patch['tel']} depuis nom")

            # nom nettoyé = premier segment avant le "/" qui n'est pas un numéro
            segs = [s.strip() for s in re.split(r'[/]', name) if s.strip()]
            clean = None
            for s in segs:
                alpha = re.sub(r'[^A-Za-zÀ-ÿ]', '', s)
                if len(alpha) >= 3 and not _CIN_RE.search(s.upper()) and 'tèl' not in s.lower() and 'tel' not in s.lower():
                    clean = s
                    break
            if clean and clean != name and (patch.get("cin") or patch.get("tel")):
                patch["nom_prenom"] = clean
                notes.append(f"nom nettoye -> {clean!r}")

        if patch:
            changes.append((cid, name, patch, notes))
    return changes


def main():
    apply = "--apply" in sys.argv
    changes = plan()
    print(f"{len(changes)} enregistrements a enrichir EN PLACE "
          f"({'APPLIQUE' if apply else 'APERCU'}):\n")
    for cid, name, patch, notes in changes:
        print(f"  [{cid}] {name!r}")
        for n in notes:
            print(f"       - {n}")
    if apply:
        for cid, name, patch, notes in changes:
            store.update_client(cid, patch)
        print(f"\n{len(changes)} enregistrements mis a jour dans la base active.")
    else:
        print("\n(aperçu — relancer avec --apply pour ecrire)")


if __name__ == "__main__":
    main()
