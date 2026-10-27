"""Récupère EN PLACE l'adresse depuis la carte CIN déjà scannée, pour les
clients qui ont un CIN mais pas d'adresse. Pas de rebuild : on relit juste la
pièce CIN de chaque dossier et on applique parse_cin_card + geo.

Usage:
    python enrich_addresses.py           # aperçu (dry-run)
    python enrich_addresses.py --apply   # ecrit via store.update_client
"""
import os
import sys
import store
import extract
import geo
try:
    import ocr
except Exception:
    ocr = None


def _cin_docs(c):
    docs = c.get("documents") or []
    cins = [d.get("rel") for d in docs if d.get("type") == "cin" and d.get("rel")]
    # repli : tout doc dont le nom evoque une CIN
    if not cins:
        for d in docs:
            n = (d.get("name") or "").lower()
            if ("cin" in n or "cni" in n or "ident" in n) and d.get("rel"):
                cins.append(d["rel"])
    return [p for p in cins if p and os.path.exists(p)]


def plan():
    clients = store.load_clients() or []
    todo = [c for c in clients
            if c.get("cin") and not (c.get("adresse1") or "").strip()]
    results = []
    for c in todo:
        cid = c.get("id")
        name = c.get("nom_prenom")
        paths = _cin_docs(c)
        if not paths:
            results.append((cid, name, None, None, "aucune piece CIN sur disque"))
            continue
        adr = ville = None; src = None
        for p in paths:
            try:
                text = ocr.ocr_file(p, max_pages=2) if ocr else ""
            except Exception as e:
                text = ""
            if not text:
                continue
            card = extract.parse_cin_card(text)
            if card.get("adresse"):
                adr = card["adresse"].strip()
                src = os.path.basename(p)
                ville = geo.ville_from_addr(adr) if hasattr(geo, "ville_from_addr") else None
                break
        if adr:
            results.append((cid, name, adr, ville, src))
        else:
            results.append((cid, name, None, None, "CIN relue, pas d'adresse lisible"))
    return results


def main():
    apply = "--apply" in sys.argv
    res = plan()
    got = [r for r in res if r[2]]
    print(f"{len(res)} clients CIN-sans-adresse examines | "
          f"{len(got)} adresses recuperees "
          f"({'APPLIQUE' if apply else 'APERCU'})\n")
    for cid, name, adr, ville, info in res:
        if adr:
            print(f"  [OK] [{cid}] {name}")
            print(f"       adresse: {adr!r}" + (f"  (ville={ville})" if ville else ""))
            print(f"       source : {info}")
        else:
            print(f"  [--] [{cid}] {name}  -> {info}")
    if apply:
        n = 0
        for cid, name, adr, ville, info in res:
            if adr:
                patch = {"adresse1": adr, "cin_adresse_src": info}
                if ville:
                    patch["ville"] = ville
                    patch["pays"] = geo.pays_from(ville, adr) if hasattr(geo, "pays_from") else None
                store.update_client(cid, patch)
                n += 1
        print(f"\n{n} adresses ecrites dans la base active.")
    else:
        print("\n(aperçu — relancer avec --apply pour ecrire)")


if __name__ == "__main__":
    main()
