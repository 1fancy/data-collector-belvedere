"""Exporte les données enrichies vers un JSON plat que le pont de sync ERP lit.
Pas de modification de la base — lecture seule.

Sortie: sync_export_erp.json  (liste d'objets normalises)
"""
import json
import os
import re
import store

OUT = os.path.join(os.path.dirname(__file__), "sync_export_erp.json")


def _norm_num(v):
    if v is None:
        return None
    s = re.sub(r"[^0-9]", "", str(v))
    return s or None


def _first_cin(c):
    cin = c.get("cin")
    if isinstance(cin, list):
        return cin[0] if cin else None
    if isinstance(cin, str):
        return cin.strip() or None
    return None


def _split_name(c):
    """Retourne (prenom, nom) best-effort depuis nom_prenom."""
    full = (c.get("nom_prenom") or "").strip()
    if c.get("prenom") or c.get("nom"):
        return (c.get("prenom") or "").strip(), (c.get("nom") or "").strip()
    # on coupe au premier "/" (co-acquereur) et on prend le segment principal
    full = re.split(r"\s*/\s*", full)[0].strip()
    parts = full.split()
    if len(parts) >= 2:
        return parts[0], " ".join(parts[1:])
    return "", full


def build():
    clients = store.load_clients() or []
    rows = []
    for c in clients:
        if c.get("_noise") or c.get("exclude_from_stats"):
            continue
        name = (c.get("nom_prenom") or "").strip()
        if not name:
            continue
        prenom, nom = _split_name(c)
        suivi = c.get("suivi") or {}
        rows.append({
            "src_id": c.get("id"),
            # --- identite bien ---
            "ilot": str(c.get("ilot") or "").strip() or None,
            "immeuble": (c.get("immeuble") or "").strip() or None,
            "numero": _norm_num(c.get("numero") or c.get("num_lot")),
            "type_bien": (c.get("type_bien") or c.get("type") or "").strip() or None,
            "etage": c.get("etage"),
            "superficie": c.get("superficie"),
            "prix_total": c.get("prix_total") or c.get("prix"),
            # --- client ---
            "nom_complet": name,
            "prenom": prenom,
            "nom": nom,
            "cin": _first_cin(c),
            "cin_all": c.get("cin") if isinstance(c.get("cin"), list) else ([c.get("cin")] if c.get("cin") else []),
            "tel": (c.get("tel") or "").strip() or None,
            "email": (c.get("email") or "").strip() or None,
            "adresse": (c.get("adresse1") or "").strip() or None,
            "ville": (c.get("ville") or "").strip() or None,
            "pays": (c.get("pays") or "").strip() or None,
            # --- deal / statut ---
            "statut_lot": (suivi.get("statut_lot") or c.get("statut_lot") or "").strip() or None,
            "is_desistement": bool(c.get("is_desistement")),
            "mode_paiement": c.get("paiement_mode") or suivi.get("mode_paiement"),
            "banque": c.get("banque"),
            "rib": c.get("rib"),
            "avance": c.get("avance"),
            "reste": c.get("reste"),
            "tag": c.get("tag"),            # commercial(e)
            # --- suivi livraison (si present) ---
            "suivi": suivi,
        })
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"count": len(rows), "rows": rows}, f, ensure_ascii=False, indent=1)
    return len(rows)


if __name__ == "__main__":
    n = build()
    print(f"{n} lignes exportees -> {OUT}")
