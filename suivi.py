#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lecture du nouveau classeur « Suivi des livraisons » (Belvédère Ilot 1).

Ce classeur suit TOUS les lots (251) d'un îlot, avec 53 colonnes couvrant le
parcours de livraison en 7 étapes (voir le schéma fonctionnel) :
  1. Contact initial      2. Visite bureau       3. Visite technique
  4. Levée des réserves   5. PV de réception     6. Transmission notaire
  7. Remise des clés

On range :
  - les champs « identité du lot / client » au niveau racine du client
    (compatibles avec le reste de l'app : ilot, immeuble, numero, nom_prenom…) ;
  - TOUS les champs de suivi (dates, statuts d'étapes…) dans un sous-dict
    `suivi` (clé -> valeur), pour alimenter le workflow de livraison.
"""

import os
import re

try:
    import openpyxl
except Exception:
    openpyxl = None

# Intitulé exact (normalisé) de la colonne -> clé interne.
# La clé racine (identité) utilise les mêmes noms que le reste de l'app ;
# les clés de suivi sont préfixées par leur étape.
HEADER_MAP = {
    # --- identité du lot / client (racine) ---
    "ilot": "ilot",
    "immeuble": "immeuble",
    "niveau": "etage",
    "n appartement": "numero",
    "n commerce": "numero",
    "n lot": "numero",
    "n local": "numero",
    "statut du lot": "statut_lot",
    "superficie habitable definitive": "superficie",
    "nom et prenom client": "nom_prenom",
    "telephone": "tel",
    "date reservation": "date_vente",
    "commercial": "commercial",
    "avance de reservation": "avance",
    "prix global": "prix_total",
    "reste a payer": "reste",
    "mode de paiement": "mode_paiement",
    "paiement du solde aupres de": "solde_aupres",
    "date de paiement du solde": "solde_date",
    "client mre": "mre",
    "statut global": "statut_global",
    "prochaine action": "prochaine_action",
    # --- étape 1 : contact initial ---
    "premier appel": "s1_appel",
    "date 1er appel": "s1_appel_date",
    "joignable": "s1_joignable",
    "relance": "s1_relance",
    "date relance": "s1_relance_date",
    "joignable a la relance": "s1_relance_joignable",
    "lettre recommandee envoyee": "s1_lr_envoyee",
    "date envoi lettre": "s1_lr_date",
    "statut lettre recommandee": "s1_lr_statut",
    # --- étape 2 : visite bureau ---
    "visite bureau": "s2_visite",
    "date visite bureau": "s2_date",
    "remarque": "s2_remarque",
    # --- étape 3 : visite technique ---
    "rdv visite technique donne": "s3_rdv",
    "date prevue visite technique": "s3_date_prevue",
    "rdv respecte": "s3_respecte",
    "date visite technique si reportee": "s3_date_reportee",
    "date reelle visite technique": "s3_date_reelle",
    # --- étape 4 : levée des réserves ---
    "delai prevu de levee des reserves jours": "s4_delai_jours",
    "visite de levee de reserve": "s4_visite",
    "date visite de levee": "s4_date",
    # --- étape 5 : PV de réception ---
    "pv de reception": "s5_pv",
    "date pv de reception": "s5_date",
    # --- étape 6 : transmission notaire ---
    "fiche notaire envoyee a nezha": "s6_fiche_nezha",
    "date envoi a nezha": "s6_nezha_date",
    "fiche envoyee au notaire par nezha": "s6_notaire",
    "passage chez le notaire": "s6_passage",
    "compromis de vente": "s6_compromis",
    "date signature compromis": "s6_compromis_date",
    "contrat definitif": "s6_contrat",
    "date contrat definitif": "s6_contrat_date",
    "credit accorde": "s6_credit",
    "cheque cdg recu": "s6_cheque_cdg",
    # --- étape 7 : remise des clés ---
    "pv de remise des cles signe": "s7_pv_cles",
    "observations generales": "observation",
}

# Quels champs vont à la RACINE du client (le reste va dans `suivi`).
ROOT_FIELDS = {
    "ilot", "immeuble", "etage", "numero", "statut_lot", "superficie",
    "nom_prenom", "tel", "date_vente", "commercial", "avance", "prix_total",
    "reste", "observation",
}


def _norm(s):
    s = str(s or "").strip().lower()
    import unicodedata
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _find_header_row(ws, max_scan=8):
    """Trouve la ligne d'en-tête : celle qui contient « nom et prenom client »
    (ou à défaut la ligne la plus « remplie » dans les premières lignes)."""
    best_row, best_score = None, -1
    for r in range(1, min(max_scan, ws.max_row) + 1):
        hits = 0
        for c in range(1, ws.max_column + 1):
            if _norm(ws.cell(r, c).value) in HEADER_MAP:
                hits += 1
        if hits > best_score:
            best_score, best_row = hits, r
    return best_row if best_score >= 3 else None


def looks_like_suivi(path):
    """Vrai si le classeur ressemble au Suivi des livraisons (en-tête reconnu)."""
    if not openpyxl:
        return False
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return False
    found = False
    for ws in wb.worksheets:
        hr = _find_header_row(ws)
        if hr:
            found = True
            break
    wb.close()
    return found


def _norm_num(v):
    """Montant/texte -> chaîne simple sans « .0 » parasite."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _clean_numero(v):
    """« A2 » / « A 12 » / « 12 » -> « 12 » (on retire le préfixe immeuble)."""
    s = str(v or "").strip().upper()
    m = re.search(r"(\d+)", s)
    return m.group(1) if m else s


def _clean_ilot(v):
    """« Îlot 1 » -> « 1 »."""
    s = str(v or "").strip()
    m = re.search(r"(\d+)", s)
    return m.group(1) if m else s


def _clean_superficie(v):
    """« 98 » / « 41m² » -> « 98 » (nombre seul, l'unité est ajoutée à l'affichage)."""
    s = str(v or "").strip()
    m = re.search(r"(\d+(?:[.,]\d+)?)", s)
    return m.group(1).replace(",", ".") if m else s


def _clean_date(v):
    """« 2025-09-30 00:00:00 » -> « 30/09/2025 »."""
    s = str(v or "").strip()
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    return s


def read_suivi(path):
    """Lit le classeur Suivi et renvoie une liste de clients (dicts).

    Chaque client : champs racine + `suivi` (dict des étapes) + xl_sheet/xl_row
    + colmap de suivi pour une éventuelle réécriture ultérieure.
    """
    if not openpyxl:
        raise RuntimeError("openpyxl manquant.")
    wb = openpyxl.load_workbook(path, data_only=True)
    base = os.path.basename(path)
    out = []
    for ws in wb.worksheets:
        hr = _find_header_row(ws)
        if not hr:
            continue
        # colonne -> clé
        colkey = {}
        for c in range(1, ws.max_column + 1):
            k = HEADER_MAP.get(_norm(ws.cell(hr, c).value))
            if k and k not in colkey.values():
                colkey[c] = k
        for r in range(hr + 1, ws.max_row + 1):
            def cell(key):
                for c, kk in colkey.items():
                    if kk == key:
                        return _norm_num(ws.cell(r, c).value)
                return ""
            numero = _clean_numero(cell("numero"))
            nom = cell("nom_prenom").strip()
            # ligne vide (pas de n° de lot) -> on saute
            if not numero and not nom:
                continue
            immeuble = cell("immeuble").strip().upper()
            is_commerce = "commerce" in ws.title.lower()
            rec = {
                "ilot": _clean_ilot(cell("ilot")),
                "immeuble": immeuble,
                "etage": cell("etage"), "numero": numero,
                "type_bien": "Commerce" if is_commerce else "Appartement",
                "nom_prenom": nom, "tel": cell("tel"),
                "superficie": _clean_superficie(cell("superficie")),
                "prix_total": cell("prix_total"), "avance": cell("avance"),
                "reste": cell("reste"), "date_vente": _clean_date(cell("date_vente")),
                "commercial": cell("commercial").strip(),
                "observation": cell("observation"),
                "statut_lot": cell("statut_lot"),
            }
            # bloc suivi (toutes les clés non-racine)
            suivi = {}
            for c, kk in colkey.items():
                if kk in ROOT_FIELDS:
                    continue
                v = _norm_num(ws.cell(r, c).value)
                if v != "":
                    suivi[kk] = v
            rec["suivi"] = suivi
            rec["xl_sheet"] = ws.title
            rec["xl_row"] = r
            rec["xl_header_row"] = hr
            rec["xl_colkey"] = {str(c): k for c, k in colkey.items()}
            out.append(rec)
    wb.close()
    return out


# ---------------------------------------------------------------- fusion / base
def _norm_name(s):
    import unicodedata
    s = unicodedata.normalize("NFKD", (s or "")).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-z ]", " ", s.lower())
    return " ".join(sorted(t for t in s.split() if len(t) > 1))


def _is_commerce(c):
    t = (c.get("type_bien", "") or "").lower()
    return "commerce" in t or "magasin" in t


# Champs d'enrichissement repris de l'ancienne extraction (CIN, adresse, docs…).
_ENRICH_FIELDS = (
    "cin", "cin_str", "cin_verifiee", "cin_verified", "cin_extras",
    "quote_part", "adresse1", "adresse2", "ville", "pays",
    "naissance", "cin_validite", "pere", "mere",
    "recu_num", "recu_montant", "recu_date",
    "documents", "provenance", "source_fiche", "sources",
    "noms", "date_courrier",
)


def merge_old_enrichment(new_clients, old_clients):
    """Ré-attache les données extraites (CIN/adresse/docs) de l'ancienne base
    sur les lots du nouveau Suivi.

    - Appartements : correspondance par îlot+immeuble+n°.
    - Commerces    : correspondance par NOM du client (numérotation différente).
    Renvoie (n_matched_key, n_matched_name).
    """
    old_by_key = {}
    old_by_name = {}
    for c in old_clients:
        if not _is_commerce(c):
            k = (str(c.get("ilot", "")).strip(),
                 str(c.get("immeuble", "")).strip().upper(),
                 str(c.get("numero", "")).strip())
            old_by_key.setdefault(k, c)
        nm = _norm_name(c.get("nom_prenom", ""))
        if nm:
            old_by_name.setdefault(nm, c)

    n_key = n_name = 0
    for n in new_clients:
        hit = None
        if n.get("type_bien") == "Appartement":
            k = (str(n.get("ilot", "")).strip(),
                 str(n.get("immeuble", "")).strip().upper(),
                 str(n.get("numero", "")).strip())
            hit = old_by_key.get(k)
            if hit:
                n_key += 1
        if not hit:
            nm = _norm_name(n.get("nom_prenom", ""))
            if nm:                       # jamais de match sur un nom vide
                hit = old_by_name.get(nm)
                if hit:
                    n_name += 1
        if not hit:
            continue
        for f in _ENRICH_FIELDS:
            if f in hit and hit[f] not in (None, "", [], {}):
                # ne pas écraser une valeur déjà présente dans le Suivi
                if n.get(f) in (None, "", [], {}):
                    n[f] = hit[f]
        # Si le Suivi n'avait pas de nom mais l'ancienne base oui -> vendu.
        old_nom = (hit.get("nom_prenom") or "").strip()
        if old_nom and not (n.get("nom_prenom") or "").strip():
            n["nom_prenom"] = old_nom
            n["noms"] = hit.get("noms") or [old_nom]
            n["statut"] = "Vendu"
            n["statut_lot"] = "Vendu"
            sv = n.get("suivi") or {}
            sv["statut_lot"] = "Vendu"
            if "renseigner" in (sv.get("statut_global") or "").lower() or "libre" in (sv.get("statut_global") or "").lower() or not sv.get("statut_global"):
                sv["statut_global"] = "À contacter"
                sv["prochaine_action"] = sv.get("prochaine_action") or "Appeler le client"
            n["suivi"] = sv
        n["_enriched_from_old"] = True
    return n_key, n_name


# Champs repris d'un scan de dossier (fiche + OCR) vers la base Suivi.
_SCAN_FIELDS = (
    "cin", "cin_str", "cin_verifiee", "cin_verified", "cin_extras",
    "quote_part", "adresse1", "adresse2", "ville", "pays", "tel",
    "etage",
    "naissance", "cin_validite", "pere", "mere",
    "recu_num", "recu_montant", "recu_date",
    "paiement_montant", "paiement_mode", "banque", "rib",
    "documents", "provenance", "source_fiche", "source_folder", "sources",
    "noms", "superficie", "date_courrier",
)


def merge_scan(base_clients, scanned, set_tag=True, fill_only=True, add_new=False):
    """Enrichit la base Suivi avec les résultats d'un scan de dossiers.

    - Appartements : correspondance par îlot+immeuble+n°.
    - Commerces    : par n° (îlot+n°+type) puis par NOM.
    - Pose le `tag` (agent) depuis le dossier scanné.
    - `fill_only` : ne remplit que les champs vides de la base.
    - `add_new`   : AJOUTE les dossiers scannés sans correspondance (utile pour
      les îlots absents du Suivi, ex. îlots 2-5). `base_clients` est modifié en
      place (append).
    Renvoie {matched, tagged, new_cin, added}.
    """
    # index base
    by_key = {}
    by_name = {}
    for c in base_clients:
        k = (str(c.get("ilot", "")).strip(),
             str(c.get("immeuble", "")).strip().upper(),
             str(c.get("numero", "")).strip(),
             "Commerce" if _is_commerce(c) else "Appartement")
        by_key[k] = c
        # clé commerce sans immeuble (le Suivi a l'immeuble vide)
        if _is_commerce(c):
            by_key.setdefault((str(c.get("ilot", "")).strip(), "",
                               str(c.get("numero", "")).strip(), "Commerce"), c)
        nm = _norm_name(c.get("nom_prenom", ""))
        if nm:
            by_name.setdefault(nm, c)

    matched = tagged = new_cin = added = 0
    conflicts = []
    # désistements déjà présents (clé lot+nom) -> éviter les doublons multi-passes
    desist_keys = set()
    for c in base_clients:
        if c.get("is_desistement"):
            desist_keys.add((str(c.get("ilot", "")).strip(),
                             str(c.get("immeuble", "")).strip().upper(),
                             str(c.get("numero", "")).strip(),
                             "Commerce" if _is_commerce(c) else "Appartement",
                             _norm_name(c.get("nom_prenom", ""))))
    added_keys = set()
    for s in scanned:
        typ = "Commerce" if _is_commerce(s) else "Appartement"
        il = str(s.get("ilot", "")).strip()
        im = str(s.get("immeuble", "")).strip().upper()
        nu = str(s.get("numero", "")).strip()
        k = (il, im, nu, typ)
        # Une clé COMPLÈTE (îlot+immeuble+n°) identifie un lot de façon unique.
        complete_key = bool(il and im and nu) if typ == "Appartement" else bool(il and nu)
        hit = by_key.get(k)
        if not hit and typ == "Commerce":
            hit = by_key.get((il, "", nu, "Commerce"))
        # Correspondance par NOM seulement si la clé est INCOMPLÈTE (sinon chaque
        # lot distinct reste distinct, même si le client possède plusieurs lots).
        if not hit and not complete_key:
            nm = _norm_name(s.get("nom_prenom", ""))
            if nm:
                hit = by_name.get(nm)
        if not hit:
            # Clé nouvelle -> on AJOUTE le lot (même si le nom existe déjà
            # ailleurs : une personne peut détenir plusieurs lots).
            if add_new and (s.get("nom_prenom") or "").strip() and k not in added_keys:
                rec = _rec_from_scan(s)
                base_clients.append(rec)
                by_key[k] = rec
                if typ == "Commerce":
                    by_key.setdefault((il, "", nu, "Commerce"), rec)
                added_keys.add(k)
                nm = _norm_name(rec.get("nom_prenom", ""))
                if nm:
                    by_name.setdefault(nm, rec)
                added += 1
            continue
        matched += 1
        # CONFLIT : le lot a déjà un NOM VRAIMENT différent du dossier scanné.
        # On ignore les variantes (ordre des mots, co-acquéreur, troncature) :
        # un conflit réel = AUCUN prénom/nom significatif en commun.
        scan_nm = _norm_name(s.get("nom_prenom", ""))
        hit_nm = _norm_name(hit.get("nom_prenom", ""))
        scan_tok = set(scan_nm.split())
        hit_tok = set(hit_nm.split())
        shared = scan_tok & hit_tok
        # bruit connu (labels usufruit/nue-propriété mal nettoyés) -> pas un conflit
        noise = {"la", "nue", "propriete", "usufruit", "pour", "le", "les", "de"}
        real_conflict = (scan_nm and hit_nm and not shared
                         and not (scan_tok <= noise or hit_tok <= noise))
        if real_conflict:
            conflicts.append({
                "ilot": hit.get("ilot"), "immeuble": hit.get("immeuble"),
                "numero": hit.get("numero"), "type_bien": hit.get("type_bien"),
                "excel_nom": hit.get("nom_prenom"), "dossier_nom": s.get("nom_prenom"),
                "dossier_cin": s.get("cin_str", ""),
                "dossier_folder": s.get("source_folder", ""),
            })
            # Règle : le NOM DU CLASSEUR reste la vente active (Vendu). Le nom
            # DU DOSSIER devient un DÉSISTEMENT distinct sur le même lot, pour
            # qu'il reste visible dans le total (on n'écrase rien).
            # Anti-doublon (passes multiples) : ne pas recréer un désistement
            # déjà présent pour ce lot + ce nom.
            dk = (il, im, nu, typ, scan_nm)
            if dk in desist_keys:
                continue
            desist_keys.add(dk)
            des = _rec_from_scan(s)
            des["statut"] = "Désistement"
            des["statut_lot"] = "Désistement"
            des["is_desistement"] = True
            des["is_duplicate"] = True          # même lot qu'un autre client
            des["conflict_with"] = hit.get("nom_prenom", "")
            sv = des.get("suivi") or {}
            sv["statut_lot"] = "Désistement"
            sv["statut_global"] = "Désistement"
            sv["prochaine_action"] = ""
            des["suivi"] = sv
            base_clients.append(des)
            added += 1
            continue
        had_cin = bool(hit.get("cin"))
        for f in _SCAN_FIELDS:
            if f not in s or s[f] in (None, "", [], {}):
                continue
            if fill_only and hit.get(f) not in (None, "", [], {}):
                continue
            hit[f] = s[f]
        # Un dossier physique existe => le lot EST vendu. Si le Suivi n'avait pas
        # de nom (lot « à renseigner » / libre), on le renseigne depuis la fiche.
        scan_nom = (s.get("nom_prenom") or "").strip()
        if scan_nom and not (hit.get("nom_prenom") or "").strip():
            hit["nom_prenom"] = scan_nom
            hit["noms"] = s.get("noms") or [scan_nom]
            if s.get("quote_part"):
                hit["quote_part"] = s["quote_part"]
            hit["statut"] = "Vendu"
            hit["statut_lot"] = "Vendu"
            sv = hit.get("suivi") or {}
            sv["statut_lot"] = "Vendu"
            # si le parcours était « à renseigner », passer à « à contacter »
            if not sv.get("statut_global") or "renseigner" in (sv.get("statut_global") or "").lower() or "libre" in (sv.get("statut_global") or "").lower():
                sv["statut_global"] = "À contacter"
                sv["prochaine_action"] = sv.get("prochaine_action") or "Appeler le client"
            hit["suivi"] = sv
        if not had_cin and hit.get("cin"):
            new_cin += 1
        if set_tag and s.get("tag") and not hit.get("tag"):
            hit["tag"] = s["tag"]
            tagged += 1
        elif set_tag and s.get("tag"):
            hit["tag"] = hit.get("tag") or s["tag"]
    # rafraîchir cin_str + dériver ville/pays depuis l'adresse (si vides).
    try:
        import geo
    except Exception:
        geo = None
    for c in base_clients:
        c["cin_str"] = c.get("cin_str") or " / ".join(c.get("cin", []))
        if geo:
            adr = c.get("adresse1", "") or ""
            if adr and not (c.get("ville", "") or "").strip():
                v = geo.ville_from_addr(adr)
                if v:
                    c["ville"] = v
            if (c.get("ville") or adr) and not (c.get("pays", "") or "").strip():
                c["pays"] = geo.pays_from(c.get("ville", ""), adr)
    return {"matched": matched, "tagged": tagged, "new_cin": new_cin,
            "added": added, "conflicts": conflicts}


def _statut_of(s):
    """Statut canonique d'un dossier scanné : Désistement / Test / Vendu."""
    folder = (s.get("source_folder", "") or "").lower()
    st = (s.get("statut", "") or "")
    nom = (s.get("nom_prenom", "") or "")
    if "désist" in st.lower() or "desist" in folder or "désist" in folder \
            or "annul" in folder:
        return "Désistement"
    if re.search(r"(?i)\b(test|wafaa\s+zine|placeholder|essai|xxx)\b", nom):
        return "Test"
    return "Vendu"


def _rec_from_scan(s):
    """Construit un enregistrement client complet depuis un dossier scanné
    (pour les îlots absents du Suivi). Mêmes champs qu'un lot du Suivi."""
    nom = (s.get("nom_prenom") or "").strip()
    noms = s.get("noms") or ([nom] if nom else [])
    rec = dict(
        ilot=str(s.get("ilot", "")), immeuble=str(s.get("immeuble", "")).upper(),
        numero=str(s.get("numero", "")), type_bien=s.get("type_bien", "Appartement"),
        etage=s.get("etage", ""), nom_prenom=nom, noms=noms,
        cin=s.get("cin", []), cin_str=s.get("cin_str", ""),
        cin_verifiee=s.get("cin_verifiee", ""), cin_verified=s.get("cin_verified", []),
        cin_extras=s.get("cin_extras", []),
        quote_part=s.get("quote_part", ["100"] if noms else []),
        adresse1=s.get("adresse1", ""), adresse2=s.get("adresse2", ""),
        ville=s.get("ville", ""), pays=s.get("pays", ""),
        tel=s.get("tel", ""), superficie=s.get("superficie", ""),
        prix_total=s.get("prix_total", ""), avance=s.get("avance", ""),
        reste=s.get("reste", ""), date_vente=s.get("date_courrier", ""),
        date_courrier=s.get("date_courrier", ""),
        statut=_statut_of(s),
        statut_lot=_statut_of(s),
        observation=s.get("observation", ""),
        commercial="", tag=s.get("tag", ""),
        suivi={"statut_global": ("Désistement" if _statut_of(s) == "Désistement"
                                 else "À contacter"),
               "prochaine_action": ("" if _statut_of(s) == "Désistement"
                                    else "Appeler le client"),
               "statut_lot": _statut_of(s)},
        documents=s.get("documents", []), provenance=s.get("provenance", {}),
        sources=s.get("sources", []), source_fiche=s.get("source_fiche", ""),
        source_folder=s.get("source_folder", ""), source_mtime=s.get("source_mtime", 0),
        naissance=s.get("naissance", ""), cin_validite=s.get("cin_validite", ""),
        pere=s.get("pere", ""), mere=s.get("mere", ""),
        recu_num=s.get("recu_num", ""), recu_montant=s.get("recu_montant", ""),
        recu_date=s.get("recu_date", ""),
        _added_from_scan=True,
    )
    return rec


def build_base(suivi_path, old_clients=None):
    """Construit la base client complète prête pour l'app depuis le Suivi.

    Renvoie (clients, stats). Chaque client a ses champs racine + `suivi` +
    enrichissement éventuel. Les ids/hash sont posés ici.
    """
    import extract
    rows = [r for r in read_suivi(suivi_path) if r.get("nom_prenom") or r.get("numero")]
    # ne garder que les lots VENDUS (avec un nom) pour la base active ; les lots
    # libres/à renseigner restent visibles mais sans enrichissement.
    base = os.path.basename(suivi_path)
    clients = []
    for r in rows:
        nom = (r.get("nom_prenom") or "").strip()
        noms = [n.strip() for n in re.split(r"\s*/\s*|\s*&\s*|\s*;\s*", nom) if n.strip()]
        rec = dict(
            ilot=r.get("ilot", ""), immeuble=r.get("immeuble", ""),
            numero=r.get("numero", ""), type_bien=r.get("type_bien", "Appartement"),
            etage=r.get("etage", ""), nom_prenom=" / ".join(noms) if noms else nom,
            noms=noms or ([nom] if nom else []),
            cin=[], cin_str="", quote_part=["100"] if noms else [],
            adresse1="", adresse2="", ville="", pays="",
            tel=r.get("tel", ""), superficie=r.get("superficie", ""),
            prix_total=r.get("prix_total", ""), avance=r.get("avance", ""),
            reste=r.get("reste", ""), date_vente=r.get("date_vente", ""),
            date_courrier="",
            statut=r.get("statut_lot", "") or ("Vendu" if nom else "Libre"),
            observation=r.get("observation", ""),
            commercial=r.get("commercial", ""),      # commercial du Suivi (repli)
            tag="",                                   # agent/dossier (posé au scan)
            suivi=r.get("suivi", {}),                 # bloc de suivi des 7 étapes
            documents=[], provenance={}, sources=[base], source_fiche="",
            source_mtime=0,
            xl_sheet=r.get("xl_sheet", ""), xl_row=r.get("xl_row", 0),
            xl_header_row=r.get("xl_header_row", 0),
            xl_colkey=r.get("xl_colkey", {}),
        )
        clients.append(rec)

    stats = {"total": len(clients)}
    if old_clients:
        nk, nn = merge_old_enrichment(clients, old_clients)
        stats["matched_key"] = nk
        stats["matched_name"] = nn
        stats["unmatched"] = len(clients) - nk - nn

    for idx, c in enumerate(clients):
        c["id"] = idx + 1
        c["cin_str"] = c.get("cin_str") or " / ".join(c.get("cin", []))
        try:
            c["hash"] = extract._content_hash(c)
        except Exception:
            c["hash"] = ""
    return clients, stats


# --------------------------------------------------------- export (gabarit)
def _backup(path):
    import shutil, glob
    d = os.path.dirname(path)
    base = os.path.splitext(os.path.basename(path))[0]
    # horodatage via le mtime (pas de Date.now ici) : suffixe incrémental.
    existing = glob.glob(os.path.join(d, base + "_backup_*.xlsx"))
    bk = os.path.join(d, f"{base}_backup_{len(existing)+1}.xlsx")
    try:
        shutil.copy2(path, bk)
    except Exception:
        bk = ""
    return bk


def write_suivi_back(path, clients):
    """Réécrit les données (racine + bloc `suivi`) dans LE classeur Suivi, en
    conservant sa mise en forme. On n'écrit QUE les lots qui ont une origine
    Excel (xl_sheet/xl_row/xl_colkey). Les lots ajoutés par scan (autres îlots)
    n'ont pas de ligne dans ce classeur -> ignorés. Sauvegarde préalable.

    Renvoie {ok, backup, written, skipped}.
    """
    if not openpyxl:
        raise RuntimeError("openpyxl manquant.")
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    backup = _backup(path)
    wb = openpyxl.load_workbook(path)   # garde le format (pas data_only)

    def put(ws, row, colkey, c):
        suivi_d = c.get("suivi") or {}
        for col_str, key in colkey.items():
            try:
                col = int(col_str)
            except (ValueError, TypeError):
                continue
            val = c.get(key, "") if (key in ROOT_FIELDS and key in c) else suivi_d.get(key, "")
            if isinstance(val, (list, tuple)):
                val = " / ".join(str(x) for x in val)
            if val not in (None, ""):
                ws.cell(row=row, column=col, value=val)

    # disposition de référence (colonne->clé) d'une feuille modèle, pour écrire
    # les lots des autres îlots qui n'ont pas de ligne dans le gabarit.
    ref_colkey = None
    ref_sheet = None
    ref_hr = None
    for ws in wb.worksheets:
        hr = _find_header_row(ws)
        if hr:
            ck = {}
            for col in range(1, ws.max_column + 1):
                k = HEADER_MAP.get(_norm(ws.cell(hr, col).value))
                if k and k not in ck.values():
                    ck[col] = k
            if ck:
                ref_colkey = {str(c): k for c, k in ck.items()}
                ref_sheet = ws.title
                ref_hr = hr
                break

    written = extra = 0
    no_origin = []
    for c in clients:
        sheet = c.get("xl_sheet")
        row = c.get("xl_row")
        colkey = c.get("xl_colkey") or {}
        if not sheet or not row or sheet not in wb.sheetnames or not colkey:
            if (c.get("nom_prenom") or "").strip():
                no_origin.append(c)
            continue
        put(wb[sheet], row, colkey, c)
        written += 1

    # Lots SANS origine Excel (îlots 2-5) -> feuille « Autres lots » avec le même
    # en-tête que le modèle (pour tout avoir dans le même classeur).
    if no_origin and ref_colkey:
        if "Autres lots" in wb.sheetnames:
            del wb["Autres lots"]
        ws2 = wb.create_sheet("Autres lots")
        src = wb[ref_sheet]
        for col_str, key in ref_colkey.items():
            col = int(col_str)
            ws2.cell(row=1, column=col, value=src.cell(ref_hr, col).value)
        r = 2
        for c in no_origin:
            put(ws2, r, ref_colkey, c)
            extra += 1
            r += 1

    wb.save(path)
    return {"ok": True, "backup": backup, "written": written,
            "extra_sheet": extra, "skipped": 0}


def export_filtered(template_path, clients, out_path, keep_only=True):
    """Exporte une SÉLECTION de lots dans une COPIE du gabarit Suivi.

    - part d'une copie du classeur modèle (format conservé) ;
    - écrit les données des `clients` fournis (ceux qui ont une origine Excel) ;
    - si keep_only : efface les lignes de lot NON sélectionnées (dans la plage
      de données de chaque feuille) pour ne garder que la sélection.
    Renvoie {ok, written, path}.
    """
    import shutil
    if not openpyxl:
        raise RuntimeError("openpyxl manquant.")
    shutil.copy2(template_path, out_path)
    wb = openpyxl.load_workbook(out_path)

    # index des clients sélectionnés par (feuille, ligne)
    sel = {}
    for c in clients:
        sh, row = c.get("xl_sheet"), c.get("xl_row")
        if sh and row:
            sel[(sh, int(row))] = c

    def put(ws, row, colkey, c):
        suivi_d = c.get("suivi") or {}
        for col_str, key in colkey.items():
            try:
                col = int(col_str)
            except (ValueError, TypeError):
                continue
            val = c.get(key, "") if (key in ROOT_FIELDS and key in c) else suivi_d.get(key, "")
            if isinstance(val, (list, tuple)):
                val = " / ".join(str(x) for x in val)
            if val not in (None, ""):
                ws.cell(row=row, column=col, value=val)

    written = 0
    ref_header_colkey = None   # layout colonne->clé d'une feuille modèle
    ref_header_row = None
    for ws in wb.worksheets:
        hr = _find_header_row(ws)
        if not hr:
            continue
        colkey_sheet = {}
        for col in range(1, ws.max_column + 1):
            k = HEADER_MAP.get(_norm(ws.cell(hr, col).value))
            if k and k not in colkey_sheet.values():
                colkey_sheet[col] = k
        if ref_header_colkey is None and colkey_sheet:
            ref_header_colkey = {str(c): k for c, k in colkey_sheet.items()}
            ref_header_row = hr
            ref_sheet_title = ws.title
        numcol = next((c for c, k in colkey_sheet.items() if k == "numero"), None)
        for row in range(hr + 1, ws.max_row + 1):
            is_lot = numcol and ws.cell(row, numcol).value not in (None, "")
            c = sel.get((ws.title, row))
            if c:
                put(ws, row, c.get("xl_colkey") or {}, c)
                written += 1
            elif keep_only and is_lot:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col, value=None)

    # Clients SANS origine Excel (îlots 2-5 absents du gabarit) -> nouvelle feuille
    # « Autres lots », avec le même en-tête que le modèle.
    extra = [c for c in clients if not (c.get("xl_sheet") and c.get("xl_row"))]
    if extra and ref_header_colkey:
        # Éviter le doublon « Autres lots1 » : on supprime l'onglet s'il existe déjà.
        if "Autres lots" in wb.sheetnames:
            del wb["Autres lots"]
        ws2 = wb.create_sheet("Autres lots")
        # recopier la ligne d'en-tête du modèle
        src = wb[ref_sheet_title]
        for col_str, key in ref_header_colkey.items():
            col = int(col_str)
            ws2.cell(row=1, column=col, value=src.cell(ref_header_row, col).value)
        r = 2
        for c in extra:
            put(ws2, r, ref_header_colkey, c)
            written += 1
            r += 1

    wb.save(out_path)
    return {"ok": True, "written": written, "path": out_path}
