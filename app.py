#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Interface web BELVEDERE — serveur local sans dépendance externe.

Fonctionne avec le seul Python 3 standard (module http.server). Les bibliothèques
optionnelles (python-docx, openpyxl, lecteurs PDF, Tesseract) sont utilisées si
présentes, mais ne sont pas requises pour démarrer.
"""

import os
import re
import sys
import json
import threading
import webbrowser
import urllib.parse
from datetime import date
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import paths
import extract
import letters
import store

DATA_FILE = paths.data("data.json")
XLSX_FILE = paths.data("BELVEDERE_clients.xlsx")
STATIC = paths.resource("static")
TEMPLATES = paths.resource("templates")

STATE = {
    "root": os.path.dirname(HERE),
    "status": "idle", "pct": 0, "msg": "Prêt.",
    "clients": [], "error": "", "generated": "",
    "phase": "",
}
LOCK = threading.Lock()

# État de la génération groupée de lettres (tâche de fond + barre de progression).
BATCH = {"status": "idle", "done": 0, "total": 0, "msg": "",
         "files": [], "error": "", "out_dir": "", "combined": ""}

# Verrouillage d'une ligne : elle se verrouille LOCK_DELAY secondes après la
# dernière étape cochée ; le code la déverrouille temporairement.
LOCK_DELAY = 120            # 2 minutes
UNLOCK_WINDOW = 300         # déverrouillage valable 5 min après le bon code
ACCESS_UNLOCK_CODE = "1111"
_UNLOCKED = {}              # bkey -> timestamp epoch de déverrouillage


_FOLDER_INDEX = {"base": None, "map": {}}
def _find_source_folder(base, folder):
    """Chemin d'un dossier client par son nom (index mis en cache)."""
    if not base or not os.path.isdir(base):
        return ""
    if _FOLDER_INDEX["base"] != base:
        m = {}
        for dp, dns, _ in os.walk(base):
            dns[:] = [d for d in dns if not d.startswith(".") and d != "Sauvegardes"]
            m.setdefault(os.path.basename(dp), dp)
        _FOLDER_INDEX["base"] = base
        _FOLDER_INDEX["map"] = m
    return _FOLDER_INDEX["map"].get(folder, "")


def _file_role(name):
    """Déduit le rôle d'une pièce d'après son nom + l'info qu'elle apporte."""
    import unicodedata
    n = "".join(ch for ch in unicodedata.normalize("NFKD", name.lower())
                if not unicodedata.combining(ch))
    if "notaire" in n or ("fiche" in n and "compta" not in n and "service" not in n):
        return ("Fiche notaire", "nom, CIN, adresse, bien, quote-part")
    if "compta" in n or "service" in n:
        return ("Fiche comptabilité", "nom, versements, date de vente")
    if "engagement" in n or "contrat" in n or "compromis" in n:
        return ("Engagement préliminaire", "adresse, identité")
    if re.search(r"\bcin\b|carte|c\.i\.n", n):
        return ("Carte d'identité (CIN)", "CIN, adresse")
    if "cheque" in n or "chèque" in n or "chéque" in n:
        return ("Chèque", "paiement")
    if "recu" in n or "reçu" in n or "reçue" in n or "récu" in n:
        return ("Reçu", "paiement, date")
    if "virement" in n or "ordre de virement" in n:
        return ("Virement", "paiement")
    if "decharge" in n or "décharge" in n:
        return ("Décharge", "")
    if "scan" in n or re.search(r"img[_\d]", n) or "whatsapp" in n:
        return ("Scan", "pièce numérisée")
    return ("Document", "")


def _row_locked(bkey, st=None):
    """La ligne est-elle verrouillée maintenant ? (st: étapes déjà chargées)"""
    import time
    if st is None:
        st = store.get_stages(bkey)
    lock_at = st.get("_lock_at")
    if not lock_at:
        return False
    # déverrouillage récent ?
    u = _UNLOCKED.get(bkey)
    if u and (time.time() - u) < UNLOCK_WINDOW:
        return False
    import datetime
    try:
        t = datetime.datetime.strptime(lock_at, "%Y-%m-%dT%H:%M:%S")
    except Exception:
        return False
    return (datetime.datetime.now() - t).total_seconds() >= LOCK_DELAY


def load_cache():
    # Source de vérité : SQLite. (data.json est importé si la base est vide,
    # pour récupérer un ancien scan.)
    try:
        if store.has_data():
            STATE["clients"] = store.load_clients()
            STATE["root"] = store.get_meta("root", STATE["root"])
            STATE["generated"] = store.get_meta("generated", "")
        elif os.path.exists(DATA_FILE):
            data = json.load(open(DATA_FILE, encoding="utf-8"))
            STATE["clients"] = data.get("clients", [])
            STATE["root"] = data.get("root", STATE["root"])
            STATE["generated"] = data.get("generated", "")
            if STATE["clients"]:
                store.save_clients(STATE["clients"], STATE["root"], STATE["generated"])
        if STATE["clients"]:
            STATE.update(status="done", pct=100,
                         msg=f"{len(STATE['clients'])} clients (dernier scan).")
    except Exception:
        import traceback; traceback.print_exc()


def save_data():
    # SQLite (vérité) + data.json (compat/sauvegarde lisible).
    store.save_clients(STATE["clients"], STATE["root"], STATE["generated"])
    try:
        data = {"root": STATE["root"], "generated": STATE["generated"],
                "clients": STATE["clients"]}
        json.dump(data, open(DATA_FILE, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
    except Exception:
        pass


def run_scan(folders, mode, auto_letters, deep=True):
    def prog(msg, pct):
        with LOCK:
            STATE["msg"] = msg
            if pct is not None:
                STATE["pct"] = int(pct * (0.7 if auto_letters else 1.0))
    locked = bool(store.get_setting("source_lock", False))
    try:
        with LOCK:
            STATE.update(status="running", pct=0, error="",
                         msg="Démarrage du scan...", phase="scan")
            existing = STATE["clients"] if (mode == "fill" or locked) else None
        # Si verrouillé sur la source (Excel importé) : on complète seulement les
        # clients existants, on n'en ajoute pas de nouveaux.
        is_suivi = bool(store.get_setting("is_suivi", False))
        scan_mode = "fill" if locked else mode
        scanned = extract.scan(folders, mode=scan_mode, existing=None,
                               deep=deep, progress=prog)
        if is_suivi and existing:
            # Base Suivi : on garde les lots existants, on les ENRICHIT
            # (CIN/docs/tag) et on AJOUTE les dossiers des autres îlots
            # (2-5) absents du Suivi (îlot 1 seulement).
            import suivi as suivi_mod
            base = [dict(c) for c in existing]
            store.apply_overrides(base)
            res = suivi_mod.merge_scan(base, scanned, set_tag=True,
                                       fill_only=True, add_new=True)
            # NORMALISATION « une vente active par lot » : si plusieurs clients
            # nommés partagent le MÊME lot (îlot+imm+n°+type), on garde UNE seule
            # vente active (priorité : nom du classeur Excel, puis le plus complet)
            # et on bascule les AUTRES en DÉSISTEMENT (anciens acquéreurs du lot).
            import unicodedata as _ud

            def _nn(s):
                s = _ud.normalize("NFKD", (s or "")).encode("ascii", "ignore").decode()
                return " ".join(sorted(t for t in re.sub(r"[^a-z ]", " ", s.lower()).split() if len(t) > 2))

            def _lotkey(c):
                typ = "C" if ("ommerce" in (c.get("type_bien", "") or "").lower()
                              or "magasin" in (c.get("type_bien", "") or "").lower()) else "A"
                return (str(c.get("ilot")), str(c.get("immeuble")).upper(),
                        str(c.get("numero")), typ)

            def _richness(c):
                # score de complétude pour choisir le meilleur enregistrement
                return (1 if c.get("xl_row") else 0) * 10 + \
                       (1 if c.get("cin") else 0) * 4 + \
                       (1 if (c.get("adresse1") or "").strip() else 0) * 2 + \
                       (1 if (c.get("tel") or "").strip() else 0)

            by_lot = {}
            for c in base:
                if (c.get("nom_prenom") or "").strip():
                    by_lot.setdefault(_lotkey(c), []).append(c)
            seen_des = set()
            for lot, members in by_lot.items():
                if len(members) < 2:
                    continue
                # l'actif = celui qui a une ligne Excel, sinon le plus complet
                members.sort(key=lambda c: (0 if c.get("xl_row") else 1,
                                            -_richness(c)))
                active = members[0]
                active["is_desistement"] = False
                for c in members[1:]:
                    nm = _nn(c.get("nom_prenom"))
                    # même personne que l'actif -> c'est un doublon, pas un désist
                    if nm == _nn(active.get("nom_prenom")):
                        c["_drop_dup"] = True
                        continue
                    key = lot + (nm,)
                    if key in seen_des:
                        c["_drop_dup"] = True
                        continue
                    seen_des.add(key)
                    c["is_desistement"] = True
                    c["statut"] = "Désistement"
                    c["statut_lot"] = "Désistement"
                    sv = c.get("suivi") or {}
                    sv["statut_lot"] = "Désistement"
                    sv["statut_global"] = "Désistement"
                    c["suivi"] = sv
                    c["conflict_with"] = active.get("nom_prenom", "")
            base = [c for c in base if not c.get("_drop_dup")]
            # Cohérence du flag : tout « statut = Désistement » -> is_desistement.
            for c in base:
                if "désist" in (c.get("statut", "") or "").lower() or \
                        "desist" in (c.get("statut", "") or "").lower():
                    c["is_desistement"] = True
            # Conflits classeur ↔ dossier (lot attribué à un autre nom) -> setting.
            store.set_setting("conflicts", res.get("conflicts", []))
            with LOCK:
                STATE["msg"] = (f"Enrichi : {res['matched']} lots, "
                                f"{res['new_cin']} CIN, {res['tagged']} tags, "
                                f"{res.get('added',0)} nouveaux, "
                                f"{len(res.get('conflicts',[]))} conflits.")
            clients = base
        elif locked and existing:
            keys = {store.bkey_of(c) for c in existing}
            clients = [c for c in scanned if store.bkey_of(c) in keys]
            for idx, c in enumerate(clients):
                c["id"] = idx + 1
        else:
            clients = scanned

        with LOCK:
            for idx, c in enumerate(clients):
                c["id"] = idx + 1
            STATE["clients"] = clients
            # Base Suivi : conserver le root de l'outil (les lettres s'y rangent),
            # ne pas le remplacer par le dossier scanné.
            if not is_suivi:
                STATE["root"] = folders[0] if folders else STATE["root"]
            STATE["generated"] = date.today().isoformat()
        # Mémoriser les dossiers scannés (pour re-scan), sans dupliquer.
        prevf = store.get_folders()
        newf = list(dict.fromkeys(prevf + list(folders))) if (mode == "fill" or is_suivi) else list(folders)
        store.set_folders(newf)
        save_data()
        extract.export_xlsx(clients, XLSX_FILE)

        if auto_letters:
            with LOCK:
                STATE["phase"] = "letters"
            def prog2(msg, pct):
                with LOCK:
                    STATE["msg"] = msg
                    STATE["pct"] = 70 + int((pct or 0) * 0.3)
            letters.generate_all(clients, STATE["root"], progress=prog2)
            save_data()

        with LOCK:
            STATE.update(status="done", pct=100, phase="",
                         msg=f"Terminé : {len(clients)} clients.")
    except Exception as e:
        import traceback
        traceback.print_exc()
        with LOCK:
            STATE.update(status="error", error=str(e), msg=f"Erreur : {e}")


def render_doc_preview(path):
    """Rend un .doc/.docx/.txt en HTML stylé (hors-ligne, sans upload).

    Si LibreOffice est présent, on convertit en PDF (mise en page fidèle) ;
    sinon on affiche le texte et les tableaux extraits, proprement mis en forme.
    """
    import html as _html
    ext = os.path.splitext(path)[1].lower()

    # 1) Si docx : garder la structure (paragraphes + tableaux).
    blocks = []
    if ext == ".docx":
        try:
            import docx
            d = docx.Document(path)
            for el in d.element.body.iterchildren():
                tag = el.tag.split("}")[-1]
                if tag == "p":
                    from docx.text.paragraph import Paragraph
                    txt = Paragraph(el, d).text.strip()
                    if txt:
                        blocks.append(("p", txt))
                elif tag == "tbl":
                    from docx.table import Table
                    t = Table(el, d)
                    rows = [[c.text.strip() for c in r.cells] for r in t.rows]
                    blocks.append(("table", rows))
        except Exception:
            blocks = []

    # 2) Sinon (ou si échec) : texte brut via le lecteur multi-format.
    if not blocks:
        text = extract.read_any(path)
        for line in text.splitlines():
            line = line.strip()
            if line:
                blocks.append(("p", line))

    parts = []
    for kind, content in blocks:
        if kind == "p":
            parts.append(f"<p>{_html.escape(content)}</p>")
        else:
            rowhtml = ""
            for r in content:
                cells = "".join(f"<td>{_html.escape(c)}</td>" for c in r)
                rowhtml += f"<tr>{cells}</tr>"
            parts.append(f"<table>{rowhtml}</table>")
    body = "\n".join(parts) or "<p style='color:#888'>Document vide.</p>"
    return f"""<!doctype html><html lang="fr"><head><meta charset="utf-8"><style>
 body{{font-family:"Times New Roman",Georgia,serif;font-size:13pt;line-height:1.55;color:#1a2332;
   max-width:19cm;margin:0 auto;padding:2.2cm 2cm;background:#fff}}
 p{{margin:.45em 0;text-align:justify}}
 table{{border-collapse:collapse;width:100%;margin:.8em 0}}
 td{{border:1px solid #ccd;padding:6px 9px;font-size:11.5pt;vertical-align:top}}
 tr:nth-child(even) td{{background:#f7f9fc}}
</style></head><body>{body}</body></html>"""


def shortcuts():
    """Dossiers usuels, pour un accès en un clic."""
    home = os.path.expanduser("~")
    items = [("Bureau", "~/Desktop"), ("Documents", "~/Documents"),
             ("Téléchargements", "~/Downloads"), ("Dossier personnel", "~")]
    out = []
    for label, p in items:
        full = os.path.expanduser(p)
        if os.path.isdir(full):
            out.append({"label": label, "path": full})
    return out


def safe_under_root(rel):
    """Autorise un fichier s'il est sous l'un des dossiers du projet (ou le root),
    y compris le dossier 'Letters' généré. `rel` peut être relatif ou absolu."""
    bases = set()
    for r in store.get_folders():
        bases.add(os.path.abspath(r))
    if STATE.get("root"):
        bases.add(os.path.abspath(STATE["root"]))
    # Dossiers "Letters" sous chaque base
    for b in list(bases):
        bases.add(os.path.join(b, "Letters"))
    if os.path.isabs(rel):
        full = os.path.abspath(rel)
    else:
        # rel relatif : essayer chaque base
        full = None
        for b in bases:
            cand = os.path.abspath(os.path.join(b, rel))
            if os.path.exists(cand):
                full = cand
                break
        if full is None:
            full = os.path.abspath(os.path.join(os.path.abspath(STATE["root"]), rel))
    for b in bases:
        if full == b or full.startswith(b + os.sep):
            return full
    return None


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    # --- Session « rester connecté » (cookie signé) ---------------------- #
    def _make_session_token(self):
        import hmac, hashlib
        sec = store.session_secret().encode()
        return hmac.new(sec, b"belvedere-session", hashlib.sha256).hexdigest()

    def _has_valid_session(self):
        if not store.has_access_code():
            return True   # pas de code défini : accès libre
        import hmac
        cookie = self.headers.get("Cookie", "")
        token = ""
        for part in cookie.split(";"):
            part = part.strip()
            if part.startswith("bv_session="):
                token = part[len("bv_session="):]
                break
        return bool(token) and hmac.compare_digest(token, self._make_session_token())

    def _session_cookie_header(self):
        # 30 jours, HttpOnly. (HTTPS via Herd → Secure acceptable.)
        return {"Set-Cookie":
                f"bv_session={self._make_session_token()}; Path=/; Max-Age=2592000; "
                "HttpOnly; SameSite=Lax"}

    def _file(self, path, attach=None):
        if not path or not os.path.exists(path):
            return self._send(404, {"error": "introuvable"})
        import mimetypes
        ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
        data = open(path, "rb").read()
        h = {"Content-Disposition": f'attachment; filename="{attach}"'} if attach else {}
        self._send(200, data, ctype, h)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        q = urllib.parse.parse_qs(u.query)
        p = u.path

        if p in ("/", "/index.html"):
            return self._file(os.path.join(TEMPLATES, "index.html"))
        if p.startswith("/static/"):
            return self._file(os.path.join(STATIC, os.path.basename(p)))

        if p == "/api/state":
            with LOCK:
                return self._send(200, {k: STATE[k] for k in
                        ("root", "status", "pct", "msg", "error", "generated", "phase")}
                        | {"count": len(STATE["clients"])})

        if p == "/api/batch_progress":
            with LOCK:
                return self._send(200, dict(BATCH))

        if p == "/api/conflicts":
            # Conflits détectés au scan : classeur ↔ dossier physique.
            return self._send(200, {"conflicts": store.get_setting("conflicts", [])})

        if p == "/api/download_export":
            # Télécharge un export Excel généré (dans le dossier temporaire).
            import tempfile
            name = os.path.basename(q.get("name", [""])[0])
            full = os.path.join(tempfile.gettempdir(), name)
            if (name.startswith("Suivi_export_") and name.endswith(".xlsx")
                    and os.path.isfile(full)):
                return self._file(full, name)
            return self._send(404, {"error": "introuvable"})

        if p == "/api/batch_combined":
            # Sert le PDF combiné de la dernière génération (chemin fiable,
            # enregistré côté serveur — pas fourni par la requête).
            with LOCK:
                cp = BATCH.get("combined", "")
            dl = q.get("dl", [""])[0] == "1"
            if cp and os.path.isfile(cp):
                return self._file(cp, os.path.basename(cp) if dl else None)
            return self._send(404, {"error": "introuvable"})

        if p == "/api/capabilities":
            return self._send(200, extract.capabilities())

        if p == "/api/clients":
            with LOCK:
                root = STATE["root"]
                cl = STATE["clients"]
            # Appliquer les overrides (édition + approbation) qui survivent aux scans.
            store.apply_overrides(cl)
            # refléter l'état réel du cache des lettres (sans régénérer)
            for c in cl:
                fol = os.path.join(root, "Letters", letters.folder_name(c))
                if letters.is_cached(c, fol):
                    c["letter_html"] = os.path.relpath(
                        os.path.join(fol, "lettre.html"), root)
                    dx = os.path.join(fol, "lettre.docx")
                    c["letter_docx"] = os.path.relpath(dx, root) if os.path.exists(dx) else ""
                    pf = os.path.join(fol, "lettre.pdf")
                    c["letter_pdf"] = os.path.relpath(pf, root) if os.path.exists(pf) else ""
                else:
                    c["letter_html"] = ""
                    c["letter_docx"] = ""
                    c["letter_pdf"] = ""
            # Étapes de lettre + verrouillage (par bien).
            stmap = store.all_stages()
            for c in cl:
                bk = store.bkey_of(c)
                st = stmap.get(bk, {})
                c["stages"] = st
                c["locked"] = _row_locked(bk, st)
            return self._send(200, {"root": root, "clients": cl})

        if p == "/api/history":
            cid = q.get("id", [None])[0]
            with LOCK:
                c = next((x for x in STATE["clients"] if str(x["id"]) == str(cid)), None)
            if not c:
                return self._send(404, {"ok": False})
            return self._send(200, {"ok": True, "history": store.get_history(store.bkey_of(c))})

        if p == "/api/source_files":
            # Liste les pièces (CIN, fiche, chèque...) du dossier source d'un client.
            cid = q.get("id", [None])[0]
            with LOCK:
                c = next((x for x in STATE["clients"] if str(x["id"]) == str(cid)), None)
            if not c:
                return self._send(404, {"ok": False})
            folder = c.get("dossier_source", "") or ""
            base = os.path.dirname(store.get_setting("source_excel_path", "") or "")
            found = _find_source_folder(base, folder) if (folder and base) else ""
            files = []
            if found:
                for f in sorted(os.listdir(found)):
                    fp = os.path.join(found, f)
                    if f.startswith(".") or not os.path.isfile(fp):
                        continue
                    kind = ("pdf" if f.lower().endswith(".pdf")
                            else "image" if f.lower().endswith((".jpg", ".jpeg", ".png", ".gif"))
                            else "doc" if f.lower().endswith((".doc", ".docx")) else "other")
                    files.append({"name": f, "kind": kind, "role": _file_role(f),
                                  "url": "/api/source_file?p=" + urllib.parse.quote(fp)})
            return self._send(200, {"ok": True, "folder": folder, "path": found, "files": files})

        if p == "/api/source_file":
            # Sert un fichier source (sous le dossier des ventes uniquement).
            full = os.path.abspath(q.get("p", [""])[0])
            base = os.path.abspath(os.path.dirname(store.get_setting("source_excel_path", "") or ""))
            if not base or not full.startswith(base) or not os.path.isfile(full):
                return self._send(403, {"error": "interdit"})
            return self._file(full)

        if p == "/api/setting":
            key = q.get("key", [""])[0]
            return self._send(200, {"key": key, "value": store.get_setting(key)})

        if p == "/api/project":
            return self._send(200, {"current": store.current_project_info(),
                                    "recent": store.recent_projects(),
                                    "folders": store.get_folders(),
                                    "source_excel": store.get_setting("source_excel", ""),
                                    "source_excel_path": store.get_setting("source_excel_path", ""),
                                    "source_lock": bool(store.get_setting("source_lock", False))})

        if p == "/api/folders":
            return self._send(200, {"folders": store.get_folders()})

        if p == "/api/access_status":
            return self._send(200, {"has_code": store.has_access_code(),
                                    "remembered": self._has_valid_session()})

        if p == "/api/excel_inspect":
            path = os.path.abspath(os.path.expanduser(q.get("path", [""])[0]))
            if not os.path.isfile(path):
                return self._send(400, {"ok": False, "error": "Fichier introuvable."})
            try:
                info = extract.inspect_excel(path)
            except Exception as e:
                return self._send(500, {"ok": False, "error": str(e)})
            info["ok"] = True
            return self._send(200, info)

        if p == "/api/excel_backups":
            path = store.get_setting("source_excel_path", "")
            if not path or not os.path.isfile(path):
                return self._send(200, {"ok": True, "path": path, "current": "", "backups": []})
            return self._send(200, {"ok": True, "path": path,
                                    "current": os.path.basename(path),
                                    "backups": extract.list_excel_backups(path)})

        if p == "/api/excel_download":
            # Télécharge l'Excel courant (which='current') ou une sauvegarde par nom.
            path = store.get_setting("source_excel_path", "")
            if not path:
                return self._send(404, {"error": "aucun excel"})
            which = q.get("which", ["current"])[0]
            if which == "current":
                target = path
            else:
                bkdir = os.path.join(os.path.dirname(path), "Sauvegardes")
                name = os.path.basename(which)   # pas de traversée
                target = os.path.join(bkdir, name)
            if not os.path.isfile(target) or not os.path.abspath(target).startswith(
                    os.path.abspath(os.path.dirname(path))):
                return self._send(403, {"error": "interdit"})
            return self._file(target, os.path.basename(target))

        if p == "/api/settings_all":
            return self._send(200, {
                "company": store.get_setting("company", {}),
                "naming": store.get_setting("naming", ""),
                "has_custom_logo": bool(store.get_setting("logo_data", "")),
            })

        if p == "/api/logo":
            # Logo courant : personnalisé (projet) sinon logo Romana par défaut.
            data_uri = store.get_setting("logo_data", "")
            if data_uri and data_uri.startswith("data:image"):
                import base64
                hdr, b64 = data_uri.split(",", 1)
                ctype = hdr.split(";")[0].replace("data:", "") or "image/png"
                return self._send(200, base64.b64decode(b64), ctype)
            return self._file(os.path.join(STATIC, "logo.png"))

        if p == "/api/browse":
            start = (q.get("path", [None])[0]) or os.path.expanduser("~/Desktop")
            sort = (q.get("sort", ["name"])[0])      # name | date
            order = (q.get("order", ["asc"])[0])     # asc | desc
            # Filtre d'extensions optionnel (ex. "files=xlsx,xls") pour lister
            # aussi des fichiers — utilisé par le sélecteur d'Excel.
            exts = tuple("." + e.strip().lstrip(".").lower()
                         for e in (q.get("files", [""])[0]).split(",") if e.strip())
            path = os.path.abspath(os.path.expanduser(start))
            if not os.path.isdir(path):
                path = os.path.expanduser("~")
            items = []
            files = []
            try:
                for d in os.listdir(path):
                    full = os.path.join(path, d)
                    if d.startswith("."):
                        continue
                    try:
                        mt = os.path.getmtime(full)
                    except OSError:
                        mt = 0
                    if os.path.isdir(full):
                        items.append({"name": d, "path": full, "mtime": mt})
                    elif exts and d.lower().endswith(exts):
                        try:
                            sz = os.path.getsize(full)
                        except OSError:
                            sz = 0
                        files.append({"name": d, "path": full, "mtime": mt, "size": sz})
            except PermissionError:
                items = []; files = []
            if sort == "date":
                # desc = le plus récent d'abord (mtime décroissant).
                items.sort(key=lambda x: x["mtime"], reverse=(order == "desc"))
                files.sort(key=lambda x: x["mtime"], reverse=(order == "desc"))
            else:
                items.sort(key=lambda x: x["name"].lower(), reverse=(order == "desc"))
                files.sort(key=lambda x: x["name"].lower(), reverse=(order == "desc"))
            # date lisible
            import datetime as _dt
            for it in items + files:
                it["date"] = _dt.datetime.fromtimestamp(it["mtime"]).strftime("%d/%m/%Y") if it["mtime"] else ""
            return self._send(200, {"path": path, "parent": os.path.dirname(path),
                "dirs": items, "files": files, "shortcuts": shortcuts(),
                "sort": sort, "order": order})

        if p == "/api/find_folders":
            term = (q.get("q", [""])[0]).strip().lower()
            if len(term) < 2:
                return self._send(200, {"results": [], "deep_more": False})
            # Recherche dans le dossier en cours de navigation (et ses sous-dossiers
            # proches). Si "base" n'est pas fourni, on retombe sur les emplacements usuels.
            base = q.get("base", [""])[0]
            if base:
                base = os.path.abspath(os.path.expanduser(base))
                bases = [base] if os.path.isdir(base) else []
            else:
                bases = [os.path.expanduser("~/Desktop"), os.path.expanduser("~/Documents"),
                         os.path.expanduser("~/Downloads"), os.path.expanduser("~")]
            max_depth = int(q.get("depth", ["2"])[0])  # profondeur limitée
            seen, results = set(), []
            deep_more = False
            for b in bases:
                if not os.path.isdir(b):
                    continue
                for dirpath, dirnames, _ in os.walk(b):
                    dirnames[:] = [d for d in dirnames if not d.startswith(".")]
                    depth = dirpath[len(b):].count(os.sep)
                    if depth >= max_depth:
                        # On arrête de descendre ; s'il reste des sous-dossiers,
                        # il peut y avoir des correspondances plus profondes.
                        if dirnames:
                            deep_more = True
                        dirnames[:] = []
                        continue
                    dirnames[:] = [d for d in dirnames if not d.startswith(".")]
                    for d in dirnames:
                        if term in d.lower():
                            full = os.path.join(dirpath, d)
                            if full not in seen:
                                seen.add(full)
                                results.append({"name": d, "path": full,
                                                "parent": os.path.basename(dirpath)})
                    if len(results) >= 60:
                        break
                if len(results) >= 60:
                    break
            results.sort(key=lambda r: (not r["name"].lower().startswith(term), len(r["name"])))
            return self._send(200, {"results": results[:60], "deep_more": deep_more,
                                    "base": bases[0] if bases else ""})

        if p == "/api/find_files":
            # Recherche récursive de fichiers (par ex. .xlsx) sous un dossier.
            term = (q.get("q", [""])[0]).strip().lower()
            exts = tuple("." + e.strip().lstrip(".").lower()
                         for e in (q.get("files", ["xlsx,xls"])[0]).split(",") if e.strip())
            base = q.get("base", [""])[0]
            base = os.path.abspath(os.path.expanduser(base)) if base else os.path.expanduser("~")
            if not os.path.isdir(base):
                return self._send(200, {"results": [], "deep_more": False})
            max_depth = int(q.get("depth", ["4"])[0])
            results, deep_more = [], False
            import datetime as _dt
            for dirpath, dirnames, filenames in os.walk(base):
                dirnames[:] = [d for d in dirnames if not d.startswith(".")]
                depth = dirpath[len(base):].count(os.sep)
                if depth >= max_depth:
                    if dirnames:
                        deep_more = True
                    dirnames[:] = []
                for f in filenames:
                    if f.startswith(".") or not f.lower().endswith(exts):
                        continue
                    if term and term not in f.lower():
                        continue
                    full = os.path.join(dirpath, f)
                    try:
                        mt = os.path.getmtime(full)
                    except OSError:
                        mt = 0
                    results.append({"name": f, "path": full, "mtime": mt,
                                    "parent": dirpath,
                                    "date": _dt.datetime.fromtimestamp(mt).strftime("%d/%m/%Y") if mt else ""})
                if len(results) >= 200:
                    deep_more = True
                    break
            results.sort(key=lambda r: r["mtime"], reverse=True)  # récent d'abord
            return self._send(200, {"results": results[:120], "deep_more": deep_more})

        if p.startswith("/letter/"):
            try:
                cid = int(p.rsplit("/", 1)[1])
            except ValueError:
                return self._send(404, {"error": "id"})
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
                root = STATE["root"]
            if not c:
                return self._send(404, {"error": "client"})
            logo = letters.logo_data_uri(store.get_setting("logo_data", ""))
            company = store.get_setting("company", {})
            paths = letters.ensure_letter(c, root, logo_uri=logo, company=company)
            c.update(paths)
            return self._file(safe_under_root(c["letter_html"]))

        if p == "/api/preview_letter":
            # Aperçu = le PDF réel (modèle Corail), identique à la lettre générée.
            try:
                cid = int(q.get("id", ["0"])[0])
            except ValueError:
                return self._send(404, {"error": "id"})
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"error": "client"})
            logo = store.get_setting("logo_data", "") or letters.logo_data_uri()
            company = store.get_setting("company", {})
            date = q.get("date", [""])[0]
            import tempfile
            tmp = os.path.join(tempfile.gettempdir(), f"preview_{cid}.pdf")
            ok = letters._pdf_via_fpdf(c, tmp, logo_uri=logo, company=company, date_override=date)
            if ok and os.path.exists(tmp):
                data = open(tmp, "rb").read()
                try: os.remove(tmp)
                except OSError: pass
                return self._send(200, data, "application/pdf")
            # repli : aperçu HTML si le PDF échoue
            html = letters.preview_html(c, logo, company, date)
            return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")

        if p == "/api/doc_preview":
            full = safe_under_root(q.get("p", [""])[0])
            if not full or not os.path.exists(full):
                return self._send(404, {"error": "introuvable"})
            html = render_doc_preview(full)
            return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")

        if p == "/file":
            full = safe_under_root(q.get("p", [""])[0])
            if not full:
                return self._send(403, {"error": "interdit"})
            name = os.path.basename(full) if q.get("dl", [""])[0] == "1" else None
            return self._file(full, name)

        if p == "/excel":
            from datetime import datetime
            stamp = datetime.now().strftime("%Y-%m-%d_%Hh%M")
            name = f"Belvedere_Clients_{stamp}.xlsx"
            return self._file(XLSX_FILE, name)

        return self._send(404, {"error": "introuvable"})

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)

        # Import Excel = multipart (lu séparément, pas en JSON).
        if u.path == "/api/import_excel":
            return self._import_excel()

        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b"{}"
        try:
            body = json.loads(raw or b"{}")
        except Exception:
            body = {}

        if u.path == "/api/excel_import":
            return self._excel_import(body)

        if u.path == "/api/excel_saveback":
            return self._excel_saveback()

        if u.path == "/api/scan":
            # Accepte "folders" (liste) ou "root" (compat). Dossiers inexistants ignorés.
            raw = body.get("folders")
            if not raw:
                raw = [body.get("root")] if body.get("root") else store.get_folders()
            folders = [os.path.abspath(os.path.expanduser(f)) for f in raw if f]
            folders = [f for f in folders if os.path.isdir(f)]
            if not folders:
                return self._send(400, {"ok": False, "error": "Aucun dossier valide à scanner."})
            with LOCK:
                if STATE["status"] == "running":
                    return self._send(409, {"ok": False, "error": "Scan déjà en cours."})
            mode = body.get("mode", "overwrite")
            auto = bool(body.get("auto_letters"))
            deep = bool(body.get("deep", True))
            threading.Thread(target=run_scan, args=(folders, mode, auto, deep),
                             daemon=True).start()
            return self._send(200, {"ok": True, "folders": folders, "mode": mode})

        if u.path == "/api/generate":
            cid = body.get("id")
            force = bool(body.get("force"))
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
                root = STATE["root"]
            if not c:
                return self._send(404, {"ok": False, "error": "client"})
            company = store.get_setting("company", {})
            logo = store.get_setting("logo_data", "") or letters.logo_data_uri()
            paths = letters.ensure_letter(c, root, logo_uri=logo, force=force, company=company)
            c.update(paths)
            save_data()
            return self._send(200, {"ok": True, **paths})

        if u.path == "/api/setting":
            key, value = body.get("key"), body.get("value")
            if key:
                store.set_setting(key, value)
            return self._send(200, {"ok": True})

        if u.path == "/api/update_client":
            # Édition manuelle PERSISTANTE (survit aux re-scans via overrides).
            cid = body.get("id")
            patch = body.get("patch", {})
            allowed = {"statut", "observation", "tel", "adresse1", "ville", "pays",
                       "date_courrier", "lettre_envoyee", "nom_prenom", "cin_str",
                       "quote_part", "prix_total", "avance", "reste", "type_bien",
                       "etage", "superficie", "date_vente"}
            patch = {k: v for k, v in patch.items() if k in allowed}
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"ok": False, "error": "client"})
            bkey = store.bkey_of(c)
            # Ligne verrouillée ? (étape cochée depuis > délai) -> refuser l'édition.
            if _row_locked(bkey):
                return self._send(423, {"ok": False, "locked": True,
                    "error": "Ligne verrouillée. Saisissez le code pour la déverrouiller."})
            # Capturer la valeur ACTUELLE (= extraite si pas encore modifiée)
            # comme « original » avant de la remplacer, pour pouvoir rétablir.
            originals = {}
            for k in patch:
                if k == "cin_str":
                    originals[k] = " / ".join(c.get("cin", []))
                elif k == "quote_part":
                    originals[k] = " / ".join(c.get("quote_part", []))
                else:
                    originals[k] = c.get(k, "")
                # Historique : une entrée par champ modifié (réversible).
                newv = patch[k]
                if isinstance(newv, (list, tuple)):
                    newv = " / ".join(str(x) for x in newv)
                if str(originals[k]) != str(newv):
                    store.add_history(bkey, k, originals[k], newv, kind="edit")
            o = store.set_override(store.bkey_of(c), patch=patch, originals=originals)
            with LOCK:
                c.update(patch)
                if "cin_str" in patch:
                    c["cin"] = [x.strip() for x in str(patch["cin_str"]).split("/") if x.strip()]
                c["edited"] = True
                c["edited_at"] = o["edited_at"]
                c["edited_fields"] = list(o["patch"].keys())
            return self._send(200, {"ok": True, "edited_at": o["edited_at"]})

        if u.path == "/api/update_suivi":
            # Met à jour un/des champs du bloc de suivi des livraisons (7 étapes).
            # Persistant via override sur la clé `suivi` (dict fusionné).
            cid = body.get("id")
            spatch = body.get("suivi", {}) or {}
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"ok": False, "error": "client"})
            bkey = store.bkey_of(c)
            cur = dict(c.get("suivi", {}) or {})
            for k, v in spatch.items():
                old = cur.get(k, "")
                if str(old) != str(v):
                    store.add_history(bkey, "suivi." + k, old, v, kind="suivi")
                if v == "" or v is None:
                    cur.pop(k, None)
                else:
                    cur[k] = v
            o = store.set_override(bkey, patch={"suivi": cur})
            with LOCK:
                c["suivi"] = cur
                c["edited"] = True
                c["edited_at"] = o["edited_at"]
            return self._send(200, {"ok": True, "suivi": cur, "edited_at": o["edited_at"]})

        if u.path == "/api/set_stage":
            # Cocher/décocher une étape de lettre ; horodate + historise + verrouille.
            cid = body.get("id"); stage = body.get("stage", ""); checked = bool(body.get("checked"))
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"ok": False})
            bkey = store.bkey_of(c)
            if _row_locked(bkey):
                return self._send(423, {"ok": False, "locked": True,
                    "error": "Ligne verrouillée. Code requis."})
            st = store.get_stages(bkey)
            now = store._now_iso()
            if checked:
                st[stage] = now
                store.add_history(bkey, "étape:" + stage, "", now, kind="stage")
                st["_lock_at"] = now        # (re)arme le verrou
            else:
                old = st.pop(stage, "")
                store.add_history(bkey, "étape:" + stage, old, "(décoché)", kind="stage")
            store.set_stages(bkey, st)
            return self._send(200, {"ok": True, "stages": st})

        if u.path == "/api/unlock_row":
            cid = body.get("id"); code = (body.get("code") or "").strip()
            if code != ACCESS_UNLOCK_CODE:
                return self._send(403, {"ok": False, "error": "Code incorrect."})
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"ok": False})
            import time
            _UNLOCKED[store.bkey_of(c)] = time.time()
            return self._send(200, {"ok": True})

        if u.path == "/api/revert_history":
            cid = body.get("id"); hist_id = body.get("hist_id")
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"ok": False})
            bkey = store.bkey_of(c)
            if _row_locked(bkey):
                return self._send(423, {"ok": False, "locked": True, "error": "Ligne verrouillée."})
            h = store.get_history_entry(hist_id)
            if not h or h["bkey"] != bkey:
                return self._send(404, {"ok": False, "error": "Entrée introuvable."})
            fld = h["field"]
            if fld.startswith("étape:"):      # rétablir une étape
                stage = fld.split(":", 1)[1]
                st = store.get_stages(bkey)
                if h["new"] == "(décoché)":
                    st[stage] = h["old"] or store._now_iso()
                else:
                    st.pop(stage, None)
                store.set_stages(bkey, st)
            else:                              # rétablir un champ édité
                patch = {fld: h["old"]}
                store.set_override(bkey, patch=patch)
                with LOCK:
                    c.update(patch)
                    if fld == "cin_str":
                        c["cin"] = [x.strip() for x in str(h["old"]).split("/") if x.strip()]
            store.add_history(bkey, fld, h["new"], h["old"], kind="revert")
            store.mark_history_reverted(hist_id)
            return self._send(200, {"ok": True})

        if u.path == "/api/approve_client":
            cid = body.get("id")
            approved = bool(body.get("approved", True))
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"ok": False})
            o = store.set_override(store.bkey_of(c), approved=approved)
            with LOCK:
                c["approved"] = approved
                c["approved_at"] = o["approved_at"]
            return self._send(200, {"ok": True, "approved_at": o["approved_at"]})

        if u.path == "/api/reset_client":
            # Annuler les modifications manuelles d'un bien (revenir à l'extrait).
            cid = body.get("id")
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if c:
                store.clear_override(store.bkey_of(c))
            return self._send(200, {"ok": True})

        if u.path == "/api/reset_field":
            # Rétablir la valeur extraite d'UN champ.
            cid = body.get("id"); field = body.get("field", "")
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if c and field:
                orig = store.clear_override_field(store.bkey_of(c), field)
                with LOCK:
                    if orig is not None:
                        c[field] = orig
                        if field == "cin_str":
                            c["cin"] = [x.strip() for x in str(orig).split("/") if x.strip()]
                        if field == "quote_part":
                            c["quote_part"] = [x.strip() for x in str(orig).split("/") if x.strip()]
                    # retirer de edited_fields
                    ef = store.get_overrides().get(store.bkey_of(c), {}).get("patch", {})
                    c["edited_fields"] = list(ef.keys())
                    c["edited"] = bool(ef)
            return self._send(200, {"ok": True})

        if u.path == "/api/source_lock":
            # Verrouiller/déverrouiller sur la source Excel importée.
            store.set_setting("source_lock", bool(body.get("locked")))
            return self._send(200, {"ok": True, "locked": bool(body.get("locked"))})

        if u.path == "/api/flush":
            # Vider les données de CE projet (clients + verrouillage source).
            # Les projets enregistrés ailleurs ne sont pas touchés.
            store.save_clients([], STATE["root"], "")
            store.set_setting("source_lock", False)
            store.set_setting("source_excel", "")
            store.set_folders([])
            with LOCK:
                STATE["clients"] = []
                STATE["status"] = "idle"; STATE["pct"] = 0
                STATE["msg"] = "Données vidées."
            save_data()
            return self._send(200, {"ok": True})

        if u.path == "/api/flush_history":
            # Vider tout l'historique des modifications (sent/edit log).
            n = store.clear_all_history()
            return self._send(200, {"ok": True, "cleared": n})

        if u.path == "/api/flush_stages":
            # Réinitialiser tous les statuts (étapes de lettre) + verrous.
            n = store.clear_all_stages()
            return self._send(200, {"ok": True, "cleared": n})

        if u.path == "/api/flush_letters":
            # Supprimer les lettres générées (dossiers Letters/ + PDF combiné)
            # pour qu'elles soient régénérées à neuf. Les données restent.
            import shutil
            removed = 0
            with LOCK:
                root = STATE["root"]
            ldir = os.path.join(root, "Letters")
            if os.path.isdir(ldir):
                for name in os.listdir(ldir):
                    full = os.path.join(ldir, name)
                    try:
                        if os.path.isdir(full):
                            shutil.rmtree(full); removed += 1
                        elif name.lower().endswith((".pdf", ".docx", ".html")):
                            os.remove(full); removed += 1
                    except OSError:
                        pass
            with LOCK:
                BATCH.update(status="idle", done=0, total=0, files=[],
                             combined="", msg="")
            return self._send(200, {"ok": True, "removed": removed})

        if u.path in ("/api/rederive_pays", "/api/rederive_geo"):
            # Recalcule la VILLE (depuis l'adresse) puis le PAYS (depuis la
            # ville, repli adresse) pour chaque client. Persistant via overrides.
            import geo
            n_ville = n_pays = 0
            with LOCK:
                clients = list(STATE["clients"])
            store.apply_overrides(clients)
            for c in clients:
                patch = {}
                adr = c.get("adresse1", "") or c.get("adresse", "")
                # Ville : on ne remplit QUE si vide (ne pas écraser une saisie).
                if not (c.get("ville", "") or "").strip():
                    v = geo.ville_from_addr(adr)
                    if v:
                        patch["ville"] = v
                        c["ville"] = v
                        n_ville += 1
                # Pays : recalcul systématique (corrige les erreurs type Qatar).
                new_pays = geo.pays_from(c.get("ville", ""), adr)
                if new_pays and new_pays != (c.get("pays", "") or "").strip():
                    patch["pays"] = new_pays
                    c["pays"] = new_pays
                    n_pays += 1
                if patch:
                    store.set_override(store.bkey_of(c), patch=patch)
            reload_from_store()
            return self._send(200, {"ok": True, "ville": n_ville, "pays": n_pays,
                                    "count": n_ville + n_pays})

        # ---- projets ----
        if u.path == "/api/project_new":
            info = store.new_project(body.get("name") or "Nouveau projet")
            with LOCK:
                STATE.update(clients=[], root=os.path.dirname(HERE), generated="",
                             status="idle", pct=0, msg="Nouveau projet.")
            return self._send(200, {"ok": True, "current": info})

        if u.path == "/api/project_open":
            info = store.open_project(body.get("path", ""))
            if not info:
                return self._send(404, {"ok": False, "error": "Projet introuvable"})
            reload_from_store()
            return self._send(200, {"ok": True, "current": info})

        if u.path == "/api/project_saveas":
            info = store.save_as(body.get("path", ""))
            return self._send(200, {"ok": True, "current": info})

        if u.path == "/api/project_rename":
            info = store.rename_current(body.get("name", ""))
            return self._send(200, {"ok": True, "current": info})

        # ---- code d'accès ----
        if u.path == "/api/access_set":
            code = (body.get("code") or "").strip()
            if len(code) < 4:
                return self._send(400, {"ok": False, "error": "Code trop court (4 min)."})
            store.set_access_code(code)
            # On mémorise la session tout de suite (cookie 30 jours).
            return self._send(200, {"ok": True}, headers=self._session_cookie_header())

        if u.path == "/api/access_check":
            ok = store.check_access_code((body.get("code") or "").strip())
            remember = body.get("remember", True)
            h = self._session_cookie_header() if (ok and remember) else None
            return self._send(200, {"ok": ok}, headers=h)

        if u.path == "/api/access_logout":
            # Oublier la session sur ce navigateur (réafficher le code).
            return self._send(200, {"ok": True}, headers={
                "Set-Cookie": "bv_session=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"})

        # ---- réglages (logo, société, nommage) ----
        if u.path == "/api/settings_save":
            if "company" in body:
                store.set_setting("company", body["company"])
            if "naming" in body:
                store.set_setting("naming", body["naming"])
            if "logo_data" in body:
                store.set_setting("logo_data", body["logo_data"])
            return self._send(200, {"ok": True})

        # ---- génération de lettres en lot ----
        if u.path == "/api/generate_batch":
            return self._generate_batch(body)

        if u.path == "/api/import_suivi":
            return self._import_suivi(body)

        if u.path == "/api/export_suivi":
            # Réécrit les données (racine + suivi) dans LE classeur Suivi d'origine
            # (le gabarit), en gardant sa mise en forme. Sauvegarde préalable.
            import suivi as suivi_mod
            path = store.get_setting("source_excel_path", "")
            if not path or not os.path.isfile(path):
                return self._send(400, {"ok": False,
                    "error": "Classeur Suivi d'origine introuvable."})
            with LOCK:
                cl = [dict(c) for c in STATE["clients"]]
            store.apply_overrides(cl)   # intègre les éditions de l'app
            try:
                res = suivi_mod.write_suivi_back(path, cl)
            except Exception as e:
                import traceback; traceback.print_exc()
                return self._send(500, {"ok": False, "error": str(e)})
            return self._send(200, {"ok": True, **res})

        if u.path == "/api/export_filtered":
            # Exporte une SÉLECTION (ids) dans une COPIE du gabarit Suivi, puis
            # renvoie le chemin à télécharger. N'écrase pas l'original.
            import suivi as suivi_mod, tempfile
            template = store.get_setting("source_excel_path", "")
            if not template or not os.path.isfile(template):
                return self._send(400, {"ok": False,
                    "error": "Classeur Suivi modèle introuvable."})
            ids = set(body.get("ids", []))
            with LOCK:
                cl = [dict(c) for c in STATE["clients"] if c["id"] in ids]
            store.apply_overrides(cl)
            from datetime import datetime
            stamp = datetime.now().strftime("%Y-%m-%d_%Hh%M")
            out = os.path.join(tempfile.gettempdir(),
                               f"Suivi_export_{stamp}.xlsx")
            try:
                res = suivi_mod.export_filtered(template, cl, out, keep_only=True)
            except Exception as e:
                import traceback; traceback.print_exc()
                return self._send(500, {"ok": False, "error": str(e)})
            return self._send(200, {"ok": True, "written": res["written"],
                                    "file": out, "name": os.path.basename(out)})

        return self._send(404, {"error": "introuvable"})

    def _import_suivi(self, body):
        """Importe le classeur « Suivi des livraisons » comme nouvelle base.

        - lit les 251 lots + le bloc de suivi des 7 étapes ;
        - ré-attache l'enrichissement (CIN/adresse/docs) de l'ancienne base si
          un fichier JSON d'archive est fourni (ou détecté) ;
        - démarre un nouveau projet verrouillé (un scan ne fera qu'enrichir).
        """
        import suivi as suivi_mod
        path = os.path.abspath(os.path.expanduser(body.get("path", "")))
        if not os.path.isfile(path):
            return self._send(400, {"ok": False, "error": "Fichier introuvable."})
        if not suivi_mod.looks_like_suivi(path):
            return self._send(400, {"ok": False,
                "error": "Ce fichier ne ressemble pas au Suivi des livraisons."})
        # Charger l'ancienne base pour l'enrichissement : d'abord les clients
        # On prend la source la PLUS RICHE (le plus de clients avec CIN) parmi :
        # les clients en mémoire ET tous les JSON d'archive (clean_base_*, extracted_*).
        import glob as _glob

        def _richcount(lst):
            return sum(1 for c in (lst or []) if c.get("cin"))

        candidates = []
        with LOCK:
            mem = [dict(c) for c in STATE.get("clients", [])]
        if mem:
            store.apply_overrides(mem)
            candidates.append(mem)
        for jp in _glob.glob(
                "/Users/younes/Downloads/Belvedere_files/_archives_*/clean_base_*.json") + \
                _glob.glob(
                "/Users/younes/Downloads/Belvedere_files/_archives_*/extracted_*clients.json"):
            try:
                candidates.append(json.load(open(jp))["clients"])
            except Exception:
                pass
        old = max(candidates, key=_richcount) if candidates else []
        try:
            clients, stats = suivi_mod.build_base(path, old or None)
        except Exception as e:
            import traceback; traceback.print_exc()
            return self._send(500, {"ok": False, "error": str(e)})
        if not clients:
            return self._send(400, {"ok": False, "error": "Aucun lot reconnu."})
        base = os.path.splitext(os.path.basename(path))[0]
        store.new_project("Suivi " + base)
        store.set_setting("source_excel", os.path.basename(path))
        store.set_setting("source_excel_path", path)
        store.set_setting("source_lock", True)
        store.set_setting("is_suivi", True)
        from datetime import date as _date
        with LOCK:
            STATE["clients"] = clients
            STATE["root"] = os.path.dirname(HERE)
            STATE["generated"] = _date.today().isoformat()
            STATE["status"] = "done"; STATE["pct"] = 100
            STATE["msg"] = f"{len(clients)} lots importés."
        save_data()
        try:
            extract.export_xlsx(clients, XLSX_FILE)
        except Exception:
            pass
        return self._send(200, {"ok": True, "count": len(clients),
                                "stats": stats, "source": os.path.basename(path)})

    def _import_excel(self):
        fname, data = self._read_multipart_file()
        if not data:
            return self._send(400, {"ok": False, "error": "Aucun fichier reçu."})
        import tempfile
        tmp = os.path.join(tempfile.gettempdir(), "import_" + (fname or "data.xlsx"))
        with open(tmp, "wb") as f:
            f.write(data)
        try:
            clients = extract.import_excel_clients(tmp)
        except Exception as e:
            import traceback; traceback.print_exc()
            return self._send(500, {"ok": False, "error": str(e)})
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
        if not clients:
            return self._send(400, {"ok": False,
                "error": "Aucune donnée client reconnue dans ce fichier."})
        # Nouveau projet basé sur le nom du fichier importé (jamais modifié).
        base = os.path.splitext(fname or "Import Excel")[0]
        store.new_project("Import " + base)
        # Marquer la source Excel et activer le verrouillage par défaut :
        # un scan de dossiers ne fera qu'ENRICHIR ces clients, sans en ajouter.
        store.set_setting("source_excel", base)
        store.set_setting("source_lock", True)
        from datetime import date as _date
        with LOCK:
            STATE["clients"] = clients
            STATE["root"] = os.path.dirname(HERE)
            STATE["generated"] = _date.today().isoformat()
            STATE["status"] = "done"; STATE["pct"] = 100
            STATE["msg"] = f"{len(clients)} clients importés."
        save_data()
        extract.export_xlsx(clients, XLSX_FILE)
        return self._send(200, {"ok": True, "count": len(clients)})

    def _excel_import(self, body):
        """Importe un Excel depuis un chemin local (le fichier reste en place,
        ce qui permet de réécrire dedans plus tard). Correspondance éventuelle
        fournie par l'utilisateur via l'écran de mapping."""
        path = os.path.abspath(os.path.expanduser(body.get("path", "")))
        mapping = body.get("mapping") or None
        if not os.path.isfile(path):
            return self._send(400, {"ok": False, "error": "Fichier introuvable."})
        try:
            clients = extract.import_excel_clients(path, mapping=mapping)
        except Exception as e:
            import traceback; traceback.print_exc()
            return self._send(500, {"ok": False, "error": str(e)})
        if not clients:
            return self._send(400, {"ok": False,
                "error": "Aucune donnée client reconnue. Vérifiez la colonne « Nom »."})
        base = os.path.splitext(os.path.basename(path))[0]
        store.new_project("Import " + base)
        store.set_setting("source_excel", os.path.basename(path))
        store.set_setting("source_excel_path", path)   # chemin complet pour réécriture
        store.set_setting("source_lock", True)
        from datetime import date as _date
        with LOCK:
            STATE["clients"] = clients
            STATE["root"] = os.path.dirname(HERE)
            STATE["generated"] = _date.today().isoformat()
            STATE["status"] = "done"; STATE["pct"] = 100
            STATE["msg"] = f"{len(clients)} clients importés."
        save_data()
        extract.export_xlsx(clients, XLSX_FILE)
        return self._send(200, {"ok": True, "count": len(clients),
                                "source": os.path.basename(path)})

    def _excel_saveback(self):
        """Réécrit les modifications dans le fichier Excel d'origine
        (sauvegarde horodatée au préalable). Nécessite un import par chemin."""
        path = store.get_setting("source_excel_path", "")
        if not path or not os.path.isfile(path):
            return self._send(400, {"ok": False, "error":
                "Aucun fichier Excel d'origine. Importez d'abord un Excel par son emplacement."})
        with LOCK:
            cl = [dict(c) for c in STATE["clients"]]
        store.apply_overrides(cl)   # intègre les éditions faites dans l'app
        try:
            res = extract.save_excel_back(path, cl)
        except Exception as e:
            import traceback; traceback.print_exc()
            return self._send(500, {"ok": False, "error": str(e)})
        return self._send(200, {"ok": True, **res})

    def _read_multipart_file(self):
        """Lit un unique fichier d'un envoi multipart/form-data (champ 'file')."""
        ctype = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in ctype or "boundary=" not in ctype:
            return None, None
        boundary = ctype.split("boundary=", 1)[1].strip().strip('"').encode()
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        parts = raw.split(b"--" + boundary)
        for part in parts:
            if b"filename=" not in part:
                continue
            head, _, data = part.partition(b"\r\n\r\n")
            fname = ""
            import re as _re
            m = _re.search(rb'filename="([^"]*)"', head)
            if m:
                fname = m.group(1).decode("utf-8", "replace")
            data = data.rstrip(b"\r\n")
            return fname, data
        return None, None

    def _generate_batch(self, body):
        # Lance la génération groupée EN TÂCHE DE FOND : générer 150+ lettres
        # dépasse le délai du proxy (nginx -> 502). Le front interroge ensuite
        # /api/batch_progress jusqu'à la fin.
        ids = body.get("ids", [])
        out_dir = os.path.abspath(os.path.expanduser(body.get("out_dir", "")))
        formats = body.get("formats", ["pdf"])
        date_override = body.get("date", "")
        if not os.path.isdir(out_dir):
            return self._send(400, {"ok": False, "error": f"Dossier invalide : {out_dir}"})
        with LOCK:
            if BATCH["status"] == "running":
                return self._send(409, {"ok": False, "error": "Génération déjà en cours."})
            root = STATE["root"]
            chosen = [c for c in STATE["clients"] if c["id"] in ids]
            BATCH.update(status="running", done=0, total=len(chosen),
                         msg="Préparation…", files=[], error="", out_dir=out_dir,
                         combined="")
        company = store.get_setting("company", {})
        naming = store.get_setting("naming", "")
        logo = letters.logo_data_uri(store.get_setting("logo_data", ""))
        threading.Thread(target=run_batch, daemon=True, args=(
            chosen, out_dir, formats, date_override, company, naming, logo, root
        )).start()
        return self._send(200, {"ok": True, "out_dir": out_dir, "total": len(chosen),
                                 "async": True})


def run_batch(chosen, out_dir, formats, date_override, company, naming, logo, root):
    """Génère les lettres en tâche de fond en publiant la progression."""
    def progress(label, pct):
        with LOCK:
            BATCH["done"] = min(BATCH["total"], int(round(pct / 100 * BATCH["total"])))
            BATCH["msg"] = f"{label}…"
    try:
        results = letters.generate_batch(
            chosen, out_dir, formats=formats, date_override=date_override,
            company=company, naming=naming, logo_uri=logo,
            template=letters.find_template(root), progress=progress)
        # PDF combiné : toutes les lettres dans un seul fichier (une par page).
        combined_rel = ""
        if "pdf" in formats and chosen:
            with LOCK:
                BATCH["msg"] = "Assemblage du PDF combiné…"
            combined_path = os.path.join(out_dir, "TOUTES_LES_LETTRES.pdf")
            if letters.combined_pdf(chosen, combined_path, company=company,
                                    date_override=date_override):
                combined_rel = combined_path
        with LOCK:
            BATCH.update(status="done", done=BATCH["total"], files=results,
                         combined=combined_rel,
                         msg=f"{len(results)} lettres générées.")
    except Exception as e:
        import traceback; traceback.print_exc()
        with LOCK:
            BATCH.update(status="error", error=str(e), msg="Erreur : " + str(e))


def reload_from_store():
    with LOCK:
        STATE["clients"] = store.load_clients()
        STATE["root"] = store.get_meta("root", os.path.dirname(HERE))
        STATE["generated"] = store.get_meta("generated", "")
        STATE["status"] = "done" if STATE["clients"] else "idle"
        STATE["pct"] = 100 if STATE["clients"] else 0
        STATE["msg"] = (f"{len(STATE['clients'])} clients." if STATE["clients"]
                        else "Projet vide.")


def main():
    load_cache()
    # Port fixe (stable) ; configurable via la variable PORT.
    port = int(os.environ.get("PORT", 8700))
    # Hôte : 127.0.0.1 par défaut. Mettre HOST=0.0.0.0 pour accès réseau/hébergé.
    host = os.environ.get("HOST", "127.0.0.1")
    # Si le port est occupé, on en essaie quelques autres (ne jamais planter).
    httpd = None
    for p in [port] + [port + i for i in range(1, 20)]:
        try:
            httpd = ThreadingHTTPServer((host, p), Handler)
            port = p
            break
        except OSError:
            continue
    if httpd is None:
        print("Impossible d'ouvrir un port réseau local.")
        return
    # URL affichée : domaine Herd si fourni (PUBLIC_URL), sinon on privilégie
    # https://belvedere.test (proxy Herd) si le port standard 8700 est utilisé,
    # et on ne retombe sur l'adresse locale que si le port a changé.
    public = os.environ.get("PUBLIC_URL", "").rstrip("/")
    if public:
        url = public + "/"
    elif port == 8700:
        url = "https://belvedere.test/"
    else:
        url = f"http://127.0.0.1:{port}/"
    print("\n" + "=" * 54)
    print("   DATA COLLECTOR — Romana Immobilier")
    print("   Resultats et lettres pour la Residence Belvedere")
    print("-" * 54)
    print(f"   Ouvrez cette adresse :  {url}")
    print("   (Elle devrait s'ouvrir toute seule dans votre navigateur.)")
    print("")
    print("   Laissez cette fenetre OUVERTE pendant l'utilisation.")
    print("   Pour quitter : fermez cette fenetre.")
    print("=" * 54 + "\n")
    if os.environ.get("NO_BROWSER") != "1":
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nArret.")


if __name__ == "__main__":
    main()
