#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Interface web BELVEDERE — serveur local sans dépendance externe.

Fonctionne avec le seul Python 3 standard (module http.server). Les bibliothèques
optionnelles (python-docx, openpyxl, lecteurs PDF, Tesseract) sont utilisées si
présentes, mais ne sont pas requises pour démarrer.
"""

import os
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
        scan_mode = "fill" if locked else mode
        clients = extract.scan(folders, mode=scan_mode, existing=existing,
                               deep=deep, progress=prog)
        if locked and existing:
            keys = {store.bkey_of(c) for c in existing}
            clients = [c for c in clients if store.bkey_of(c) in keys]
            for idx, c in enumerate(clients):
                c["id"] = idx + 1

        with LOCK:
            STATE["clients"] = clients
            STATE["root"] = folders[0] if folders else STATE["root"]
            STATE["generated"] = date.today().isoformat()
        store.set_folders(folders if mode != "fill" else store.get_folders() + list(folders))
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
                else:
                    c["letter_html"] = ""
                    c["letter_docx"] = ""
            return self._send(200, {"root": root, "clients": cl})

        if p == "/api/setting":
            key = q.get("key", [""])[0]
            return self._send(200, {"key": key, "value": store.get_setting(key)})

        if p == "/api/project":
            return self._send(200, {"current": store.current_project_info(),
                                    "recent": store.recent_projects(),
                                    "folders": store.get_folders(),
                                    "source_excel": store.get_setting("source_excel", ""),
                                    "source_lock": bool(store.get_setting("source_lock", False))})

        if p == "/api/folders":
            return self._send(200, {"folders": store.get_folders()})

        if p == "/api/access_status":
            return self._send(200, {"has_code": store.has_access_code()})

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
            path = os.path.abspath(os.path.expanduser(start))
            if not os.path.isdir(path):
                path = os.path.expanduser("~")
            items = []
            try:
                for d in os.listdir(path):
                    full = os.path.join(path, d)
                    if d.startswith(".") or not os.path.isdir(full):
                        continue
                    try:
                        mt = os.path.getmtime(full)
                    except OSError:
                        mt = 0
                    items.append({"name": d, "path": full, "mtime": mt})
            except PermissionError:
                items = []
            if sort == "date":
                items.sort(key=lambda x: x["mtime"], reverse=(order == "asc"))
            else:
                items.sort(key=lambda x: x["name"].lower(), reverse=(order == "desc"))
            # date lisible
            import datetime as _dt
            for it in items:
                it["date"] = _dt.datetime.fromtimestamp(it["mtime"]).strftime("%d/%m/%Y") if it["mtime"] else ""
            return self._send(200, {"path": path, "parent": os.path.dirname(path),
                "dirs": items, "shortcuts": shortcuts(), "sort": sort, "order": order})

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
            try:
                cid = int(q.get("id", ["0"])[0])
            except ValueError:
                return self._send(404, {"error": "id"})
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"error": "client"})
            logo = letters.logo_data_uri(store.get_setting("logo_data", ""))
            company = store.get_setting("company", {})
            date = q.get("date", [""])[0]
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
            paths = letters.ensure_letter(c, root, logo_uri=letters.logo_data_uri(), force=force)
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
                       "etage", "superficie"}
            patch = {k: v for k, v in patch.items() if k in allowed}
            with LOCK:
                c = next((x for x in STATE["clients"] if x["id"] == cid), None)
            if not c:
                return self._send(404, {"ok": False, "error": "client"})
            o = store.set_override(store.bkey_of(c), patch=patch)
            with LOCK:
                c.update(patch)
                c["edited"] = True
                c["edited_at"] = o["edited_at"]
            return self._send(200, {"ok": True, "edited_at": o["edited_at"]})

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
            return self._send(200, {"ok": True})

        if u.path == "/api/access_check":
            ok = store.check_access_code((body.get("code") or "").strip())
            return self._send(200, {"ok": ok})

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

        return self._send(404, {"error": "introuvable"})

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
        ids = body.get("ids", [])
        out_dir = os.path.abspath(os.path.expanduser(body.get("out_dir", "")))
        formats = body.get("formats", ["pdf"])
        date_override = body.get("date", "")
        if not os.path.isdir(out_dir):
            return self._send(400, {"ok": False, "error": f"Dossier invalide : {out_dir}"})
        with LOCK:
            root = STATE["root"]
            chosen = [c for c in STATE["clients"] if c["id"] in ids]
        company = store.get_setting("company", {})
        naming = store.get_setting("naming", "")
        logo = letters.logo_data_uri(store.get_setting("logo_data", ""))
        try:
            results = letters.generate_batch(
                chosen, out_dir, formats=formats, date_override=date_override,
                company=company, naming=naming, logo_uri=logo,
                template=letters.find_template(root))
        except Exception as e:
            import traceback; traceback.print_exc()
            return self._send(500, {"ok": False, "error": str(e)})
        return self._send(200, {"ok": True, "out_dir": out_dir, "files": results})


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
    # URL affichée : domaine Herd si fourni, sinon l'adresse locale.
    public = os.environ.get("PUBLIC_URL", "").rstrip("/")
    url = public + "/" if public else f"http://127.0.0.1:{port}/"
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
