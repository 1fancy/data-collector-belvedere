# -*- coding: utf-8 -*-
"""Stockage SQLite pour la console BELVEDERE.

SQLite est inclus dans Python (aucune installation). On y garde :
  - clients   : une ligne par bien, avec toutes les données (JSON pour les
                champs composés : noms, cin, quote_part, documents, provenance…)
  - owners    : une ligne par acquéreur (nom, CIN, quote-part) rattachée au bien,
                pour pouvoir traiter chaque co-acquéreur séparément
  - settings  : préférences (dernier dossier, taille de page…)

Le fichier reste `belvedere.db` à côté du programme. L'export Excel et le cache
des lettres continuent de fonctionner comme avant.
"""

import os
import json
import shutil
import sqlite3

import paths

HERE = os.path.dirname(os.path.abspath(__file__))

# --- Config applicative (hors projet) : projet courant, liste récente, code… ---
APP_CONFIG = paths.data("app_config.json")
PROJECTS_DIR = paths.data("Projets")          # projets créés par défaut
DEFAULT_DB = os.path.join(PROJECTS_DIR, "Projet par defaut.db")


def _load_app_config():
    try:
        return json.load(open(APP_CONFIG, encoding="utf-8"))
    except Exception:
        return {}


def _save_app_config(cfg):
    try:
        os.makedirs(os.path.dirname(APP_CONFIG), exist_ok=True)
        json.dump(cfg, open(APP_CONFIG, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
    except OSError:
        pass


def current_db():
    """Chemin de la base du projet courant. Crée le projet par défaut au besoin."""
    cfg = _load_app_config()
    p = cfg.get("current_project")
    if p and os.path.exists(os.path.dirname(p) or "."):
        return p
    os.makedirs(PROJECTS_DIR, exist_ok=True)
    set_current_project(DEFAULT_DB)
    return DEFAULT_DB


def set_current_project(path):
    cfg = _load_app_config()
    cfg["current_project"] = path
    recent = [x for x in cfg.get("recent", []) if x != path]
    recent.insert(0, path)
    cfg["recent"] = recent[:12]
    _save_app_config(cfg)


def recent_projects():
    cfg = _load_app_config()
    out = []
    for p in cfg.get("recent", []):
        if os.path.exists(p):
            out.append({"path": p, "name": os.path.splitext(os.path.basename(p))[0]})
    return out


def current_project_info():
    db = current_db()
    return {"path": db, "name": os.path.splitext(os.path.basename(db))[0],
            "exists": os.path.exists(db)}


# Champs composés stockés en JSON.
_JSON_FIELDS = ("noms", "cin", "quote_part", "documents", "provenance", "sources")


def _connect():
    db = current_db()
    d = os.path.dirname(db)
    if d and not os.path.exists(d):
        os.makedirs(d, exist_ok=True)
    con = sqlite3.connect(db)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    return con


def init():
    con = _connect()
    con.executescript("""
    CREATE TABLE IF NOT EXISTS meta (
        key TEXT PRIMARY KEY, value TEXT
    );
    CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY,
        bkey TEXT UNIQUE,
        data TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS owners (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        idx INTEGER NOT NULL,
        nom TEXT, cin TEXT, quote_part TEXT,
        FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY, value TEXT
    );
    """)
    con.commit()
    con.close()


def _bkey(c):
    return f"{c.get('ilot')}|{c.get('immeuble')}|{c.get('numero')}|{(c.get('type_bien') or '')[:4].lower()}"


def save_clients(clients, root, generated):
    """Remplace tout le contenu par la liste `clients` (après un scan)."""
    init()
    con = _connect()
    con.execute("DELETE FROM clients")
    con.execute("DELETE FROM owners")
    for c in clients:
        con.execute("INSERT OR REPLACE INTO clients(id, bkey, data) VALUES (?,?,?)",
                    (c["id"], _bkey(c), json.dumps(c, ensure_ascii=False)))
        noms = c.get("noms") or ([c["nom_prenom"]] if c.get("nom_prenom") else [])
        cins = c.get("cin") or []
        qps = c.get("quote_part") or []
        for i, nom in enumerate(noms):
            con.execute(
                "INSERT INTO owners(client_id, idx, nom, cin, quote_part) VALUES (?,?,?,?,?)",
                (c["id"], i, nom, cins[i] if i < len(cins) else "",
                 qps[i] if i < len(qps) else (qps[0] if len(qps) == 1 and len(noms) == 1 else "")))
    set_meta(con, "root", root)
    set_meta(con, "generated", generated)
    con.commit()
    con.close()


def set_meta(con, key, value):
    con.execute("INSERT OR REPLACE INTO meta(key, value) VALUES (?,?)", (key, str(value)))


def get_meta(key, default=""):
    init()
    con = _connect()
    row = con.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    con.close()
    return row["value"] if row else default


def load_clients():
    init()
    con = _connect()
    rows = con.execute("SELECT data FROM clients ORDER BY id").fetchall()
    con.close()
    return [json.loads(r["data"]) for r in rows]


def get_client(cid):
    init()
    con = _connect()
    row = con.execute("SELECT data FROM clients WHERE id=?", (cid,)).fetchone()
    con.close()
    return json.loads(row["data"]) if row else None


def update_client(cid, patch):
    """Modifie certains champs d'un client (édition manuelle)."""
    init()
    con = _connect()
    row = con.execute("SELECT data FROM clients WHERE id=?", (cid,)).fetchone()
    if not row:
        con.close()
        return None
    c = json.loads(row["data"])
    c.update(patch)
    con.execute("UPDATE clients SET data=? WHERE id=?",
                (json.dumps(c, ensure_ascii=False), cid))
    con.commit()
    con.close()
    return c


def owners_of(cid):
    init()
    con = _connect()
    rows = con.execute(
        "SELECT idx, nom, cin, quote_part FROM owners WHERE client_id=? ORDER BY idx",
        (cid,)).fetchall()
    con.close()
    return [dict(r) for r in rows]


# ---- préférences (côté serveur, en plus de localStorage côté navigateur) ----

def get_setting(key, default=None):
    init()
    con = _connect()
    row = con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    con.close()
    if not row:
        return default
    try:
        return json.loads(row["value"])
    except (ValueError, TypeError):
        return row["value"]


def set_setting(key, value):
    init()
    con = _connect()
    con.execute("INSERT OR REPLACE INTO settings(key, value) VALUES (?,?)",
                (key, json.dumps(value, ensure_ascii=False)))
    con.commit()
    con.close()


def has_data():
    init()
    con = _connect()
    n = con.execute("SELECT COUNT(*) n FROM clients").fetchone()["n"]
    con.close()
    return n > 0


def get_folders():
    """Liste des dossiers scannés de ce projet."""
    return get_setting("scan_folders", []) or []


def set_folders(folders):
    uniq = []
    for f in folders:
        import os as _os
        ap = _os.path.abspath(_os.path.expanduser(f))
        if ap and ap not in uniq:
            uniq.append(ap)
    set_setting("scan_folders", uniq)
    return uniq


def add_folders(folders):
    return set_folders(get_folders() + list(folders))


# --------------------------------------------------------------------------- #
#  Gestion des projets (nouveau / ouvrir / enregistrer sous)                  #
# --------------------------------------------------------------------------- #

def _safe_filename(name):
    import re
    name = re.sub(r"[^\w\s().-]", "", name or "", flags=re.U).strip() or "Projet"
    return name[:80]


def new_project(name="Nouveau projet"):
    """Crée un nouveau projet vide et bascule dessus."""
    os.makedirs(PROJECTS_DIR, exist_ok=True)
    base = _safe_filename(name)
    path = os.path.join(PROJECTS_DIR, base + ".db")
    i = 2
    while os.path.exists(path):
        path = os.path.join(PROJECTS_DIR, f"{base} ({i}).db")
        i += 1
    set_current_project(path)
    init()  # crée les tables dans la nouvelle base
    return current_project_info()


def open_project(path):
    if not os.path.exists(path):
        return None
    set_current_project(path)
    init()
    return current_project_info()


def save_as(path):
    """Enregistre le projet courant sous un nouveau fichier et bascule dessus."""
    src = current_db()
    # Forcer l'écriture du WAL dans le fichier principal avant la copie.
    try:
        con = sqlite3.connect(src)
        con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        con.close()
    except Exception:
        pass
    if not path.lower().endswith(".db"):
        path += ".db"
    d = os.path.dirname(path)
    if d:
        os.makedirs(d, exist_ok=True)
    if os.path.abspath(path) != os.path.abspath(src) and os.path.exists(src):
        shutil.copy2(src, path)
    set_current_project(path)
    return current_project_info()


# --------------------------------------------------------------------------- #
#  Config applicative : code d'accès, réglages globaux                         #
# --------------------------------------------------------------------------- #

def app_get(key, default=None):
    return _load_app_config().get(key, default)


def app_set(key, value):
    cfg = _load_app_config()
    cfg[key] = value
    _save_app_config(cfg)


def _hash_code(code, salt):
    import hashlib
    return hashlib.sha256((salt + ":" + code).encode("utf-8")).hexdigest()


def has_access_code():
    cfg = _load_app_config()
    return bool(cfg.get("access_hash"))


def set_access_code(code):
    import secrets
    salt = secrets.token_hex(8)
    cfg = _load_app_config()
    cfg["access_salt"] = salt
    cfg["access_hash"] = _hash_code(code, salt)
    _save_app_config(cfg)


def check_access_code(code):
    cfg = _load_app_config()
    if not cfg.get("access_hash"):
        return True
    return _hash_code(code, cfg.get("access_salt", "")) == cfg["access_hash"]


def clear_access_code():
    cfg = _load_app_config()
    cfg.pop("access_hash", None)
    cfg.pop("access_salt", None)
    _save_app_config(cfg)


def rename_current(new_name):
    """Renomme le projet courant (déplace le fichier)."""
    src = current_db()
    new_path = os.path.join(os.path.dirname(src), _safe_filename(new_name) + ".db")
    if os.path.abspath(new_path) != os.path.abspath(src):
        try:
            con = sqlite3.connect(src)
            con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            con.close()
        except Exception:
            pass
        if os.path.exists(src):
            shutil.move(src, new_path)
        for ext in ("-wal", "-shm"):
            if os.path.exists(src + ext):
                try:
                    os.remove(src + ext)
                except OSError:
                    pass
        set_current_project(new_path)
    return current_project_info()
