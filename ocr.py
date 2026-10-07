# -*- coding: utf-8 -*-
"""Moteur OCR multi-moteurs, sans installation obligatoire.

Ordre de préférence :
  1. Vision (Apple) — intégré à macOS, aucune installation, français + arabe.
  2. Tesseract — s'il est installé (Windows/Linux, ou macOS via Homebrew).

Les PDF scannés sont d'abord rastérisés en images (pdftoppm, PyMuPDF ou sips),
puis chaque page est passée à l'OCR. Si aucun moteur n'est disponible, les
fonctions renvoient "" sans erreur : le reste du logiciel continue de marcher.
"""

import os
import glob
import json
import shutil
import hashlib
import tempfile
import subprocess

import paths

HERE = os.path.dirname(os.path.abspath(__file__))
_SWIFT = paths.resource("static", "ocr_vision.swift")
_CACHE_DIR = paths.data(".ocr_cache")

_CACHE = {}


def _cache_key(path, max_pages):
    """Clé de cache stable même si le dossier est déplacé.

    On se base sur le contenu du fichier (taille + début/fin) plutôt que sur
    son chemin absolu, pour que le cache reste valable après un déplacement.
    """
    try:
        st = os.stat(path)
        sig = f"{st.st_size}"
        with open(path, "rb") as f:
            head = f.read(65536)
            if st.st_size > 131072:
                f.seek(-65536, os.SEEK_END)
                sig += "|" + hashlib.md5(f.read(65536)).hexdigest()
        sig += "|" + hashlib.md5(head).hexdigest()
    except OSError:
        sig = "0"
    raw = f"{sig}|{max_pages}|{engine()}"
    return hashlib.md5(raw.encode("utf-8")).hexdigest()


def _cache_get(path, max_pages):
    f = os.path.join(_CACHE_DIR, _cache_key(path, max_pages) + ".txt")
    if os.path.exists(f):
        try:
            return open(f, encoding="utf-8").read()
        except OSError:
            return None
    return None


def _cache_put(path, max_pages, text):
    try:
        os.makedirs(_CACHE_DIR, exist_ok=True)
        f = os.path.join(_CACHE_DIR, _cache_key(path, max_pages) + ".txt")
        open(f, "w", encoding="utf-8").write(text)
    except OSError:
        pass


def _have(binary):
    return shutil.which(binary) is not None


def _is_macos():
    import platform
    return platform.system() == "Darwin"


def _has_rapidocr():
    try:
        import rapidocr_onnxruntime  # noqa: F401
        return True
    except Exception:
        return False


def engine():
    """Moteur OCR disponible, par ordre de préférence :
    'vision' (macOS, intégré) > 'tesseract' (si installé) > 'rapidocr' (inclus).
    """
    if "engine" in _CACHE:
        return _CACHE["engine"]
    eng = ""
    if _is_macos() and _have("swift") and os.path.exists(_SWIFT):
        eng = "vision"
    elif _have("tesseract"):
        eng = "tesseract"
    elif _has_rapidocr():
        eng = "rapidocr"
    _CACHE["engine"] = eng
    return eng


_RAPID = None


def _rapid_reader():
    global _RAPID
    if _RAPID is None:
        from rapidocr_onnxruntime import RapidOCR
        _RAPID = RapidOCR()
    return _RAPID


def available():
    return engine() != ""


def _rasterize(pdf_path, out_dir, dpi=200, max_pages=4):
    """PDF -> liste d'images PNG (une par page, limité)."""
    imgs = []
    try:
        import fitz
        doc = fitz.open(pdf_path)
        for i, pg in enumerate(doc):
            if i >= max_pages:
                break
            out = os.path.join(out_dir, f"p{i}.png")
            pg.get_pixmap(dpi=dpi).save(out)
            imgs.append(out)
        if imgs:
            return imgs
    except Exception:
        pass
    if _have("pdftoppm"):
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", "-l", str(max_pages),
                        pdf_path, os.path.join(out_dir, "p")],
                       capture_output=True, timeout=120)
        imgs = sorted(glob.glob(os.path.join(out_dir, "p*.png")))
        if imgs:
            return imgs
    if _is_macos() and _have("sips"):
        out = os.path.join(out_dir, "p0.png")
        r = subprocess.run(["sips", "-s", "format", "png", pdf_path, "--out", out],
                           capture_output=True, timeout=60)
        if r.returncode == 0 and os.path.exists(out):
            imgs = [out]
    return imgs


def ocr_image(img_path, langs=("fr-FR", "en-US", "ar-SA")):
    eng = engine()
    if eng == "vision":
        r = subprocess.run(["swift", _SWIFT, img_path, ",".join(langs)],
                           capture_output=True, timeout=90)
        if r.returncode == 0:
            return r.stdout.decode("utf-8", errors="replace")
        return ""
    if eng == "tesseract":
        r = subprocess.run(["tesseract", img_path, "-", "-l", "fra+eng+ara"],
                           capture_output=True, timeout=120)
        if r.returncode == 0:
            return r.stdout.decode("utf-8", errors="replace")
    if eng == "rapidocr":
        try:
            result, _ = _rapid_reader()(img_path)
            if result:
                # result = liste de [box, texte, score] -> on garde les textes.
                return "\n".join(line[1] for line in result)
        except Exception:
            return ""
    return ""


def ocr_pdf(pdf_path, max_pages=4):
    if not available():
        return ""
    with tempfile.TemporaryDirectory() as td:
        pages = _rasterize(pdf_path, td, max_pages=max_pages)
        return "\n".join(ocr_image(p) for p in pages)


def ocr_file(path, max_pages=4):
    """OCR d'un fichier avec cache (clé = chemin + date de modif + moteur)."""
    if not available():
        return ""
    cached = _cache_get(path, max_pages)
    if cached is not None:
        return cached
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        text = ocr_pdf(path, max_pages=max_pages)
    elif ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"):
        text = ocr_image(path)
    else:
        text = ""
    _cache_put(path, max_pages, text)
    return text


def warm_cache(paths, max_pages=1, workers=None, progress=None):
    """OCR en parallèle d'une liste de fichiers, pour remplir le cache.

    Accélère nettement la recherche approfondie : les pages sont traitées
    simultanément plutôt qu'une par une. Les fichiers déjà en cache sont ignorés.
    """
    if not available():
        return
    todo = [p for p in paths if _cache_get(p, max_pages) is None]
    if not todo:
        if progress:
            progress(len(paths), len(paths))
        return
    import concurrent.futures
    if workers is None:
        workers = min(8, (os.cpu_count() or 4))
    done = len(paths) - len(todo)
    total = len(paths)
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(ocr_file, p, max_pages): p for p in todo}
        for fut in concurrent.futures.as_completed(futs):
            done += 1
            if progress:
                progress(done, total)
