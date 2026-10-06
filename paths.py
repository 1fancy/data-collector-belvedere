# -*- coding: utf-8 -*-
"""Résolution des chemins — fonctionne en script Python ET en .exe (PyInstaller).

Deux familles de chemins :
  - resource(...) : fichiers FOURNIS avec le programme, en lecture seule
      (templates/, static/, vendor/, modèle de lettre). Dans un .exe, ils sont
      extraits par PyInstaller dans un dossier temporaire (sys._MEIPASS).
  - data(...)     : fichiers ÉCRITS par le programme (base SQLite, export Excel,
      cache OCR, lettres générées). Ils doivent persister : on les met à côté de
      l'exécutable (ou du script), jamais dans le dossier temporaire.
"""

import os
import sys

_FROZEN = getattr(sys, "frozen", False)


def _resource_base():
    if _FROZEN:
        # Dossier d'extraction PyInstaller (--onefile) ou dossier de l'exe (--onedir)
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _data_base():
    if _FROZEN:
        # À côté de l'exécutable, pour que les données restent après fermeture.
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


RESOURCE_DIR = _resource_base()
DATA_DIR = _data_base()


def resource(*parts):
    return os.path.join(RESOURCE_DIR, *parts)


def data(*parts):
    p = os.path.join(DATA_DIR, *parts)
    d = os.path.dirname(p)
    if d and not os.path.exists(d):
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            pass
    return p
