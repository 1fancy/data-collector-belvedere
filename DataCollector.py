#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Point d'entrée de l'application empaquetée (Data Collector).

C'est le fichier que PyInstaller transforme en .exe. Il lance simplement le
serveur local et ouvre le navigateur. Toute la logique est dans app.py.
"""

import os
import sys

# Dossier du script / de l'exe
if getattr(sys, "frozen", False):
    sys.path.insert(0, getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)))
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if __name__ == "__main__":
    try:
        import app
        app.main()
    except Exception:
        import traceback
        err = traceback.format_exc()
        # Écrire l'erreur à côté de l'exe pour diagnostic.
        try:
            base = os.path.dirname(os.path.abspath(sys.executable)) \
                if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
            with open(os.path.join(base, "erreur_demarrage.txt"), "w", encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        sys.stderr.write(err)
        sys.stderr.flush()
        try:
            input("\nUne erreur est survenue. Appuyez sur Entrée pour fermer.")
        except Exception:
            pass
        sys.exit(1)
