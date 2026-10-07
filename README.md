# Data Collector — Romana Immobilier

Application de consolidation des données clients et de publipostage pour la
Résidence Belvédère.

- **Hors-ligne** : fonctionne sans internet (sauf la toute première installation).
- **Portable** : s'installe dans un dossier, fonctionne sur n'importe quel PC.
- **Privé** : toutes les données restent sur l'ordinateur. Ce dépôt ne contient
  **que le code** du logiciel, aucune donnée client.
- **OCR intégré** : lit les documents scannés (CIN, reçus) sans rien installer.

---

## Installation (Windows)

1. Téléchargez **`INSTALLER.bat`** (bouton vert **Code → Download ZIP**, ou
   téléchargez le fichier seul).
2. Double-cliquez dessus.
3. Entrez le **code d'installation** quand il est demandé.
4. Patientez : le logiciel se télécharge et se construit tout seul. À la fin,
   un raccourci **« Data Collector »** apparaît sur le Bureau.

> Il faut **Python 3** sur le PC. S'il manque, l'installateur ouvre la page de
> téléchargement et vous indique quoi faire (cochez « Add Python to PATH »).

## Mise à jour

Double-cliquez sur **`METTRE_A_JOUR.bat`** (placé à côté du dossier
« Data Collector »). Il télécharge la dernière version et reconstruit
l'application **sans toucher à vos données**.

## macOS

Double-cliquez sur **`DEMARRER_Mac.command`** (le logiciel utilise l'OCR intégré
de macOS, Apple Vision — rien à installer).

---

## Pour les développeurs

- `app.py` — serveur web local (http.server, sans dépendance obligatoire)
- `extract.py` — extraction multi-format (doc/docx/txt/pdf + OCR)
- `ocr.py` — OCR : Apple Vision (Mac) → Tesseract → RapidOCR (intégré)
- `letters.py` — génération des lettres (modèle .docx + PDF/HTML)
- `store.py` — stockage SQLite (projets, clients, réglages)
- `templates/index.html` — interface
- `DataCollector.spec` — configuration PyInstaller (build de l'exe)

Construire l'exe : `pyinstaller DataCollector.spec` (résultat dans `dist/`).
