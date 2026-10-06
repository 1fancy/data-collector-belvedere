@echo off
chcp 65001 >nul
title Romana - Residence Belvedere
cd /d "%~dp0"

echo ============================================================
echo    RESIDENCE BELVEDERE - Console clients et lettres
echo ============================================================
echo.

rem --- 1. Trouver Python -------------------------------------------------
set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )

if not defined PY (
  echo [!] Python n'est pas installe sur cet ordinateur.
  echo.
  echo     Telechargez-le ici ^(cochez "Add Python to PATH" a l'installation^) :
  echo         https://www.python.org/downloads/
  echo.
  echo     Puis relancez ce fichier.
  echo.
  start "" "https://www.python.org/downloads/"
  pause
  exit /b 1
)

echo [1/3] Python detecte.

rem --- 2. Installer les bibliotheques optionnelles -----------------------
echo [2/3] Verification des bibliotheques (Word, Excel)...
%PY% -c "import docx, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo       Installation en cours ^(une seule fois^)...
  if exist "vendor\" (
    rem Installation hors-ligne depuis les paquets fournis
    %PY% -m pip install --no-index --find-links vendor python-docx openpyxl >nul 2>&1
  ) else (
    %PY% -m pip install --quiet python-docx openpyxl >nul 2>&1
  )
  if errorlevel 1 (
    echo       [i] Installation impossible. Le logiciel fonctionnera quand meme,
    echo           mais sans lecture .docx ni export Excel enrichi.
  ) else (
    echo       Bibliotheques installees.
  )
) else (
  echo       Deja installees.
)

rem --- 2b. OCR (optionnel) : lecture des scans CIN/recus -----------------
where tesseract >nul 2>&1
if errorlevel 1 (
  echo       [i] OCR non detecte. La lecture des documents scannes ^(CIN, recus^)
  echo           est desactivee. Pour l'activer : installez Tesseract OCR
  echo           ^(https://github.com/UB-Mannheim/tesseract/wiki^) puis relancez.
  echo           Le logiciel fonctionne tres bien sans.
) else (
  echo       OCR actif ^(Tesseract^).
)

rem --- 3. Lancer ---------------------------------------------------------
echo [3/3] Ouverture de l'interface dans le navigateur...
echo.
echo     Laissez cette fenetre OUVERTE pendant l'utilisation.
echo     Fermez-la pour quitter le logiciel.
echo.
set "PORT=8700"
%PY% app.py

pause
