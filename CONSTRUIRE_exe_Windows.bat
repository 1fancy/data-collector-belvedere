@echo off
chcp 65001 >nul
title Construction de Data Collector (.exe)
cd /d "%~dp0"

echo ============================================================
echo    Construction de "Data Collector.exe"
echo    (a faire UNE SEULE FOIS, sur un PC Windows)
echo ============================================================
echo.

rem --- Python ---
set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )
if not defined PY (
  echo [!] Python manquant. Installez-le depuis https://www.python.org/downloads/
  echo     (cochez "Add Python to PATH"), puis relancez ce fichier.
  start "" "https://www.python.org/downloads/"
  pause & exit /b 1
)

echo [1/3] Installation des outils de construction...
%PY% -m pip install --quiet --upgrade pip
%PY% -m pip install --quiet pyinstaller python-docx openpyxl
if errorlevel 1 (
  echo [!] Installation impossible (verifiez la connexion internet).
  pause & exit /b 1
)

echo [2/3] Construction de l'executable (quelques minutes)...
%PY% -m PyInstaller --noconfirm DataCollector.spec
if errorlevel 1 (
  echo [!] La construction a echoue.
  pause & exit /b 1
)

echo [3/3] Termine.
echo.
echo     Votre fichier est pret :  dist\Data Collector.exe
echo     Copiez-le ou vous voulez et double-cliquez dessus.
echo.
if exist "dist\Data Collector.exe" explorer dist
pause
