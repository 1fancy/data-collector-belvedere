@echo off
chcp 65001 >nul
title Data Collector - Installation
color 0F
cd /d "%~dp0"

echo.
echo   ============================================================
echo      DATA COLLECTOR - Romana Immobilier
echo      Installation de l'application
echo   ============================================================
echo.

rem ---------------------------------------------------------------
rem  1) Code d'installation (empeche une utilisation non autorisee)
rem ---------------------------------------------------------------
set /p CODE=  Entrez le code d'installation :
if not "%CODE%"=="1234" (
  echo.
  echo   [X] Code incorrect. Installation annulee.
  echo.
  pause
  exit /b 1
)
echo.
echo   Code accepte.
echo.

rem ---------------------------------------------------------------
rem  Parametres (depot public contenant UNIQUEMENT le code)
rem ---------------------------------------------------------------
set "REPO_ZIP=https://github.com/1fancy/data-collector-belvedere/archive/refs/heads/main.zip"
set "WORKDIR=%CD%\DataCollector_build"
set "APPDIR=%CD%\Data Collector"

rem ---------------------------------------------------------------
rem  2) Verifier Python
rem ---------------------------------------------------------------
set "PY="
where py  >nul 2>&1 && set "PY=py"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )
if not defined PY (
  echo   [!] Python n'est pas installe.
  echo       Ouverture de la page de telechargement...
  echo       Cochez "Add Python to PATH" puis relancez ce fichier.
  start "" "https://www.python.org/downloads/"
  pause
  exit /b 1
)
echo   [1/5] Python detecte.

rem ---------------------------------------------------------------
rem  3) Telecharger le code depuis GitHub (zip, sans git)
rem ---------------------------------------------------------------
echo   [2/5] Telechargement de l'application...
if exist "%WORKDIR%" rmdir /s /q "%WORKDIR%"
mkdir "%WORKDIR%"
powershell -NoProfile -Command ^
  "try { Invoke-WebRequest -Uri '%REPO_ZIP%' -OutFile '%WORKDIR%\src.zip' -UseBasicParsing } catch { exit 1 }"
if errorlevel 1 (
  echo   [X] Telechargement impossible. Verifiez la connexion internet.
  pause & exit /b 1
)
echo   [3/5] Extraction...
powershell -NoProfile -Command ^
  "Expand-Archive -Path '%WORKDIR%\src.zip' -DestinationPath '%WORKDIR%' -Force"
rem Le zip contient un sous-dossier <repo>-main : on le retrouve.
for /d %%D in ("%WORKDIR%\*-main") do set "SRC=%%D"

rem ---------------------------------------------------------------
rem  4) Installer les dependances + construire l'exe
rem ---------------------------------------------------------------
echo   [4/5] Preparation (quelques minutes la premiere fois)...
%PY% -m pip install --quiet --upgrade pip
%PY% -m pip install --quiet pyinstaller python-docx openpyxl
pushd "%SRC%"
%PY% -m PyInstaller --noconfirm DataCollector.spec
popd
if not exist "%SRC%\dist\Data Collector\Data Collector.exe" (
  echo   [X] La construction a echoue.
  pause & exit /b 1
)

rem ---------------------------------------------------------------
rem  5) Installer proprement + archiver les fichiers de build
rem ---------------------------------------------------------------
echo   [5/5] Finalisation...
if exist "%APPDIR%" rmdir /s /q "%APPDIR%"
move "%SRC%\dist\Data Collector" "%APPDIR%" >nul

rem Archive des fichiers de construction (cache, zip, build) dans "installer\"
if not exist "%CD%\installer" mkdir "%CD%\installer"
move "%WORKDIR%" "%CD%\installer\build_%RANDOM%" >nul 2>&1

rem Raccourci sur le Bureau
set "LNK=%USERPROFILE%\Desktop\Data Collector.lnk"
powershell -NoProfile -Command ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%LNK%'); $s.TargetPath='%APPDIR%\Data Collector.exe'; $s.WorkingDirectory='%APPDIR%'; $s.IconLocation='%APPDIR%\Data Collector.exe'; $s.Save()"

echo.
echo   ============================================================
echo      Installation terminee !
echo      L'application est dans le dossier : "Data Collector"
echo      Un raccourci a ete cree sur le Bureau.
echo   ============================================================
echo.
echo   Lancement de l'application...
start "" "%APPDIR%\Data Collector.exe"
timeout /t 3 >nul
exit /b 0
