@echo off
chcp 65001 >nul
title Data Collector - Mise a jour
cd /d "%~dp0"

echo.
echo   ============================================================
echo      DATA COLLECTOR - Mise a jour
echo   ============================================================
echo.
echo   Cette operation telecharge la derniere version et reconstruit
echo   l'application. VOS DONNEES (projets, clients) ne sont PAS
echo   touchees : elles restent dans le dossier "Data Collector".
echo.
set /p OK=  Continuer ? (O/N) :
if /i not "%OK%"=="O" ( echo Annule. & timeout /t 2 >nul & exit /b 0 )
echo.

set "REPO_ZIP=https://github.com/1fancy/data-collector-belvedere/archive/refs/heads/main.zip"
set "WORKDIR=%CD%\DataCollector_update"
set "APPDIR=%CD%\Data Collector"

set "PY="
where py  >nul 2>&1 && set "PY=py"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )
if not defined PY ( echo [!] Python manquant. & pause & exit /b 1 )

echo   [1/4] Telechargement de la derniere version...
if exist "%WORKDIR%" rmdir /s /q "%WORKDIR%"
mkdir "%WORKDIR%"
powershell -NoProfile -Command ^
  "try { Invoke-WebRequest -Uri '%REPO_ZIP%' -OutFile '%WORKDIR%\src.zip' -UseBasicParsing } catch { exit 1 }"
if errorlevel 1 ( echo [X] Telechargement impossible. & pause & exit /b 1 )
powershell -NoProfile -Command "Expand-Archive -Path '%WORKDIR%\src.zip' -DestinationPath '%WORKDIR%' -Force"
for /d %%D in ("%WORKDIR%\*-main") do set "SRC=%%D"

echo   [2/4] Mise a jour des outils...
%PY% -m pip install --quiet --upgrade pyinstaller python-docx openpyxl rapidocr_onnxruntime

echo   [3/4] Reconstruction (quelques minutes)...
pushd "%SRC%"
%PY% -m PyInstaller --noconfirm DataCollector.spec
popd
if not exist "%SRC%\dist\Data Collector\Data Collector.exe" (
  echo   [X] La reconstruction a echoue. L'ancienne version reste en place.
  pause & exit /b 1
)

echo   [4/4] Remplacement de l'application (donnees conservees)...
rem Sauvegarder les donnees utilisateur de l'ancienne version
set "KEEP=%WORKDIR%\keep"
mkdir "%KEEP%" 2>nul
for %%F in ("belvedere.db" "app_config.json" "data.json" "BELVEDERE_clients.xlsx") do (
  if exist "%APPDIR%\%%~F" copy /y "%APPDIR%\%%~F" "%KEEP%\%%~F" >nul
)
for %%D in ("Projets" ".ocr_cache") do (
  if exist "%APPDIR%\%%~D" xcopy /e /i /y /q "%APPDIR%\%%~D" "%KEEP%\%%~D" >nul
)
rem Remplacer l'app
if exist "%APPDIR%" rmdir /s /q "%APPDIR%"
move "%SRC%\dist\Data Collector" "%APPDIR%" >nul
rem Restaurer les donnees
for %%F in ("belvedere.db" "app_config.json" "data.json" "BELVEDERE_clients.xlsx") do (
  if exist "%KEEP%\%%~F" copy /y "%KEEP%\%%~F" "%APPDIR%\%%~F" >nul
)
for %%D in ("Projets" ".ocr_cache") do (
  if exist "%KEEP%\%%~D" xcopy /e /i /y /q "%KEEP%\%%~D" "%APPDIR%\%%~D" >nul
)
rmdir /s /q "%WORKDIR%" 2>nul

echo.
echo   Mise a jour terminee ! Vos donnees ont ete conservees.
echo.
start "" "%APPDIR%\Data Collector.exe"
timeout /t 3 >nul
exit /b 0
