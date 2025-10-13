@echo off
chcp 65001
cd BKB-Tool

REM ---------------------------------------------------------
REM Beende laufendes BKB-Tool.exe
REM ---------------------------------------------------------
echo Prüfe, ob BKB-Tool.exe läuft ...
tasklist | find /I "BKB-Tool.exe" >nul
if not errorlevel 1 (
    echo Beende BKB-Tool.exe ...
    taskkill /IM "BKB-Tool.exe" /F
    :waitforend
    tasklist | find /I "BKB-Tool.exe" >nul
    if not errorlevel 1 (
        timeout /t 1 >nul
        goto waitforend
    )
)

REM ---------------------------------------------------------
REM Windows-Build
REM ---------------------------------------------------------
echo.
echo === Erstelle Windows x64 Build ===
dotnet publish -c Release -r win-x64 --self-contained true ^
    /p:PublishSingleFile=true /p:IncludeAllContentForSelfExtract=true

REM ---------------------------------------------------------
REM Linux-Build
REM ---------------------------------------------------------
echo.
echo === Erstelle Linux x64 Build ===
dotnet publish -c Release -r linux-x64 --self-contained true ^
    /p:PublishSingleFile=true /p:IncludeAllContentForSelfExtract=true

REM ---------------------------------------------------------
REM Ziel- und Quellverzeichnisse
REM ---------------------------------------------------------
set ZIEL=C:\Users\bm\BKB-Tool\
set QUELLE_WIN=bin\Release\net8.0\win-x64\publish
set QUELLE_LINUX=bin\Release\net8.0\linux-x64\publish

echo Lösche alle Dateien im Zielverzeichnis...
del /Q %ZIEL%*

REM ---------------------------------------------------------
REM Windows-Build kopieren
REM ---------------------------------------------------------
echo Kopiere Windows-Build nach Zielverzeichnis...
robocopy %QUELLE_WIN% %ZIEL% /XO /E /R:3 /W:5

REM ---------------------------------------------------------
REM Linux-Build kopieren (gleicher Ordner)
REM ---------------------------------------------------------
echo Kopiere Linux-Build nach Zielverzeichnis...
robocopy %QUELLE_LINUX% %ZIEL% /XO /E /R:3 /W:5

REM ---------------------------------------------------------
REM .env kopieren
REM ---------------------------------------------------------
copy /Y ".env" "%ZIEL%"

REM ---------------------------------------------------------
REM Icon-Datei kopieren
REM ---------------------------------------------------------
if exist "BKB-Tool_256.png" (
    echo Kopiere Icon BKB-Tool_256.png ins Zielverzeichnis...
    copy /Y "BKB-Tool_256.png" "%ZIEL%"
    set ICON_PATH=%ZIEL%BKB-Tool_256.png
) else (
    echo WARNUNG: Icon BKB-Tool_256.png nicht gefunden! Desktop-Datei wird nicht erstellt.
    set ICON_PATH=
)

REM ---------------------------------------------------------
REM BKB-Tool.desktop erstellen (nur wenn Icon gefunden)
REM ---------------------------------------------------------
if defined ICON_PATH (
    echo Erstelle BKB-Tool.desktop ...
    (
    echo [Desktop Entry]
    echo Type=Application
    echo Name=BKB-Tool
    echo Exec=%ZIEL%BKB-Tool
    echo Icon=%ICON_PATH%
    echo Comment=BKB-Tool
    echo Terminal=false
    echo Categories=Utility;
    ) > "%ZIEL%BKB-Tool.desktop"
)

REM ---------------------------------------------------------
REM Abschluss
REM ---------------------------------------------------------
echo Zielordner öffnen...
start "" %ZIEL%

echo.
echo Publishing abgeschlossen!
pause
