chcp 65001
cd BKB-Tool

REM BKB-Tool.exe beenden, falls offen
echo Prüfe, ob BKB-Tool.exe läuft ...
tasklist | find /I "BKB-Tool.exe" >nul
if not errorlevel 1 (
    echo Beende BKB-Tool.exe ...
    taskkill /IM "BKB-Tool.exe" /F
    REM Warten, bis Prozess wirklich beendet ist
    :waitforend
    tasklist | find /I "BKB-Tool.exe" >nul
    if not errorlevel 1 (
        timeout /t 1 >nul
        goto waitforend
    )
)

dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeAllContentForSelfExtract=true

REM Zielverzeichnis definieren
set ZIEL=C:\Users\bm\BKB-Tool\

REM Quellverzeichnis definieren
set QUELLE=bin\Release\net8.0\win-x64\publish

echo Lösche alle Dateien im Zielverzeichnis...
del /Q %ZIEL%*

robocopy %QUELLE% %ZIEL% /XO /E /R:3 /W:5

echo Zielordner öffnen...
start "" %ZIEL%


echo Publishing abgeschlossen!