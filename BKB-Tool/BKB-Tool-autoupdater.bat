@echo off
echo.
echo BKB-Tool
echo =========
echo Warte auf Beenden von BKB-Tool.exe ...
set /a counter=0
:waitforend
tasklist | find /I "BKB-Tool.exe" >nul
if not errorlevel 1 (
    timeout /t 1 >nul
    set /a counter+=1
    if %counter%==7 (
        echo Das Beenden dauert zu lange. Versuchen Sie BKB-Tool im Taskmanager zu beenden oder starten Sie den Rechner neu.
    )
    goto waitforend
)
echo Ersetze alte Version ...
del /F /Q BKB-Tool.exe
if exist BKB-Tool.exe (
    echo Fehler: Die alte BKB-Tool.exe konnte nicht gelöscht werden. Bitte schließen Sie alle Instanzen und versuchen Sie es erneut.
    pause
    exit /b 1
)
rename BKB-Tool_neu.exe BKB-Tool.exe
echo Starte neue Version ...
start "" BKB-Tool.exe
exit
