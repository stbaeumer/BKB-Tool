@echo off
echo.
echo BKB-Tool
echo =========
echo Warte auf Beenden von BKB-Tool.exe ...
:waitforend
tasklist | find /I "BKB-Tool.exe" >nul
if not errorlevel 1 (
    timeout /t 1 >nul
    goto waitforend
)
echo Ersetze alte Version ...
del /F /Q BKB-Tool.exe
rename BKB-Tool_neu.exe BKB-Tool.exe
echo Starte neue Version ...
start "" BKB-Tool.exe
exit
