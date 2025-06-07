chcp 65001
cd BKB-Tool
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