cd BKB-Tool
dotnet publish -c Release -r win-x64 --self-contained true

REM Zielverzeichnis definieren
set ZIEL=\\fs01\SchILD-NRW\publish

REM Quellverzeichnis definieren
set QUELLE=bin\Release\net8.0\win-x64\publish

robocopy %QUELLE% %ZIEL% /XO /E /R:3 /W:5

echo Veröffentlichung abgeschlossen!