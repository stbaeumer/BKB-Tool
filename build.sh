#!/bin/bash

cd /home/stefan/RiderProjects/BKB-Tool/BKB-Tool
/home/stefan/.dotnet/dotnet publish -c Release -r win-x64 --self-contained true
cp -r /home/stefan/RiderProjects/BKB-Tool/BKB-Tool/bin/Release/net8.0/win-x64/publish/* /home/stefan/Windows/SchILD-NRW/publish/
/home/stefan/.dotnet/dotnet publish -c Release -r linux-x64 --self-contained true 
cp -r /home/stefan/RiderProjects/BKB-Tool/BKB-Tool/bin/Release/net8.0/linux-x64/publish/* /home/stefan/bin/BKB-Tool/

chmod +x /home/stefan/bin/BKB-Tool/BKB-Tool

echo "Veröffentlichung abgeschlossen!"

