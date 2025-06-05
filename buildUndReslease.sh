#!/bin/bash

set -e

# === Konfiguration ===
VERSION="v$(date +%Y%m%d%H%M%S)"
REPO="stbaeumer/BKB-Tool"
RELEASE_DIR="/home/stefan/RiderProjects/BKB-Tool/releases/$VERSION"
mkdir -p "$RELEASE_DIR"

# === Letzten Git-Tag finden und Changelog generieren ===
cd /home/stefan/RiderProjects/BKB-Tool
LAST_TAG=$(git describe --tags --abbrev=0 2>/dev/null || echo "")
if [ -n "$LAST_TAG" ]; then
    CHANGELOG=$(git log "$LAST_TAG"..HEAD --pretty=format:"- %s" --no-merges)
else
    CHANGELOG=$(git log --pretty=format:"- %s" --no-merges)
fi
if [ -z "$CHANGELOG" ]; then
    CHANGELOG="Keine Änderungen seit dem letzten Release."
fi

# === Build-Prozess ===
cd BKB-Tool
/home/stefan/.dotnet/dotnet publish -c Release -r win-x64 --self-contained true
cp -r bin/Release/net8.0/win-x64/publish/* "$RELEASE_DIR/"
/home/stefan/.dotnet/dotnet publish -c Release -r linux-x64 --self-contained true 
cp -r bin/Release/net8.0/linux-x64/publish/* "$RELEASE_DIR/"
chmod +x "$RELEASE_DIR/BKB-Tool"

# === ZIP erstellen ===
cd "$RELEASE_DIR"
zip -r "../BKB-${VERSION}.zip" *

# === GitHub Release erstellen ===
cd ..
gh release create "$VERSION" "BKB-${VERSION}.zip" \
  --repo "$REPO" \
  --title "Version $VERSION" \
  --notes "$CHANGELOG"

echo "✅ Veröffentlichung & GitHub Release abgeschlossen für $VERSION!"

