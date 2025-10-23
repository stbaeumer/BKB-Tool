#!/bin/bash
set -e

VERSION=$1
if [ -z "$VERSION" ]; then
    echo "Usage: $0 <version>"
    exit 1
fi

echo "Creating AppImage for version $VERSION"

# AppImage-Tool herunterladen
if [ ! -f "/tmp/appimagetool" ]; then
    wget -q https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage -O /tmp/appimagetool
    chmod +x /tmp/appimagetool
fi

# AppDir erstellen
rm -rf AppDir
mkdir -p AppDir/usr/bin
mkdir -p AppDir/usr/share/icons/hicolor/256x256/apps
mkdir -p AppDir/usr/share/icons/hicolor/512x512/apps
mkdir -p AppDir/usr/share/applications

# Binary kopieren
cp ./publish/linux/BKB-Tool AppDir/usr/bin/
chmod +x AppDir/usr/bin/BKB-Tool

# Icons kopieren
cp BKB-Tool/Icons/BKB-Tool_256.png AppDir/usr/share/icons/hicolor/256x256/apps/bkb-tool.png
cp BKB-Tool/Icons/BKB-Tool_512.png AppDir/usr/share/icons/hicolor/512x512/apps/bkb-tool.png
cp BKB-Tool/Icons/BKB-Tool_256.png AppDir/bkb-tool.png

# Desktop-Datei erstellen
cat > AppDir/bkb-tool.desktop << 'EOF'
[Desktop Entry]
Type=Application
Name=BKB-Tool
Comment=Ein Werkzeug an der Schnittstelle zwischen SchILD und Webuntis
Exec=BKB-Tool
Icon=bkb-tool
Categories=Utility;Education;Office;
Terminal=false
StartupNotify=true
Version=1.0
EOF

cp AppDir/bkb-tool.desktop AppDir/usr/share/applications/

# AppRun erstellen
cat > AppDir/AppRun << 'EOF'
#!/bin/bash
SELF=$(readlink -f "$0")
HERE=${SELF%/*}
export PATH="${HERE}/usr/bin:${PATH}"
export LD_LIBRARY_PATH="${HERE}/usr/lib:${LD_LIBRARY_PATH}"
exec "${HERE}/usr/bin/BKB-Tool" "$@"
EOF
chmod +x AppDir/AppRun

# AppImage erstellen
ARCH=x86_64 /tmp/appimagetool AppDir "BKB-Tool-${VERSION}.AppImage"

echo "AppImage created: BKB-Tool-${VERSION}.AppImage"