#!/bin/bash
set -e

VERSION=$1
if [ -z "$VERSION" ]; then
    echo "Usage: $0 <version>"
    exit 1
fi

APP_NAME="BKB-Tool"
RUNTIME="linux-x64"
CONFIGURATION="Release"
OUTPUT_DIR="dist/${APP_NAME}-${VERSION}-${RUNTIME}"

echo "Building ${APP_NAME} version ${VERSION}"

# Clean
rm -rf dist
mkdir -p "$OUTPUT_DIR"

# .NET publish
dotnet publish \
    ./BKB-Tool/BKB-Tool.csproj \
    -c "$CONFIGURATION" \
    -r "$RUNTIME" \
    --self-contained true \
    /p:Version="$VERSION" \
    -o "$OUTPUT_DIR"

# Ensure executable
chmod +x "$OUTPUT_DIR/$APP_NAME"

# Optional: Archiv für GitHub Release / Artifact
ARCHIVE_NAME="${APP_NAME}-${VERSION}-${RUNTIME}.tar.gz"
tar -czf "dist/$ARCHIVE_NAME" -C dist "$(basename "$OUTPUT_DIR")"

echo "Build completed:"
echo " - Output directory: $OUTPUT_DIR"
echo " - Archive: dist/$ARCHIVE_NAME"
