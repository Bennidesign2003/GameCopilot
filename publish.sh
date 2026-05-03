#!/bin/bash
set -e

# ============================================================
# GameCopilot Publish & Release Script
# Builds Windows x64 exe, creates update.json, uploads to
# Bennidesign2003/GodotRenderingAI releases
#
# The old WPF updater expects:
#   - Zip named: MSFS24.Game.Manager.Update.zip
#   - Folder inside: MSFS24.Game.Manager.Update/
#   - Exe named: MSFS Mod Manager.exe
#   - Downloaded from MSFS24 release tag
# ============================================================

REPO="Bennidesign2003/GodotRenderingAI"
VERSION="3.5.2"
TAG="v${VERSION}"
PROJECT_DIR="$(cd "$(dirname "$0")/GameCopilot" && pwd)"
PUBLISH_DIR="${PROJECT_DIR}/bin/publish"
UPDATE_DIR="${PUBLISH_DIR}/MSFS24.Game.Manager.Update"
ZIP_NAME="MSFS24.Game.Manager.Update.zip"
ZIP_PATH="${PUBLISH_DIR}/${ZIP_NAME}"
UPDATE_JSON="${PUBLISH_DIR}/update.json"
DOWNLOAD_URL="https://github.com/${REPO}/releases/download/MSFS24/${ZIP_NAME}"

echo "=========================================="
echo "  GameCopilot Publish v${VERSION}"
echo "=========================================="

# Check prerequisites
if ! command -v dotnet &>/dev/null; then
    echo "ERROR: dotnet SDK not found. Install .NET 8 SDK first."
    exit 1
fi

if ! command -v gh &>/dev/null; then
    echo "ERROR: GitHub CLI (gh) not found. Install with: brew install gh"
    exit 1
fi

if ! gh auth status &>/dev/null; then
    echo "ERROR: Not logged into GitHub. Run: gh auth login"
    exit 1
fi

# Step 1: Clean previous build
echo ""
echo "[1/6] Cleaning previous build..."
rm -rf "${PUBLISH_DIR}"
mkdir -p "${PUBLISH_DIR}"

# Step 2: Publish for Windows x64 (self-contained single file)
echo "[2/6] Building Windows x64 release..."
dotnet publish "${PROJECT_DIR}/GameCopilot.csproj" \
    -c Release \
    -r win-x64 \
    -o "${PUBLISH_DIR}/app" \
    -p:Version="${VERSION}" \
    -p:PublishSingleFile=true \
    -p:SelfContained=true \
    -p:IncludeNativeLibrariesForSelfExtract=true

echo "    Build output:"
ls -lh "${PUBLISH_DIR}/app/GameCopilot.exe"

# Step 3: Create update folder with old updater-compatible structure
echo "[3/6] Creating update package (old updater compatible)..."
rm -rf "${UPDATE_DIR}"
mkdir -p "${UPDATE_DIR}"
cp "${PUBLISH_DIR}/app/GameCopilot.exe" "${UPDATE_DIR}/MSFS Mod Manager.exe"
cat > "${UPDATE_DIR}/appconfig.json" <<EOF
{
  "CurrentVersion": "${VERSION}"
}
EOF

# Step 4: Create zip with correct folder structure
echo "[4/6] Creating zip archive..."
cd "${PUBLISH_DIR}"
zip -r "${ZIP_PATH}" "MSFS24.Game.Manager.Update/" -x "*.pdb"
cd - >/dev/null
echo "    Archive: ${ZIP_PATH}"
ls -lh "${ZIP_PATH}"

# Step 5: Generate update.json
echo "[5/6] Generating update.json..."
cat > "${UPDATE_JSON}" <<EOF
{
  "LatestVersion": "${VERSION}",
  "DownloadUrl": "${DOWNLOAD_URL}",
  "Changelog": [
    "Bugfix: MCP startet automatisch wenn nicht aktiv - AI nutzt immer Tools",
    "Bugfix: UI blockiert nicht mehr bei Tool-Aufrufen (Background-Thread)",
    "Bugfix: AI ruft Tools einzeln auf statt alles auf einmal",
    "Bugfix: Weniger Tools (10 statt 22) fuer schnellere Verarbeitung",
    "Bugfix: OpenXR Toolkit Auto-Discovery (alle MSFS exe-Namen)"
  ]
}
EOF
echo "    update.json created"
cat "${UPDATE_JSON}"

# Step 6: Upload to MSFS24 release (where old app checks for updates)
echo ""
echo "[6/6] Uploading to MSFS24 release..."
if gh release view "MSFS24" --repo "${REPO}" &>/dev/null; then
    gh release upload "MSFS24" \
        "${ZIP_PATH}" \
        "${UPDATE_JSON}" \
        --repo "${REPO}" \
        --clobber
else
    gh release create "MSFS24" \
        "${ZIP_PATH}" \
        "${UPDATE_JSON}" \
        --repo "${REPO}" \
        --title "MSFS24 Update Channel" \
        --notes "Auto-update channel for GameCopilot MSFS 2024 edition."
fi

# Also create a versioned release tag
echo "    Creating versioned release ${TAG}..."
if gh release view "${TAG}" --repo "${REPO}" &>/dev/null; then
    gh release upload "${TAG}" \
        "${ZIP_PATH}" \
        "${UPDATE_JSON}" \
        --repo "${REPO}" \
        --clobber
else
    gh release create "${TAG}" \
        "${ZIP_PATH}" \
        "${UPDATE_JSON}" \
        --repo "${REPO}" \
        --title "GameCopilot ${VERSION}" \
        --notes "## GameCopilot ${VERSION}

### Bugfixes
- **AI erfindet keine Daten mehr**: Wenn MCP-Tools Fehler zurueckgeben (ReShade nicht installiert, OpenXR Toolkit nicht gefunden, Pimax Config fehlt), zeigt die AI jetzt ehrlich den Fehler an statt Tabellen mit erfundenen Werten
- **Tool-Fehler sichtbar**: Agent-Steps zeigen Warnungs-Icon bei Fehlern statt gruenen Haken
- **Qwen3 /no_think**: Schnellere Tool-Aufrufe durch deaktivierten Thinking-Modus

### Neu
- **Bessere Steam Library Erkennung**: ReShade-Suche liest libraryfolders.vdf und findet MSFS auf allen Laufwerken automatisch"
fi

echo ""
echo "=========================================="
echo "  DONE! Release ${TAG} published"
echo "=========================================="
echo ""
echo "  MSFS24 tag: https://github.com/${REPO}/releases/tag/MSFS24"
echo "  Version tag: https://github.com/${REPO}/releases/tag/${TAG}"
echo "  Download: ${DOWNLOAD_URL}"
echo ""
