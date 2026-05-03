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
VERSION="3.5.4"
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
    "HOTFIX: App-Crash beim Start in 3.5.3 behoben (FormatException 'Easing Linear was not found')",
    "(aus 3.5.3) MCP-Server wird beim App-Start vorgewaermt",
    "(aus 3.5.3) Halluzinationen behoben durch groesseres Kontext-Fenster (16k)",
    "(aus 3.5.3) Tool-Auswahl deterministischer (temperature 0.1), finale Antwort natuerlicher (0.65)",
    "(aus 3.5.3) Klare Fehlermeldung wenn der Agent nach 15 Tool-Runden ohne Antwort endet",
    "(aus 3.5.3) Live-Streaming Indikator (typing dots) waehrend die Antwort eintrifft",
    "(aus 3.5.2) Neue Modelle: Qwen 3, Qwen 2.5 Coder, GPT-OSS, Llama 3.3, Gemma 3",
    "(aus 3.5.2) Auto-Modellwahl basierend auf erkanntem VRAM",
    "(aus 3.5.2) Modell bleibt 30 Min im VRAM zwischen Fragen (keep_alive)",
    "(aus 3.5.2) Streaming auch im Tool-Modus statt 30s Spinner"
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
        --notes "## GameCopilot ${VERSION} — Hotfix

**3.5.3 stuerzt beim Start ab** (\`FormatException: Easing 'Linear' was not found\`). Bitte direkt auf 3.5.4 updaten.

### Hotfix
- Streaming-Cursor-Animation in MainWindow.axaml verwendete \`Easing=\"Linear\"\` - Avalonia erwartet aber \`LinearEasing\` (Suffix). Der XAML-Compiler hat das durchgelassen, weil \`Easing.Parse\` erst zur Laufzeit aufgeloest wird, und die App ist daher in 3.5.3 sofort beim Window-Init gecrasht.

### Aus 3.5.3 (jetzt nutzbar)
- MCP-Server wird waehrend des Splash-Screens vorgewaermt (erste Tool-Anfrage sofort schnell)
- Halluzinationen behoben durch groesseres Kontext-Fenster (\`num_ctx\` 8192 → 16384) im Tool-Modus
- Tool-Auswahl deterministischer (temperature 0.1), finale Antwort natuerlicher (temperature 0.65)
- System-Prompt erweitert um \`diagnose_msfs_config\`, \`set_msfs_setting\`, \`fix_msfs\`, \`check_and_install_driver\`
- Klare Fehlermeldung wenn der Agent alle 15 Tool-Runden braucht ohne fertig zu werden
- Streaming-Indikator (typing dots) ab dem ersten Token

### Aus 3.5.2
- Neue Modelle: Qwen 3 (4B/8B/14B/30B-MoE/32B), Qwen 2.5 Coder (7B/14B/32B), GPT-OSS 20B/120B, Llama 3.3 70B, Gemma 3 12B/27B
- Auto-Modellwahl basierend auf erkanntem VRAM
- \`keep_alive 30m\` haelt Modell zwischen Fragen im VRAM
- Streaming auch im Tool-Modus, finale Antwort Wort-fuer-Wort statt 30s Spinner
- Komplett neu geschriebener System-Prompt (Identitaet / Stil / Ehrlichkeit / Tools)"
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
