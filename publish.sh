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
    "Neue Modelle im Picker: Qwen 3 (4B/8B/14B/30B-MoE/32B), Qwen 2.5 Coder, GPT-OSS 20B/120B, Llama 3.3 70B, Gemma 3 12B/27B",
    "Auto-Modellwahl beim Start: VRAM/Unified-Memory wird erkannt und das passende Modell vorausgewaehlt",
    "Live-Streaming auch im Tool-Modus: finale Antwort erscheint Wort-fuer-Wort statt nach 30s am Stueck",
    "Schnellere Folge-Anfragen: Modell bleibt 30 Min im VRAM (keep_alive) - kein Cold-Load mehr zwischen Fragen",
    "Professionellerer System-Prompt: klare Rolle, Markdown-Tabellen mit Einheiten, Diagnose vor Empfehlung, keine erfundenen Daten"
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

Komplett ueberarbeiteter AI-Chat - schneller, professioneller, modernere Modelle.

### Neue Modelle
- **Qwen 3** in 4B / 8B / 14B / **30B-A3B (MoE)** / 32B
- **Qwen 2.5 Coder** in 7B / 14B / 32B fuer Code- und Tool-Workloads
- **GPT-OSS** 20B und 120B (OpenAI Open-Weights)
- **Llama 3.3** 70B (Meta-Flagship)
- **Gemma 3** 12B und 27B mit 128k Kontext
- DeepSeek R1 8B / 14B / 32B bleiben fuer Reasoning

### Schneller
- **Auto-Modellwahl beim Start**: VRAM (NVIDIA via nvidia-smi) bzw. Unified Memory (Apple Silicon via sysctl) wird erkannt und das passende Modell aus der Liste vorausgewaehlt - Status zeigt z.B. \`qwen3:14b bereit · 24 GB erkannt\`
- **keep_alive 30m**: Modell bleibt 30 Min nach jedem Chat im VRAM. Folgefragen starten ohne 5-30s Cold-Load
- **Streaming auch im Tool-Modus**: die finale Antwort erscheint Wort-fuer-Wort statt nach 30s am Stueck. Thinking-Spinner verschwindet beim ersten Token

### Professioneller
- Neuer System-Prompt mit vier expliziten Sektionen (Identitaet / Stil / Ehrlichkeit / Tools)
- Tool-Output immer als Markdown-Tabelle \`| Einstellung | Aktuell | Empfehlung |\`
- Werte mit Einheiten, Diagnose vor Empfehlung, klare Annahmen-Markierung
- \`/no_think\` fuer Qwen3 im Tool-Modus bleibt - vermeidet versteckte Reasoning-Tokens"
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
