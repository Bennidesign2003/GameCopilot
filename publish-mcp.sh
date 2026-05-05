#!/bin/bash
set -e

# ============================================================
# Pimax-Graphics-MCP — Publish Script
# Auto-increments __mcp_version__, commits + pushes to
# https://github.com/Bennidesign2003/Pimax-Graphics-MCP
# ============================================================

REPO="Bennidesign2003/Pimax-Graphics-MCP"
MCP_SRC="$(cd "$(dirname "$0")/GameCopilot/Assets" && pwd)/mcp-server.py"
WORK_DIR="$(mktemp -d)"
CLONE_DIR="${WORK_DIR}/Pimax-Graphics-MCP"

echo "=========================================="
echo "  Pimax-Graphics-MCP Publisher"
echo "=========================================="

# ── Prerequisites ────────────────────────────────────────────
if ! command -v gh &>/dev/null; then
    echo "ERROR: GitHub CLI (gh) not found. Install with: brew install gh"
    exit 1
fi
if ! gh auth status &>/dev/null; then
    echo "ERROR: Not logged in. Run: gh auth login"
    exit 1
fi
if [ ! -f "${MCP_SRC}" ]; then
    echo "ERROR: mcp-server.py not found at ${MCP_SRC}"
    exit 1
fi

# ── Read current version ─────────────────────────────────────
CURRENT_VERSION=$(grep -m1 '__mcp_version__' "${MCP_SRC}" | sed 's/.*"\(.*\)".*/\1/')
if [ -z "${CURRENT_VERSION}" ]; then
    echo "ERROR: Could not read __mcp_version__ from mcp-server.py"
    exit 1
fi
echo "[1/5] Current version: ${CURRENT_VERSION}"

# ── Auto-increment patch version ─────────────────────────────
MAJOR=$(echo "${CURRENT_VERSION}" | cut -d. -f1)
MINOR=$(echo "${CURRENT_VERSION}" | cut -d. -f2)
PATCH=$(echo "${CURRENT_VERSION}" | cut -d. -f3)
NEW_PATCH=$((PATCH + 1))
NEW_VERSION="${MAJOR}.${MINOR}.${NEW_PATCH}"
echo "[2/5] New version:     ${NEW_VERSION}"

# ── Update version in mcp-server.py ──────────────────────────
echo "[3/5] Updating __mcp_version__ in mcp-server.py..."
# macOS-compatible sed (no -i '' issue)
TMP_FILE="${WORK_DIR}/mcp-server.py"
sed "s/__mcp_version__ = \"${CURRENT_VERSION}\"/__mcp_version__ = \"${NEW_VERSION}\"/" "${MCP_SRC}" > "${TMP_FILE}"

# Verify the replacement worked
if ! grep -q "__mcp_version__ = \"${NEW_VERSION}\"" "${TMP_FILE}"; then
    echo "ERROR: Version replacement failed"
    rm -rf "${WORK_DIR}"
    exit 1
fi

# Write back to source
cp "${TMP_FILE}" "${MCP_SRC}"
echo "    Updated: ${CURRENT_VERSION} → ${NEW_VERSION}"

# ── Clone + push to Pimax-Graphics-MCP ───────────────────────
echo "[4/5] Cloning ${REPO}..."
gh repo clone "${REPO}" "${CLONE_DIR}" -- --depth=1

echo "    Copying mcp-server.py..."
cp "${MCP_SRC}" "${CLONE_DIR}/mcp-server.py"

# Update or create README badge line with version
README="${CLONE_DIR}/README.md"
if [ -f "${README}" ]; then
    # Replace version badge if present, else leave as-is
    sed -i.bak "s/version-[0-9]*\.[0-9]*\.[0-9]*/version-${NEW_VERSION}/g" "${README}" 2>/dev/null && rm -f "${README}.bak" || true
fi

cd "${CLONE_DIR}"
git add mcp-server.py README.md 2>/dev/null || git add mcp-server.py
git commit -m "chore: release v${NEW_VERSION}

Auto-published by publish-mcp.sh
Previous version: ${CURRENT_VERSION}"

git push origin main
echo "    Pushed to https://github.com/${REPO}"

# ── Create GitHub Release ─────────────────────────────────────
echo "[5/5] Creating GitHub release v${NEW_VERSION}..."
TAG="v${NEW_VERSION}"

if gh release view "${TAG}" --repo "${REPO}" &>/dev/null; then
    echo "    Release ${TAG} already exists — skipping"
else
    gh release create "${TAG}" \
        "${CLONE_DIR}/mcp-server.py" \
        --repo "${REPO}" \
        --title "Pimax-Graphics-MCP ${NEW_VERSION}" \
        --notes "## Pimax-Graphics-MCP ${NEW_VERSION}

Auto-release from publish-mcp.sh.

### Install / Update
The MCP server updates itself automatically via the AI chat.
To manually update: replace \`mcp-server.py\` in AppData with this release asset."
    echo "    Release created: https://github.com/${REPO}/releases/tag/${TAG}"
fi

# ── Cleanup ───────────────────────────────────────────────────
cd - >/dev/null
rm -rf "${WORK_DIR}"

echo ""
echo "=========================================="
echo "  DONE! v${NEW_VERSION} published"
echo "=========================================="
echo ""
echo "  Repo:    https://github.com/${REPO}"
echo "  Release: https://github.com/${REPO}/releases/tag/v${NEW_VERSION}"
echo "  Raw:     https://raw.githubusercontent.com/${REPO}/main/mcp-server.py"
echo ""
