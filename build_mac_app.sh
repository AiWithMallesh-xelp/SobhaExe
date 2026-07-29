#!/usr/bin/env bash
# build_mac_app.sh — Build Sobha Reconciliation.app for macOS
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

APP_NAME="Sobha Reconciliation"
APP_PATH="dist/${APP_NAME}.app"
MACOS_DIR="${APP_PATH}/Contents/MacOS"

echo ""
echo "============================================="
echo "  Building ${APP_NAME}.app"
echo "============================================="
echo ""

# ── 1. Virtual environment ───────────────────────────────────
if [[ ! -d "venv" ]]; then
  echo "[1/7] Creating virtual environment..."
  python3 -m venv venv
else
  echo "[1/7] Virtual environment already exists."
fi

# shellcheck disable=SC1091
source venv/bin/activate

# ── 2. Dependencies ────────────────────────────────────────────
echo "[2/7] Installing dependencies..."
python -m pip install --upgrade pip --quiet
pip install -r requirements.txt --quiet
pip install pyinstaller --quiet

# ── 3. Icon asset ────────────────────────────────────────────
echo "[3/7] Preparing icon..."
if [[ ! -f "SobhaReconciliation.icns" ]]; then
  if [[ ! -f "sobha_logo_brand.png" ]]; then
    echo "ERROR: sobha_logo_brand.png not found." >&2
    exit 1
  fi
  mkdir -p SobhaReconciliation.iconset
  cp sobha_logo_brand.png SobhaReconciliation.iconset/icon_512x512@2x.png
  sips -z 16 16   sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_16x16.png >/dev/null
  sips -z 32 32   sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_16x16@2x.png >/dev/null
  sips -z 32 32   sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_32x32.png >/dev/null
  sips -z 64 64   sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_32x32@2x.png >/dev/null
  sips -z 128 128 sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_128x128.png >/dev/null
  sips -z 256 256 sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_128x128@2x.png >/dev/null
  sips -z 256 256 sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_256x256.png >/dev/null
  sips -z 512 512 sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_256x256@2x.png >/dev/null
  sips -z 512 512 sobha_logo_brand.png --out SobhaReconciliation.iconset/icon_512x512.png >/dev/null
  iconutil -c icns SobhaReconciliation.iconset -o SobhaReconciliation.icns
  rm -rf SobhaReconciliation.iconset
fi

# ── 4. Playwright Chromium (portable) ────────────────────────
echo "[4/7] Ensuring Playwright Chromium in pw-browsers/..."
export PLAYWRIGHT_BROWSERS_PATH="${ROOT}/pw-browsers"
python -m playwright install chromium

# ── 5. PyInstaller build ─────────────────────────────────────
echo "[5/7] Running PyInstaller..."
pyinstaller sobha_reconciliation.spec --noconfirm --clean

if [[ ! -d "$APP_PATH" ]]; then
  echo "ERROR: ${APP_PATH} was not created." >&2
  exit 1
fi

# ── 6. Post-build copy (browsers + logo next to executable) ──
echo "[6/7] Copying pw-browsers and logo into app bundle..."
rm -rf "${MACOS_DIR}/pw-browsers"
cp -R pw-browsers "${MACOS_DIR}/"
cp sobha_logo_brand.png "${MACOS_DIR}/"

# ── 7. Gatekeeper / ad-hoc sign ──────────────────────────────
echo "[7/7] Clearing quarantine and ad-hoc signing..."
xattr -cr "$APP_PATH" 2>/dev/null || true
codesign --force --deep --sign - "$APP_PATH" 2>/dev/null || true

echo ""
echo "============================================="
echo "  DONE: ${ROOT}/${APP_PATH}"
echo "============================================="
echo ""
echo "  Install:  cp -R \"${APP_PATH}\" /Applications/"
echo "  Launch:   open \"${APP_PATH}\""
echo ""
