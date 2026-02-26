# ============================================================
#  build_windows.ps1  –  One-click portable Windows build
#  Run from the project folder:  .\build_windows.ps1
# ============================================================
$ErrorActionPreference = "Stop"

$ENTRY   = "sales_receipt_generation.py"
$APPNAME = "sobha"

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Building $APPNAME.exe from $ENTRY"         -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Ensure we run from the script's own directory
Set-Location -Path $PSScriptRoot

# ── 1. Create venv if missing ────────────────────────────────
if (!(Test-Path ".venv")) {
    Write-Host "[1/6] Creating virtual environment..." -ForegroundColor Yellow
    py -m venv .venv
} else {
    Write-Host "[1/6] Virtual environment already exists." -ForegroundColor Green
}

# ── 2. Activate venv ────────────────────────────────────────
Write-Host "[2/6] Activating venv..." -ForegroundColor Yellow
& ".\.venv\Scripts\Activate.ps1"

# ── 3. Install Python dependencies ──────────────────────────
Write-Host "[3/6] Installing dependencies..." -ForegroundColor Yellow
python -m pip install --upgrade pip --quiet

if (Test-Path "requirements.txt") {
    pip install -r requirements.txt --quiet
} else {
    Write-Host "  requirements.txt not found – installing playwright only..." -ForegroundColor DarkYellow
    pip install playwright==1.58.0 --quiet
}

pip install pyinstaller --quiet
Write-Host "  Dependencies installed." -ForegroundColor Green

# ── 4. Download Chromium into portable pw-browsers/ ─────────
Write-Host "[4/6] Downloading Playwright Chromium (portable)..." -ForegroundColor Yellow
$env:PLAYWRIGHT_BROWSERS_PATH = "$PSScriptRoot\pw-browsers"
python -m playwright install chromium
Write-Host "  Chromium ready in: $PSScriptRoot\pw-browsers" -ForegroundColor Green

# ── 5. Build exe with PyInstaller ───────────────────────────
Write-Host "[5/6] Building $APPNAME.exe with PyInstaller..." -ForegroundColor Yellow

$pyinstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--name", $APPNAME,
    "--collect-all", "playwright"
)

# Include forest-light.tcl theme if it exists
if (Test-Path "forest-light.tcl") {
    $pyinstallerArgs += "--add-data"
    $pyinstallerArgs += "forest-light.tcl;."
    Write-Host "  Including forest-light.tcl theme." -ForegroundColor DarkCyan
}

$pyinstallerArgs += $ENTRY

& pyinstaller @pyinstallerArgs

if (!(Test-Path "dist\$APPNAME.exe")) {
    Write-Host "ERROR: dist\$APPNAME.exe was not created. Check PyInstaller output above." -ForegroundColor Red
    exit 1
}
Write-Host "  Build successful: dist\$APPNAME.exe" -ForegroundColor Green

# ── 6. Assemble release/ folder ─────────────────────────────
Write-Host "[6/6] Assembling release folder..." -ForegroundColor Yellow

if (Test-Path "release") {
    Remove-Item -Recurse -Force "release"
}
New-Item -ItemType Directory -Path "release" | Out-Null

# Core files
Copy-Item ".\dist\$APPNAME.exe"  ".\release\$APPNAME.exe"  -Force
Copy-Item ".\config.json"        ".\release\config.json"    -Force

# Optional: pre-bundled auth session (client starts logged in)
if (Test-Path "auth.json") {
    Copy-Item ".\auth.json" ".\release\auth.json" -Force
    Write-Host "  auth.json included (client starts logged in)." -ForegroundColor DarkCyan
}

# Playwright browser binaries (required for no-setup run)
if (Test-Path "pw-browsers") {
    Copy-Item ".\pw-browsers" ".\release\pw-browsers" -Recurse -Force
} else {
    Write-Host "ERROR: pw-browsers/ folder missing – Playwright install step failed." -ForegroundColor Red
    exit 1
}

# Client readme
if (Test-Path "README_CLIENT.md") {
    Copy-Item ".\README_CLIENT.md" ".\release\README_CLIENT.md" -Force
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  DONE! Release folder ready:" -ForegroundColor Green
Write-Host "  $PSScriptRoot\release\" -ForegroundColor White
Write-Host ""
Write-Host "  Contents sent to client:" -ForegroundColor Yellow
Get-ChildItem ".\release" | Format-Table Name, Length -AutoSize
Write-Host ""
Write-Host "  Zip the 'release' folder and send to client." -ForegroundColor Yellow
Write-Host "  Client: double-click auto.exe  ✅"           -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
