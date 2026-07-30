# ============================================================
#  build_windows.ps1  –  Portable EXE + MSI Windows build
#  Run from the project folder:  .\build_windows.ps1
# ============================================================
$ErrorActionPreference = "Stop"

$ENTRY      = "sales_receipt_generation.py"
$APPNAME    = "sobha"
$RELEASEDIR = "release\sobha-app"

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Building $APPNAME (onedir) from $ENTRY"    -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

Set-Location -Path $PSScriptRoot

# ── 1. Create venv if missing ────────────────────────────────
if (!(Test-Path ".venv")) {
    Write-Host "[1/8] Creating virtual environment..." -ForegroundColor Yellow
    py -m venv .venv
} else {
    Write-Host "[1/8] Virtual environment already exists." -ForegroundColor Green
}

# ── 2. Activate venv ────────────────────────────────────────
Write-Host "[2/8] Activating venv..." -ForegroundColor Yellow
& ".\.venv\Scripts\Activate.ps1"

# ── 3. Install Python dependencies ──────────────────────────
Write-Host "[3/8] Installing dependencies..." -ForegroundColor Yellow
python -m pip install --upgrade pip --quiet

if (Test-Path "requirements.txt") {
    pip install -r requirements.txt --quiet
} else {
    Write-Host "  requirements.txt not found – installing playwright only..." -ForegroundColor DarkYellow
    pip install playwright==1.58.0 --quiet
}

pip install pyinstaller --quiet
Write-Host "  Dependencies installed." -ForegroundColor Green

# ── 4. Optional app icon ────────────────────────────────────
Write-Host "[4/8] Preparing icon..." -ForegroundColor Yellow
if (Test-Path "scripts\generate_icon.py") {
    pip install pillow --quiet
    python scripts\generate_icon.py
}

# ── 5. Download Chromium into portable pw-browsers/ ─────────
Write-Host "[5/8] Downloading Playwright Chromium (portable)..." -ForegroundColor Yellow
$env:PLAYWRIGHT_BROWSERS_PATH = "$PSScriptRoot\pw-browsers"
python -m playwright install chromium
Write-Host "  Chromium ready in: $PSScriptRoot\pw-browsers" -ForegroundColor Green

# ── 6. Build exe with PyInstaller (onedir) ──────────────────
Write-Host "[6/8] Building $APPNAME.exe with PyInstaller (onedir)..." -ForegroundColor Yellow

$pyinstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--onedir",
    "--windowed",
    "--name", $APPNAME,
    "--collect-all", "playwright"
)

if (Test-Path "installer\sobha.ico") {
    $pyinstallerArgs += "--icon"
    $pyinstallerArgs += "installer\sobha.ico"
    Write-Host "  Using installer\sobha.ico" -ForegroundColor DarkCyan
}

if (Test-Path "forest-light.tcl") {
    $pyinstallerArgs += "--add-data"
    $pyinstallerArgs += "forest-light.tcl;."
    Write-Host "  Including forest-light.tcl theme." -ForegroundColor DarkCyan
}

$pyinstallerArgs += $ENTRY

& pyinstaller @pyinstallerArgs

if (!(Test-Path "dist\$APPNAME\$APPNAME.exe")) {
    Write-Host "ERROR: dist\$APPNAME\$APPNAME.exe was not created. Check PyInstaller output above." -ForegroundColor Red
    exit 1
}
Write-Host "  Build successful: dist\$APPNAME\$APPNAME.exe" -ForegroundColor Green

# ── 7. Assemble release/sobha-app/ folder ───────────────────
Write-Host "[7/8] Assembling $RELEASEDIR folder..." -ForegroundColor Yellow

if (Test-Path "release") {
    Remove-Item -Recurse -Force "release"
}
New-Item -ItemType Directory -Path $RELEASEDIR -Force | Out-Null

Copy-Item ".\dist\$APPNAME\*" $RELEASEDIR -Recurse -Force
Copy-Item ".\config.json" "$RELEASEDIR\config.json" -Force

if (Test-Path "sobha_logo_brand.png") {
    Copy-Item ".\sobha_logo_brand.png" "$RELEASEDIR\sobha_logo_brand.png" -Force
    Write-Host "  sobha_logo_brand.png included." -ForegroundColor DarkCyan
} else {
    Write-Host "WARNING: sobha_logo_brand.png not found – app header logo will be missing." -ForegroundColor DarkYellow
}

if (Test-Path "auth.json") {
    Copy-Item ".\auth.json" "$RELEASEDIR\auth.json" -Force
    Write-Host "  auth.json included (client starts logged in)." -ForegroundColor DarkCyan
}

if (Test-Path "pw-browsers") {
    Copy-Item ".\pw-browsers" "$RELEASEDIR\pw-browsers" -Recurse -Force
} else {
    Write-Host "ERROR: pw-browsers/ folder missing – Playwright install step failed." -ForegroundColor Red
    exit 1
}

if (Test-Path "README_CLIENT.md") {
    Copy-Item ".\README_CLIENT.md" "$RELEASEDIR\README_CLIENT.md" -Force
}

Write-Host "[8/8] Creating ZIP packages..." -ForegroundColor Yellow
$appZipItems = @(
    "$RELEASEDIR\$APPNAME.exe",
    "$RELEASEDIR\_internal",
    "$RELEASEDIR\config.json"
)
if (Test-Path "$RELEASEDIR\README_CLIENT.md") {
    $appZipItems += "$RELEASEDIR\README_CLIENT.md"
}
if (Test-Path "$RELEASEDIR\sobha_logo_brand.png") {
    $appZipItems += "$RELEASEDIR\sobha_logo_brand.png"
}

$portableZipItems = @(
    "$RELEASEDIR\$APPNAME.exe",
    "$RELEASEDIR\_internal",
    "$RELEASEDIR\config.json",
    "$RELEASEDIR\pw-browsers"
)
if (Test-Path "$RELEASEDIR\README_CLIENT.md") {
    $portableZipItems += "$RELEASEDIR\README_CLIENT.md"
}
if (Test-Path "$RELEASEDIR\sobha_logo_brand.png") {
    $portableZipItems += "$RELEASEDIR\sobha_logo_brand.png"
}

Compress-Archive -Path $appZipItems -DestinationPath ".\release\sobha-app-only.zip" -Force
Compress-Archive -Path "$RELEASEDIR\pw-browsers" `
    -DestinationPath ".\release\sobha-pw-browsers.zip" -Force
Compress-Archive -Path $portableZipItems `
    -DestinationPath ".\release\sobha-windows-portable.zip" -Force

Write-Host ""
Write-Host "  Building MSI installer..." -ForegroundColor Yellow
& "$PSScriptRoot\installer\build_msi.ps1" -SourceDir "$PSScriptRoot\$RELEASEDIR" -OutputMsi "$PSScriptRoot\release\sobha-setup.msi"

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  DONE! Release folder ready:" -ForegroundColor Green
Write-Host "  $PSScriptRoot\release\" -ForegroundColor White
Write-Host ""
Write-Host "  Packages:" -ForegroundColor Yellow
Write-Host "    release\sobha-windows-portable.zip  (portable test package)" -ForegroundColor White
Write-Host "    release\sobha-setup.msi             (Windows installer)" -ForegroundColor White
Write-Host "    release\sobha-app-only.zip" -ForegroundColor White
Write-Host "    release\sobha-pw-browsers.zip" -ForegroundColor White
Write-Host ""
Get-ChildItem ".\release" | Format-Table Name, Length -AutoSize
Write-Host "=============================================" -ForegroundColor Cyan
