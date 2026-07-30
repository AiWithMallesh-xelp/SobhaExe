param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputMsi,

    [string]$LogFile
)

# Native tools report failures via exit code, so keep PowerShell from turning
# their stderr output into terminating errors before the code can be checked.
$ErrorActionPreference = "Continue"
if (Test-Path "variable:PSNativeCommandUseErrorActionPreference") {
    $PSNativeCommandUseErrorActionPreference = $false
}

$WixVersion = "5.0.2"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

if (-not $LogFile) {
    $LogFile = Join-Path $Root "release\msi-build.log"
}
$logDir = Split-Path -Parent $LogFile
if ($logDir -and !(Test-Path $logDir)) {
    New-Item -ItemType Directory -Force -Path $logDir | Out-Null
}
Set-Content -Path $LogFile -Value "MSI build log $(Get-Date -Format o)" -Encoding utf8

function Write-Log {
    param([string]$Message)
    Write-Host $Message
    Add-Content -Path $LogFile -Value $Message -Encoding utf8
}

function Invoke-Native {
    param(
        [string]$Label,
        [string]$FilePath,
        [string[]]$Arguments,
        [switch]$AllowFailure
    )

    Write-Log ""
    Write-Log "=== $Label ==="
    Write-Log "> $FilePath $($Arguments -join ' ')"

    $output = & $FilePath @Arguments 2>&1
    $code = $LASTEXITCODE
    foreach ($line in $output) {
        Write-Log "    $line"
    }
    Write-Log "exit code: $code"

    if ($code -ne 0 -and -not $AllowFailure) {
        throw "$Label failed with exit code $code"
    }
    return $code
}

if (!(Test-Path $SourceDir)) {
    Write-Log "ERROR: MSI source directory not found: $SourceDir"
    throw "MSI source directory not found: $SourceDir"
}

$SourceDirResolved = (Resolve-Path $SourceDir).Path
$fileCount = (Get-ChildItem $SourceDirResolved -Recurse -File).Count
Write-Log "Source directory: $SourceDirResolved ($fileCount files)"

# ── Install WiX as a global .NET tool ───────────────────────
$toolsDir = Join-Path $env:USERPROFILE ".dotnet\tools"
$installCode = Invoke-Native -Label "dotnet tool install wix $WixVersion" `
    -FilePath "dotnet" `
    -Arguments @("tool", "install", "--global", "wix", "--version", $WixVersion) `
    -AllowFailure

if ($installCode -ne 0) {
    Write-Log "Install returned $installCode (likely already installed); trying update."
    Invoke-Native -Label "dotnet tool update wix $WixVersion" `
        -FilePath "dotnet" `
        -Arguments @("tool", "update", "--global", "wix", "--version", $WixVersion) `
        -AllowFailure | Out-Null
}

if ($env:Path -notlike "*$toolsDir*") {
    $env:Path = "$env:Path;$toolsDir"
}

$wixExe = Join-Path $toolsDir "wix.exe"
if (!(Test-Path $wixExe)) {
    $resolved = Get-Command wix -ErrorAction SilentlyContinue
    if ($resolved) {
        $wixExe = $resolved.Source
    } else {
        Write-Log "ERROR: wix.exe not found in $toolsDir and not on PATH."
        Write-Log "PATH = $env:Path"
        throw "wix.exe not found after install"
    }
}
Write-Log "Using wix: $wixExe"

Invoke-Native -Label "wix --version" -FilePath $wixExe -Arguments @("--version") -AllowFailure | Out-Null

Invoke-Native -Label "wix extension add UI" `
    -FilePath $wixExe `
    -Arguments @("extension", "add", "-g", "WixToolset.UI.wixext/$WixVersion") `
    -AllowFailure | Out-Null

Invoke-Native -Label "wix extension list" `
    -FilePath $wixExe `
    -Arguments @("extension", "list", "-g") `
    -AllowFailure | Out-Null

# ── Build the MSI ───────────────────────────────────────────
$outputDir = Split-Path -Parent $OutputMsi
if ($outputDir -and !(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

Invoke-Native -Label "wix build" `
    -FilePath $wixExe `
    -Arguments @(
        "build",
        (Join-Path $Root "installer\sobha.wxs"),
        "-ext", "WixToolset.UI.wixext",
        "-arch", "x64",
        "-d", "SourceDir=$SourceDirResolved",
        "-o", $OutputMsi
    ) | Out-Null

if (!(Test-Path $OutputMsi)) {
    Write-Log "ERROR: MSI was not created: $OutputMsi"
    throw "MSI was not created: $OutputMsi"
}

$sizeMb = [math]::Round((Get-Item $OutputMsi).Length / 1MB, 1)
Write-Log "MSI build successful: $OutputMsi ($sizeMb MB)"
