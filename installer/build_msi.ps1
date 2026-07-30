param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputMsi
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

if (!(Test-Path $SourceDir)) {
    Write-Error "MSI source directory not found: $SourceDir"
}

if (!(Test-Path "$Root\installer\sobha.wxs")) {
    Write-Error "WiX source not found: installer\sobha.wxs"
}

Write-Host "  Installing WiX Toolset..." -ForegroundColor DarkCyan
dotnet tool install --global wix --version 5.0.2 2>$null
if ($LASTEXITCODE -ne 0) {
    dotnet tool update --global wix --version 5.0.2
}

$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + `
            [System.Environment]::GetEnvironmentVariable("Path", "User")

wix extension add WixToolset.UI.wixext --force
wix extension add WixToolset.Heat.wixext --force

$HarvestFile = Join-Path $Root "installer\harvested.wxs"
$SourceDirResolved = (Resolve-Path $SourceDir).Path

Write-Host "  Harvesting files from $SourceDirResolved ..." -ForegroundColor DarkCyan
wix heat dir $SourceDirResolved `
    -cg AppFiles `
    -dr INSTALLFOLDER `
    -ke -scom -sreg -sfrag -srd `
    -var var.SourceDir `
    -out $HarvestFile

Write-Host "  Building MSI -> $OutputMsi ..." -ForegroundColor DarkCyan
$outputDir = Split-Path -Parent $OutputMsi
if ($outputDir -and !(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

wix build "$Root\installer\sobha.wxs" $HarvestFile `
    -ext WixToolset.UI.wixext `
    -d "SourceDir=$SourceDirResolved" `
    -o $OutputMsi

if (!(Test-Path $OutputMsi)) {
    Write-Error "MSI was not created: $OutputMsi"
}

$sizeMb = [math]::Round((Get-Item $OutputMsi).Length / 1MB, 1)
Write-Host "  MSI build successful: $OutputMsi ($sizeMb MB)" -ForegroundColor Green
