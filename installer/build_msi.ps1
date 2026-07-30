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

Write-Host "  Installing WiX Toolset..." -ForegroundColor DarkCyan
dotnet tool install --global wix
if ($LASTEXITCODE -ne 0) {
    dotnet tool update --global wix
}

$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + `
            [System.Environment]::GetEnvironmentVariable("Path", "User")

wix extension add WixToolset.UI.wixext --force

$HarvestFile = Join-Path $Root "installer\harvested.wxs"
$SourceDirResolved = (Resolve-Path $SourceDir).Path
$outputDir = Split-Path -Parent $OutputMsi
if ($outputDir -and !(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

Write-Host "  Harvesting $SourceDirResolved ..." -ForegroundColor DarkCyan
heat dir $SourceDirResolved `
    -cg AppFiles `
    -dr INSTALLFOLDER `
    -ke -scom -sreg -sfrag -srd `
    -var var.SourceDir `
    -out $HarvestFile

Write-Host "  Building MSI -> $OutputMsi ..." -ForegroundColor DarkCyan
wix build "$Root\installer\sobha.wxs" $HarvestFile `
    -ext WixToolset.UI.wixext `
    -d "SourceDir=$SourceDirResolved" `
    -o $OutputMsi

if (!(Test-Path $OutputMsi)) {
    Write-Error "MSI was not created: $OutputMsi"
}

$sizeMb = [math]::Round((Get-Item $OutputMsi).Length / 1MB, 1)
Write-Host "  MSI build successful: $OutputMsi ($sizeMb MB)" -ForegroundColor Green
