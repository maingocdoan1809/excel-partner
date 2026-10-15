<#
.SYNOPSIS
  Sideload the Excel Partner add-in into Excel desktop on Windows.

.DESCRIPTION
  Copies manifest.xml into a shared folder and registers that folder as a
  trusted Office add-in catalog, so the add-in shows up under
  Insert > My Add-ins > Shared Folder without editing the manifest by hand.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File sideload.ps1 `
    -ManifestPath 'C:\path\to\excel-partner\addin\manifest.xml'
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$ManifestPath,

  [string]$ShareDir = "$env:USERPROFILE\ExcelAIAddin"
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ManifestPath)) {
  throw "Manifest not found: $ManifestPath"
}

if (-not (Test-Path -LiteralPath $ShareDir)) {
  New-Item -ItemType Directory -Path $ShareDir | Out-Null
  Write-Host "Created shared folder: $ShareDir"
}

Copy-Item -LiteralPath $ManifestPath -Destination (Join-Path $ShareDir 'manifest.xml') -Force
Write-Host "Copied manifest to $ShareDir\manifest.xml"

$catalogGuid = '{E5B1A3C7-2F4D-4A9B-9C1E-8D7F6A5B4C3D}'
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogGuid"

if (-not (Test-Path $catalogKey)) {
  New-Item -Path $catalogKey -Force | Out-Null
}
New-ItemProperty -Path $catalogKey -Name 'Id' -Value $catalogGuid -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name 'Url' -Value $ShareDir -PropertyType String -Force | Out-Null
New-ItemProperty -Path $catalogKey -Name 'Flags' -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "Registered trusted add-in catalog at $catalogKey"

# Developer registration: shows the add-in under Insert > My Add-ins > Developer Add-ins
# even where Office refuses a local-path shared-folder catalog.
$manifestId = ([xml](Get-Content -LiteralPath $ManifestPath -Raw)).OfficeApp.Id
$devKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
if (-not (Test-Path $devKey)) { New-Item -Path $devKey -Force | Out-Null }
New-ItemProperty -Path $devKey -Name $manifestId -Value (Join-Path $ShareDir 'manifest.xml') -PropertyType String -Force | Out-Null
Write-Host "Registered developer add-in $manifestId"

Write-Host ''
Write-Host 'Next steps:'
Write-Host '  1. Make sure the bridge is running in WSL: bash scripts/start.sh'
Write-Host '  2. Trust the localhost certificate in certs/localhost.crt (docs/install.md)'
Write-Host '  3. Restart Excel, then click Home > Excel Partner'
Write-Host '     (or Insert > My Add-ins > Developer Add-ins > Excel Partner)'
