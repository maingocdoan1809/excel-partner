<#
.SYNOPSIS
  Install "Excel Partner" from scratch.

.DESCRIPTION
  Entry point you can run from anywhere, even before cloning. It:
    1. reuses the current checkout if this file sits inside one, otherwise
       clones the repo (with gh or git) into -Dir;
    2. hands over to scripts\install.ps1, which builds, trusts the localhost
       certificate, starts the bridge and sideloads the task pane into Excel.

  WSL is optional: the bridge and provider CLIs run natively on Windows or in
  WSL2, chosen by the installer (or forced with -Mode).

.EXAMPLE
  # already cloned
  powershell -ExecutionPolicy Bypass -File install.ps1

.EXAMPLE
  # straight from scratch (no clone yet)
  powershell -ExecutionPolicy Bypass -File install.ps1 -Providers copilot,codex -Open

.EXAMPLE
  # from a one-liner when the file is downloaded and gh is authenticated
  powershell -ExecutionPolicy Bypass -File .\install.ps1 -Dir C:\tools\excel-partner
#>
[CmdletBinding()]
param(
  [string]$Repo = '',
  [string]$Dir = "$env:USERPROFILE\excel-partner",
  [ValidateSet('auto', 'native', 'wsl')] [string]$Mode = 'auto',
  [string]$Providers = '',
  [string]$ReferenceDir = '',
  [switch]$Open,
  [switch]$SkipProviders
)

$ErrorActionPreference = 'Stop'

function Write-Info2($m) { Write-Host "    $m" }
function Fail2($m) { Write-Host "`nERROR: $m" -ForegroundColor Red; exit 1 }

$insideRepo = Test-Path (Join-Path $PSScriptRoot 'package.json') -and `
              Test-Path (Join-Path $PSScriptRoot 'scripts\install.ps1')

if ($insideRepo) {
  $RepoRoot = $PSScriptRoot
}
else {
  if ([string]::IsNullOrWhiteSpace($Repo) -and -not (Test-Path (Join-Path $Dir '.git'))) {
    Fail2 'Pass -Repo <repository-url> when installing outside an existing checkout.'
  }
  Write-Host "Excel Partner - bootstrap" -ForegroundColor Green
  Write-Info2 "cloning $Repo"
  Write-Info2 "into     $Dir"

  if (Test-Path (Join-Path $Dir '.git')) {
    Write-Info2 "using existing checkout at $Dir"
    Push-Location $Dir
    git pull --ff-only
    Pop-Location
  }
  elseif ((Test-Path $Dir) -and ((Get-ChildItem -Force $Dir | Measure-Object).Count -gt 0)) {
    Fail2 "$Dir already exists and is not a git checkout. Pass -Dir with an empty/absent path."
  }
  elseif (Get-Command gh -ErrorAction SilentlyContinue) {
    $slug = ($Repo -replace '\.git$', '') -replace '.*github\.com/', ''
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Dir) | Out-Null
    gh repo clone $slug $Dir
    if ($LASTEXITCODE -ne 0) { Fail2 "gh repo clone failed (are you logged in? run 'gh auth login')." }
  }
  elseif (Get-Command git -ErrorAction SilentlyContinue) {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Dir) | Out-Null
    git clone $Repo $Dir
    if ($LASTEXITCODE -ne 0) { Fail2 "git clone failed (a private repo needs credentials; install gh and run 'gh auth login')." }
  }
  else {
    Fail2 "Neither gh nor git was found. Install Git for Windows (https://git-scm.com) or GitHub CLI."
  }

  $RepoRoot = $Dir
}

$installer = Join-Path $RepoRoot 'scripts\install.ps1'
if (-not (Test-Path $installer)) { Fail2 "installer not found at $installer" }

$params = @{ Mode = $Mode }
if (-not [string]::IsNullOrWhiteSpace($Providers)) { $params.Providers = $Providers }
if (-not [string]::IsNullOrWhiteSpace($ReferenceDir)) { $params.ReferenceDir = $ReferenceDir }
if ($Open) { $params.Open = $true }
if ($SkipProviders) { $params.SkipProviders = $true }

& $installer @params
