<#
.SYNOPSIS
  Update an existing "Excel Partner" (excel-partner) install.

.DESCRIPTION
  Pulls the latest source, reinstalls dependencies, rebuilds, re-sideloads the
  manifest (in case it changed) and restarts the bridge. Works for a native
  Windows install or a WSL2 install; the mode is auto-detected like the installer.
  It never installs provider CLIs or re-trusts the certificate.

.PARAMETER Mode
  auto (default) | native | wsl.

.PARAMETER Open
  Start Excel when the update finishes.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\update.ps1

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\update.ps1 -Mode wsl
#>
[CmdletBinding()]
param(
  [ValidateSet('auto', 'native', 'wsl')] [string]$Mode = 'auto',
  [switch]$Open
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
. (Join-Path $PSScriptRoot 'windows-common.ps1')

Write-Host "Excel Partner - update" -ForegroundColor Green
Write-Info "repo: $Root"

$Mode = Resolve-Mode $Mode $Root

# ------------------------------------------------------------------- WSL2 ----
if ($Mode -eq 'wsl') {
  $wslRoot = Get-WslRoot $Root
  $cmd = "cd '$wslRoot' && bash scripts/update.sh"
  if ($Open) { $cmd += " --open" }
  Write-Step "running the WSL update"
  Write-Info $cmd
  wsl.exe bash -lc $cmd
  if ($LASTEXITCODE -ne 0) { Fail "WSL update failed." }
  Write-Host "`nUpdate complete (WSL mode)." -ForegroundColor Green
  exit 0
}

# ------------------------------------------------------------- native Windows -
Write-Step "checking Node.js"
Assert-Node

if (Test-Path (Join-Path $Root '.git')) {
  Write-Step "pulling latest changes"
  git pull --ff-only
  if ($LASTEXITCODE -ne 0) { Write-Info "git pull failed; continuing with the current checkout" }
}

Write-Step "installing dependencies"
if (Test-Path (Join-Path $Root 'package-lock.json')) { npm ci } else { npm install }
if ($LASTEXITCODE -ne 0) { Fail "dependency install failed." }

Write-Step "rebuilding"
npm run build
if ($LASTEXITCODE -ne 0) { Fail "build failed." }

Write-Step "localhost certificate"
Ensure-Cert $Root
Trust-Cert $Root

Write-Step "restarting the bridge"
Stop-Bridge $Root
Start-Bridge $Root

Write-Step "re-sideloading into Excel"
Invoke-Sideload $Root

if ($Open) {
  Write-Step "starting Excel"
  Start-Process 'C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE' -ArgumentList '/x' -ErrorAction SilentlyContinue
}

Write-Host "`nUpdate complete." -ForegroundColor Green
