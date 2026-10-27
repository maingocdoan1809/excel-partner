<#
.SYNOPSIS
  One-shot installer for "Excel Partner" (excel-partner).

.DESCRIPTION
  Sets up the add-in on a Windows machine. WSL is optional: the bridge and the
  provider CLIs can run natively on Windows or inside WSL2. The installer asks
  which provider CLI(s) to install (default: copilot) unless -Providers is given,
  then builds, trusts the localhost certificate, starts the bridge and sideloads
  the task pane into Excel desktop.

.PARAMETER Mode
  auto (default) | native | wsl. In auto it prefers a native Windows install when
  node.exe is present, otherwise it falls back to WSL2.

.PARAMETER Providers
  Comma-separated subset of: copilot, claude, codex, opencode. Default: copilot.
  Use 'none' to skip provider installation (e.g. CLIs are already installed elsewhere).

.PARAMETER ReferenceDir
  Optional folder the assistant may read .xlsx templates from. Written to
  excel-partner.config.json (read-only access, only that folder).

.PARAMETER AllowedHost
  Extra hostname or IP the bridge should answer to (for example a WSL IP when
  BRIDGE_HOST=0.0.0.0). Requests whose Host header is not localhost, 127.0.0.1,
  ::1 or listed here are refused. Written to excel-partner.config.json.

.PARAMETER Open
  Start Excel when the install finishes.

.PARAMETER SkipProviders
  Do not install or ask about provider CLIs.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\install.ps1

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\install.ps1 -Mode wsl -Providers copilot,codex -Open
#>
[CmdletBinding()]
param(
  [ValidateSet('auto', 'native', 'wsl')] [string]$Mode = 'auto',
  [string]$Providers = '',
  [string]$ReferenceDir = '',
  [string[]]$AllowedHost = @(),
  [switch]$Open,
  [switch]$SkipProviders
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
. (Join-Path $PSScriptRoot 'windows-common.ps1')

Write-Host "Excel Partner - installer" -ForegroundColor Green
Write-Info "repo: $Root"

if (-not $SkipProviders -and [string]::IsNullOrWhiteSpace($Providers)) {
  Write-Host ""
  Write-Host "Which provider CLI(s) should this machine install?" -ForegroundColor Yellow
  Write-Info "options: copilot, claude, codex, opencode   (comma separated, or 'none')"
  $answer = Read-Host "Providers [copilot]"
  if ([string]::IsNullOrWhiteSpace($answer)) { $answer = 'copilot' }
  $Providers = $answer
}

$Mode = Resolve-Mode $Mode $Root

# ------------------------------------------------------------------- WSL2 ----
if ($Mode -eq 'wsl') {
  $wslRoot = Get-WslRoot $Root
  $parts = @("cd '$wslRoot'", "bash scripts/install-windows.sh")
  if (-not $SkipProviders -and -not [string]::IsNullOrWhiteSpace($Providers)) { $parts += "--providers '$Providers'" }
  if (-not [string]::IsNullOrWhiteSpace($ReferenceDir)) { $parts += "--reference-dir '$ReferenceDir'" }
  foreach ($h in $AllowedHost) { if (-not [string]::IsNullOrWhiteSpace($h)) { $parts += "--allowed-host '$h'" } }
  if ($Open) { $parts += "--open" }
  $cmd = $parts -join ' '

  Write-Step "running the WSL install"
  Write-Info $cmd
  wsl.exe bash -lc $cmd
  if ($LASTEXITCODE -ne 0) { Fail "WSL install failed." }
  Write-Host "`nDone (WSL mode). In Excel: Home > Excel Partner." -ForegroundColor Green
  exit 0
}

# ------------------------------------------------------------- native Windows -
Write-Step "checking Node.js"
Assert-Node

if (-not $SkipProviders) {
  Write-Step "provider CLIs"
  Install-Providers $Providers
}

Write-Step "installing dependencies"
if (Test-Path (Join-Path $Root 'package-lock.json')) { npm ci } else { npm install }
if ($LASTEXITCODE -ne 0) { Fail "dependency install failed." }

Write-Step "building bridge, MCP server and task pane"
npm run build
if ($LASTEXITCODE -ne 0) { Fail "build failed." }

Write-Step "localhost certificate"
Ensure-Cert $Root
Trust-Cert $Root

Write-Step "writing local config"
Write-Config $Root $ReferenceDir $AllowedHost

Write-Step "starting the bridge"
Stop-Bridge $Root
Start-Bridge $Root

Write-Step "sideloading into Excel"
Invoke-Sideload $Root

if ($Open) {
  Write-Step "starting Excel"
  Start-Process 'C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE' -ArgumentList '/x' -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "Done. In Excel: Home > Excel Partner." -ForegroundColor Green
Write-Host "Restart any Excel window that was already open before this install."
Write-Host ""
Write-Host "The bridge is running now but does not auto-start on reboot."
Write-Host "To update later:  powershell -File scripts\update.ps1"
