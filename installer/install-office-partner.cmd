@echo off
rem Office Partner - all-in-one installer for the current Windows user. No administrator rights needed.
rem One file: the batch lines below start PowerShell, which runs everything after the marker line.
rem Switches: /quiet (no prompts, no pause; for Intune/SCCM in user context), /version x.y.z (install that release),
rem           /uninstall (remove Office Partner and the portable Node.js this installer added).
rem Environment overrides: OP_QUIET=1, OP_VERSION, OP_PROVIDERS (default copilot), OP_NODE_VERSION (exact Node.js 22+ version),
rem           OP_REGISTRY (private npm registry), OP_NO_SIDELOAD=1 (manifest deployed from Microsoft 365), OP_PURGE=1 (with /uninstall: also delete chats, logs and provider sign-ins).
rem Exit codes: 0 ok, 10 Node.js, 20 npm install, 30 setup, 40 certificate trust, 50 Office is running (quiet mode),
rem           60 unsupported context (SYSTEM account), 1 anything else.
setlocal
:args
if "%~1"=="" goto run
if /i "%~1"=="/quiet" set "OP_QUIET=1"
if /i "%~1"=="/uninstall" set "OP_UNINSTALL=1"
if /i "%~1"=="/version" (set "OP_VERSION=%~2" & shift)
shift
goto args
:run
set "OP_SELF=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:OP_SELF); $i=$t.LastIndexOf('#'+'#PS-BEGIN##'); Invoke-Expression $t.Substring($i)"
set "code=%errorlevel%"
if not defined OP_QUIET (echo. & pause)
exit /b %code%
##PS-BEGIN##
# Office Partner installer for Windows (PowerShell part). Runs as the signed-in user: no administrator rights needed.
$Quiet = $env:OP_QUIET -eq '1'
$Uninstall = $env:OP_UNINSTALL -eq '1'
$Version = $env:OP_VERSION
$Providers = if ($env:OP_PROVIDERS) { $env:OP_PROVIDERS } else { 'copilot' }
# An exact Node.js release (checked against the SHA256 list nodejs.org publishes) so every machine gets the same one.
$NodeVersion = if ($env:OP_NODE_VERSION) { $env:OP_NODE_VERSION.TrimStart('v') } else { '22.23.5' }
$Registry = $env:OP_REGISTRY
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
if ($ExecutionContext.SessionState.LanguageMode -eq 'FullLanguage') {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  # Corporate proxies that need the Windows sign-in.
  [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
}

function Done([int] $code) { if ($script:logging) { Stop-Transcript | Out-Null }; exit $code }
function Step([string] $message) { Write-Host ""; Write-Host "==> $message" -ForegroundColor Cyan }
function Fail([string] $message, [int] $code = 1) {
  Write-Host ""; Write-Host "ERROR: $message" -ForegroundColor Red
  if ($script:logging) { Write-Host "Install log: $log" }
  Done $code
}

# Each Windows user has their own add-in, certificate and bridge: the SYSTEM account has no Office profile to install into.
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if ($identity.IsSystem -or -not $env:LOCALAPPDATA) {
  Write-Host 'ERROR: Run this installer as the signed-in user (Intune: Install behavior = User), not as SYSTEM.' -ForegroundColor Red
  exit 60
}

$dataDir = Join-Path $env:LOCALAPPDATA 'Excel Partner'
$nodeHome = Join-Path $env:LOCALAPPDATA 'Programs\nodejs'
New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
$log = Join-Path $dataDir 'install.log'
Start-Transcript -Path $log -Force | Out-Null
$script:logging = $true

# Node.js supported by the package: >=22.12.0 <26.0.0
function Test-NodeVersion([string] $value) {
  if ($value -notmatch '^v?(\d+)\.(\d+)\.(\d+)') { return $false }
  $major = [int]$Matches[1]; $minor = [int]$Matches[2]
  return ($major -eq 22 -and $minor -ge 12) -or ($major -gt 22 -and $major -lt 26)
}

function Get-InstalledNode {
  $candidates = @((Get-Command node.exe -ErrorAction SilentlyContinue).Source, (Join-Path $nodeHome 'node.exe')) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
  foreach ($candidate in $candidates) {
    $found = (& $candidate --version 2>$null)
    if ($LASTEXITCODE -eq 0 -and (Test-NodeVersion $found) -and (Test-Path -LiteralPath (Join-Path (Split-Path $candidate) 'npm.cmd'))) { return $candidate }
  }
  return $null
}

function Install-PortableNode {
  Step "Downloading Node.js $NodeVersion for this user only (no administrator rights)"
  if (-not (Test-NodeVersion $NodeVersion)) { Fail "OP_NODE_VERSION $NodeVersion is not supported (22.12 or later, before 26)." 10 }
  $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
  $name = "node-v$NodeVersion-win-$arch"
  $work = Join-Path $env:TEMP "office-partner-node-$([guid]::NewGuid().ToString('N'))"
  New-Item -ItemType Directory -Path $work -Force | Out-Null
  try {
    $zip = Join-Path $work "$name.zip"
    try {
      Invoke-WebRequest -Uri "https://nodejs.org/dist/v$NodeVersion/$name.zip" -OutFile $zip -UseBasicParsing
      $sums = (Invoke-WebRequest -Uri "https://nodejs.org/dist/v$NodeVersion/SHASUMS256.txt" -UseBasicParsing).Content
    } catch { Fail "Could not download Node.js $NodeVersion from nodejs.org: $($_.Exception.Message). Check the network or proxy." 10 }
    $expected = ($sums -split "`n" | Where-Object { $_ -match "\s$([regex]::Escape("$name.zip"))\s*$" } | Select-Object -First 1) -replace '\s.*$', ''
    $actual = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
    if (-not $expected -or $expected.ToLower() -ne $actual.ToLower()) { Fail 'The downloaded Node.js does not match its SHA256 checksum. Nothing was installed.' 10 }
    Expand-Archive -LiteralPath $zip -DestinationPath $work -Force
    if (Test-Path -LiteralPath $nodeHome) { Remove-Item -LiteralPath $nodeHome -Recurse -Force }
    New-Item -ItemType Directory -Path (Split-Path $nodeHome) -Force | Out-Null
    Move-Item -LiteralPath (Join-Path $work $name) -Destination $nodeHome
    Set-Content -LiteralPath (Join-Path $nodeHome '.installed-by-office-partner') -Value $NodeVersion
  } finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}

function Add-UserPath([string] $directory) {
  $current = [Environment]::GetEnvironmentVariable('Path', 'User')
  $parts = @($current -split ';' | Where-Object { $_ })
  if ($parts -notcontains $directory) { [Environment]::SetEnvironmentVariable('Path', (@($directory) + $parts) -join ';', 'User') }
  if (($env:Path -split ';') -notcontains $directory) { $env:Path = "$directory;$env:Path" }
}

function Remove-UserPath([string] $directory) {
  $parts = @([Environment]::GetEnvironmentVariable('Path', 'User') -split ';' | Where-Object { $_ -and $_ -ne $directory })
  [Environment]::SetEnvironmentVariable('Path', $parts -join ';', 'User')
}

function Wait-OfficeClosed {
  Step 'Checking Word, Excel and PowerPoint'
  $office = Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue
  if (-not $office) { return }
  if ($Quiet) { Fail "Close Word, Excel and PowerPoint first (open: $(($office.ProcessName | Sort-Object -Unique) -join ', '))." 50 }
  Write-Host 'Open: ' ($office.ProcessName | Sort-Object -Unique) -ForegroundColor Yellow
  Write-Host 'Save your work and close Word, Excel and PowerPoint.'
  Read-Host 'Press Enter when they are closed'
  if (Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue) { Fail 'Word, Excel or PowerPoint is still open. Close them and run the installer again.' 50 }
}

if ($Uninstall) {
  Write-Host 'Office Partner - uninstall' -ForegroundColor Green
  Wait-OfficeClosed
  $node = Get-InstalledNode
  if ($node) {
    $npm = Join-Path (Split-Path $node) 'npm.cmd'
    $command = Join-Path ((& $npm prefix -g).Trim()) 'office-partner.cmd'
    if (Test-Path -LiteralPath $command) {
      Step 'Removing the add-in registration, startup entry and certificate trust'
      & $command uninstall
      if ($LASTEXITCODE -ne 0) { Fail "office-partner uninstall failed (exit $LASTEXITCODE)." 30 }
    }
    & $npm uninstall -g office-partner | Out-Null
  }
  if (Test-Path -LiteralPath (Join-Path $nodeHome '.installed-by-office-partner')) {
    Step 'Removing the portable Node.js this installer added'
    Remove-UserPath $nodeHome
    Remove-Item -LiteralPath $nodeHome -Recurse -Force -ErrorAction SilentlyContinue
  }
  if ($env:OP_PURGE -eq '1') {
    Step 'Deleting chats, logs and provider sign-ins'
    Stop-Transcript | Out-Null; $script:logging = $false
    Remove-Item -LiteralPath $dataDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  Write-Host ''; Write-Host 'Office Partner was removed.' -ForegroundColor Green
  Done 0
}

Write-Host 'Office Partner - install' -ForegroundColor Green
Wait-OfficeClosed

Step 'Checking Node.js'
$node = Get-InstalledNode
if ($node) { Write-Host "Using the installed Node.js: $node ($(& $node --version))" } else { Install-PortableNode; $node = Join-Path $nodeHome 'node.exe' }
$nodeDir = Split-Path $node
Add-UserPath $nodeDir
$npm = Join-Path $nodeDir 'npm.cmd'
Write-Host "Node.js $(& $node --version), npm $(& $npm --version)"

if (-not $Version) { Write-Host 'No version given (/version or OP_VERSION): installing the latest release. Pin a version for managed rollouts.' -ForegroundColor Yellow }
$spec = if ($Version) { "office-partner@$Version" } else { 'office-partner@latest' }
$registryArgs = if ($Registry) { @('--registry', $Registry) } else { @() }
$existing = Test-Path -LiteralPath (Join-Path $dataDir 'deployment.json')

Step "Installing $spec"
& $npm install -g $spec @registryArgs
if ($LASTEXITCODE -ne 0) { Fail "npm install failed (exit $LASTEXITCODE). Check the network or the company proxy (npm config set proxy ...)." 20 }
$prefix = (& $npm prefix -g).Trim()
Add-UserPath $prefix
$command = Join-Path $prefix 'office-partner.cmd'
if (-not (Test-Path -LiteralPath $command)) { Fail "The office-partner command was not found in $prefix." 20 }

if ($existing) {
  # Setup keeps the release that is already active; an update switches to the one just installed (and rolls back if it fails).
  Step "Switching the active release to $spec"
  & $command update --from $spec
  if ($LASTEXITCODE -ne 0) { Fail "office-partner update failed (exit $LASTEXITCODE)." 30 }
}
Step 'Registering the add-in and starting the bridge'
$setupArgs = @('setup', '--providers', $Providers)
if ($env:OP_NO_SIDELOAD -eq '1') { $setupArgs += '--no-sideload' }
& $command @setupArgs
if ($LASTEXITCODE -ne 0) {
  $trustFailed = Select-String -LiteralPath $log -Pattern 'CERTIFICATE_TRUST_FAILED|Could not trust the localhost certificate' -Quiet
  if ($trustFailed) { Fail 'Windows did not trust the localhost certificate. Accept the security prompt, or ask IT whether user root certificates are blocked.' 40 }
  Fail "office-partner setup failed (exit $LASTEXITCODE)." 30
}
& $command status

Write-Host ""
Write-Host 'Office Partner is installed.' -ForegroundColor Green
Write-Host '1. Open Excel, Word or PowerPoint, then Home > Office Partner.'
Write-Host '2. First time: open Settings in the Office Partner pane and sign in to your AI provider.'
Write-Host "If something goes wrong, run: office-partner doctor   (install log: $log)"
Done 0
