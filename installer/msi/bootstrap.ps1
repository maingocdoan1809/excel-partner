# Office Partner MSI: the steps an installer cannot do with files alone. Runs as the signed-in user (per-user MSI, no administrator rights).
#   -Action install    download the pinned Node.js, register the add-in, install the pinned Copilot CLI, start the bridge
#   -Action stop       stop the bridge (before files are replaced by an upgrade)
#   -Action uninstall  remove the registration, startup entry and certificate trust, and the Node.js this script added
# Every version and SHA-256 comes from bin/runtime-manifest.json of the package being installed: nothing here says "latest".
# Exit codes (same as installer/install-office-partner.cmd): 0 ok, 10 Node.js, 30 setup, 40 certificate trust, 60 unsupported context.
param(
  [Parameter(Mandatory = $true)] [ValidateSet('install', 'stop', 'uninstall')] [string] $Action,
  # The package folder is the parent of this script's folder (<package>\installer\bootstrap.ps1); a trailing backslash on a command-line path would break quoting.
  [string] $Package = (Split-Path -Parent $PSScriptRoot),
  [string] $Providers = 'copilot',
  [switch] $NoSideload
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
if ($ExecutionContext.SessionState.LanguageMode -eq 'FullLanguage') {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  # Corporate proxies that need the Windows sign-in.
  [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
}

$dataDir = Join-Path $env:LOCALAPPDATA 'Excel Partner'
$nodeHome = Join-Path $env:LOCALAPPDATA 'Programs\nodejs'
$marker = Join-Path $nodeHome '.installed-by-office-partner'
$cli = Join-Path $Package 'bin\office-partner.mjs'
$nodeExe = Join-Path $nodeHome 'node.exe'
New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
$log = Join-Path $dataDir 'msi-install.log'
Start-Transcript -Path $log -Append | Out-Null

function Done([int] $code) { Stop-Transcript | Out-Null; exit $code }
function Step([string] $message) { Write-Host ''; Write-Host "==> $message" -ForegroundColor Cyan }
function Fail([string] $message, [int] $code = 1) { Write-Host ''; Write-Host "ERROR: $message" -ForegroundColor Red; Write-Host "Log: $log"; Done $code }

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if ($identity.IsSystem -or -not $env:LOCALAPPDATA) { Fail 'Install Office Partner as the signed-in user (Intune: Install behavior = User), not as SYSTEM.' 60 }

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

# A security program scans each new file and can deny access for a moment: try again a few times before giving up.
function Invoke-Retry([string] $what, [scriptblock] $work, [int] $attempts = 4) {
  for ($try = 1; $try -le $attempts; $try++) {
    try { & $work; return }
    catch {
      if ($try -eq $attempts) { Fail "$what failed: $($_.Exception.Message). A security program is probably blocking it; ask IT to allow %LOCALAPPDATA%\Programs\nodejs and %LOCALAPPDATA%\Excel Partner." 10 }
      Write-Host "$what was denied (attempt $try of $attempts). Waiting for the security scan..." -ForegroundColor Yellow
      Start-Sleep -Seconds (2 * $try)
    }
  }
}

function Read-Manifest {
  $file = Join-Path $Package 'bin\runtime-manifest.json'
  if (-not (Test-Path -LiteralPath $file)) { Fail "The package has no bin\runtime-manifest.json ($Package)." 10 }
  return Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
}

# The exact Node.js release of the manifest, checked against the SHA-256 written in the manifest itself (not one fetched from the download site).
function Install-PinnedNode($manifest) {
  $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { 'arm64' } else { 'x64' }
  $version = $manifest.node.version
  $expected = $manifest.node.win32.$arch.sha256
  if (-not $version -or -not $expected) { Fail "The manifest pins no Node.js for win-$arch." 10 }
  if ((Test-Path -LiteralPath $nodeExe) -and ((& $nodeExe --version 2>$null) -eq "v$version")) { Write-Host "Node.js v$version is already installed: $nodeExe"; return }
  if ((Test-Path -LiteralPath $nodeExe) -and -not (Test-Path -LiteralPath $marker)) { Fail "$nodeHome holds a Node.js this installer did not put there. Remove it or move it, then run the installer again." 10 }
  Step "Downloading Node.js v$version for this user only (pinned, no administrator rights)"
  $name = "node-v$version-win-$arch"
  $work = Join-Path $env:TEMP "office-partner-node-$([guid]::NewGuid().ToString('N'))"
  New-Item -ItemType Directory -Path $work -Force | Out-Null
  try {
    $zip = Join-Path $work "$name.zip"
    $source = if ($env:OP_NODE_MIRROR) { $env:OP_NODE_MIRROR.TrimEnd('/') } else { 'https://nodejs.org/dist' }
    try { Invoke-WebRequest -Uri "$source/v$version/$name.zip" -OutFile $zip -UseBasicParsing }
    catch { Fail "Could not download Node.js v$version from ${source}: $($_.Exception.Message). Check the network or proxy, or set OP_NODE_MIRROR to an internal mirror of nodejs.org/dist." 10 }
    $actual = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
    if ($actual.ToLower() -ne $expected.ToLower()) { Fail 'The downloaded Node.js does not match the SHA-256 pinned in this release. Nothing was installed.' 10 }
    Invoke-Retry 'Unpacking Node.js' { Expand-Archive -LiteralPath $zip -DestinationPath $work -Force }
    Invoke-Retry 'Replacing the old Node.js folder' { if (Test-Path -LiteralPath $nodeHome) { Remove-Item -LiteralPath $nodeHome -Recurse -Force } }
    New-Item -ItemType Directory -Path (Split-Path $nodeHome) -Force | Out-Null
    Invoke-Retry 'Moving Node.js into place' { Move-Item -LiteralPath (Join-Path $work $name) -Destination $nodeHome }
    Set-Content -LiteralPath $marker -Value $version
  } finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}

if (-not (Test-Path -LiteralPath $cli)) { if ($Action -eq 'install') { Fail "The package is not where the installer put it: $cli" 10 } else { Done 0 } }

if ($Action -eq 'stop') {
  if (Test-Path -LiteralPath $nodeExe) { & $nodeExe $cli stop 2>&1 | Out-Host }
  Done 0
}

if ($Action -eq 'uninstall') {
  Step 'Removing the add-in registration, startup entry and certificate trust'
  if (Test-Path -LiteralPath $nodeExe) { & $nodeExe $cli uninstall 2>&1 | Out-Host }
  if (Test-Path -LiteralPath $marker) {
    Step 'Removing the Node.js this installer added'
    Remove-UserPath $nodeHome
    Remove-Item -LiteralPath $nodeHome -Recurse -Force -ErrorAction SilentlyContinue
  }
  Done 0
}

$manifest = Read-Manifest
Install-PinnedNode $manifest
Add-UserPath $nodeHome
Add-UserPath (Join-Path $env:APPDATA 'npm')
Write-Host "Node.js $(& $nodeExe --version)"

# `office-partner update` needs the package in the folder npm calls its global root; this is what the MSI file layout assumes.
$npm = Join-Path $nodeHome 'npm.cmd'
$globalRoot = (& $npm root -g).Trim()
$packageParent = Split-Path -Parent $Package
if ($globalRoot.TrimEnd('\').ToLower() -ne $packageParent.TrimEnd('\').ToLower()) {
  Fail "npm's global folder is $globalRoot but the package is in $packageParent, so 'office-partner update' would not find it. Remove any npm 'prefix' setting (npm config delete prefix), then install again." 30
}

Step "Registering the add-in, installing the pinned Copilot CLI and starting the bridge ($Providers)"
$setupArgs = @($cli, 'setup', '--providers', $Providers)
if ($NoSideload) { $setupArgs += '--no-sideload' }
& $nodeExe @setupArgs
if ($LASTEXITCODE -ne 0) {
  if (Select-String -LiteralPath $log -Pattern 'CERTIFICATE_TRUST_FAILED|Could not trust the localhost certificate' -Quiet) { Fail 'Windows did not trust the localhost certificate. Accept the security prompt, or ask IT whether user root certificates are blocked.' 40 }
  Fail "office-partner setup failed (exit $LASTEXITCODE)." 30
}
Write-Host ''
Write-Host 'Office Partner is installed. Open Excel, Word or PowerPoint, then Home > Office Partner, and sign in from Settings.' -ForegroundColor Green
Done 0
