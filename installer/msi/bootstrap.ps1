# Office Partner MSI: the steps an installer cannot do with files alone. Runs as the signed-in user (per-user MSI, no administrator rights).
#   -Action install    remove whatever an older version left behind, adopt or install the pinned Node.js, set up the parts
#                      ticked in the wizard (the add-in, the gateway, or both), install the pinned CLI, start them
#   -Action stop       stop the bridge, the gateway and every process this product started (runs before files are replaced)
#   -Action prepare    the same stop, then delete the folder an earlier copy (npm or an interrupted install) left at the install path
#   -Action uninstall  remove the registration, startup entries, shortcut, certificate trust and the Node.js this script added
#   -Action check      print what this machine has and change nothing
#   -Action status     show the result of the last install on demand
# Every version and SHA-256 comes from bin/runtime-manifest.json of the package being installed: nothing here says "latest".
# This script never fails the installation. A machine that already has an old copy, a foreign Node.js, a different npm prefix
# or a running bridge is repaired and overwritten; anything that still cannot be finished is written to the log and to
# install-status.json, and the person keeps a working (or clearly reported) installation. The MSI runs it with Return="ignore".
param(
  [ValidateSet('install', 'prepare', 'stop', 'uninstall', 'check', 'status')] [string] $Action = 'install',
  # The package folder is the parent of this script's folder (<package>\installer\bootstrap.ps1). It is resolved below:
  # $PSScriptRoot is not available while parameter defaults are bound in Windows PowerShell 5.1.
  [string] $Package,
  [string] $Providers = 'copilot',
  # The parts chosen in the wizard: addin, gateway, both (or empty to read the MSI's markers, see Get-Targets).
  [string] $Targets = '',
  [switch] $NoSideload
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
if ($ExecutionContext.SessionState.LanguageMode -eq 'FullLanguage') {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  # Corporate proxies that need the Windows sign-in.
  [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
}

# ---------------------------------------------------------------------------------------------------------------------
# Paths. A deferred MSI custom action inherits the environment of the Windows Installer service, so %LOCALAPPDATA% and
# friends can point at the SYSTEM profile. Ask Windows for the folders of the signed-in user instead.
# ---------------------------------------------------------------------------------------------------------------------
function Get-UserFolder([string] $special, [string] $variable) {
  $value = $null
  try { $value = [Environment]::GetFolderPath($special) } catch { }
  if (-not $value) {
    try { $value = (Get-ItemProperty -Path 'HKCU:\Volatile Environment' -ErrorAction Stop).$variable } catch { }
  }
  if (-not $value) { $value = [Environment]::GetEnvironmentVariable($variable) }
  return $value
}
$localAppData = Get-UserFolder 'LocalApplicationData' 'LOCALAPPDATA'
$appData = Get-UserFolder 'ApplicationData' 'APPDATA'
$userProfile = Get-UserFolder 'UserProfile' 'USERPROFILE'
$programs = Join-Path $localAppData 'Programs'

if (-not $Package) {
  $scriptRoot = $PSScriptRoot
  if (-not $scriptRoot) { try { $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path } catch { } }
  if ($scriptRoot) { $Package = Split-Path -Parent $scriptRoot } else { $Package = (Get-Location).Path }
}
$Package = $Package.TrimEnd('\')

$dataDir = Join-Path $localAppData 'Excel Partner'
$nodeHome = Join-Path $programs 'nodejs'
$altNodeHome = Join-Path $programs 'office-partner-node'
$marker = Join-Path $nodeHome '.installed-by-office-partner'
$altMarker = Join-Path $altNodeHome '.installed-by-office-partner'
$cli = Join-Path $Package 'bin\office-partner.mjs'
$npmCmd = Join-Path $nodeHome 'npm.cmd'
$gatewayDir = Join-Path $localAppData 'Agent Gateway'
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$developerKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$stateKey = 'HKCU:\Software\Office Partner'

# ---------------------------------------------------------------------------------------------------------------------
# Logging. Written where the person can find it; falls back to %TEMP% when the data folder cannot be created yet.
# ---------------------------------------------------------------------------------------------------------------------
$logFile = Join-Path $dataDir 'msi-install.log'
try { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null } catch { $logFile = Join-Path ([IO.Path]::GetTempPath()) 'office-partner-msi-install.log' }
$problems = New-Object System.Collections.ArrayList
. (Join-Path $PSScriptRoot 'stop-processes.ps1')
function Add-Problem([string] $message) { [void]$problems.Add($message); Write-Log "PROBLEM: $message" }
function Step([string] $name, [scriptblock] $work) {
  Write-Log "==> $name"
  try { & $work } catch { Add-Problem "${name}: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------------------------------------------------------------
# Which parts the wizard ticked. The MSI writes one registry marker per chosen feature before this script runs, so the
# wizard, a silent install (ADDLOCAL=Files,AddIn,Gateway) and a manual run all end up in the same code path.
# ---------------------------------------------------------------------------------------------------------------------
function Get-Targets {
  if ($Targets) { return $Targets }
  $selected = New-Object System.Collections.ArrayList
  foreach ($name in 'AddIn', 'Gateway') {
    $value = $null
    try { $value = (Get-ItemProperty -Path $stateKey -Name $name -ErrorAction Stop).$name } catch { }
    if ("$value" -eq '1') { [void]$selected.Add($name.ToLower()) }
  }
  # The marker this MSI always writes: the person really chose nothing (or chose only one part).
  $fromThisInstaller = $null
  try { $fromThisInstaller = (Get-ItemProperty -Path $stateKey -Name 'Selected' -ErrorAction Stop).Selected } catch { }
  if ("$fromThisInstaller" -eq '1') { return ($selected -join ',') }
  # No marker at all: an installer from before the choice existed, or a person running this script by hand.
  return 'addin'
}

Write-Log ''
Write-Log "==== Office Partner $Action ===="
Write-Log "user=$([Environment]::UserName) powershell=$($PSVersionTable.PSVersion)"
Write-Log "package=$Package"
Write-Log "data=$dataDir"
Write-Log "env LOCALAPPDATA=$env:LOCALAPPDATA (resolved $localAppData), APPDATA=$env:APPDATA (resolved $appData)"

$targets = Get-Targets
$wantAddin = $targets -match 'addin'
$wantGateway = $targets -match 'gateway'
Write-Log "targets=$targets (addin=$wantAddin gateway=$wantGateway)"

# ---------------------------------------------------------------------------------------------------------------------
# npm's global folder for the Node.js used by this installation.
# ---------------------------------------------------------------------------------------------------------------------
function Get-NpmRoot([string] $nodeExe) {
  try {
    $npm = Join-Path (Split-Path -Parent $nodeExe) 'npm.cmd'
    if (Test-Path -LiteralPath $npm) { return (& $npm root -g 2>$null | Select-Object -First 1).Trim() }
    return (& $nodeExe -e "process.stdout.write(require('path').join(process.env.APPDATA || '', 'npm', 'node_modules'))" 2>$null).Trim()
  } catch { return $null }
}

# ---------------------------------------------------------------------------------------------------------------------
# Previous installations. Any of these may exist on a machine that already used the product; all of them go.
# ---------------------------------------------------------------------------------------------------------------------
function Remove-PreviousInstall([string] $nodeExe) {
  $candidates = New-Object System.Collections.ArrayList
  $recorded = $null
  try { $recorded = (Get-ItemProperty -Path $stateKey -Name InstallDir -ErrorAction Stop).InstallDir } catch { }
  if ($recorded) { [void]$candidates.Add($recorded) }
  [void]$candidates.Add((Join-Path $appData 'npm\node_modules\office-partner'))
  [void]$candidates.Add((Join-Path $nodeHome 'node_modules\office-partner'))
  if ($nodeExe) { $npmRoot = Get-NpmRoot $nodeExe; if ($npmRoot) { [void]$candidates.Add((Join-Path $npmRoot 'office-partner')) } }
  foreach ($candidate in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
    $path = "$candidate".TrimEnd('\')
    if ($path.ToLower() -eq $Package.ToLower()) { continue }
    if (-not (Test-Path -LiteralPath $path)) { continue }
    Write-Log "    removing an older installation: $path"
    try {
      $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
      if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { [IO.Directory]::Delete($path) }
      else { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop }
    } catch { Write-Log "    could not remove $path now: $($_.Exception.Message)" }
  }
  # Command shims of older versions, in every folder that may be on PATH. The shim this installer placed next to the
  # package (in the prefix, one level above node_modules) stays.
  $prefixDir = Split-Path -Parent (Split-Path -Parent $Package)
  $keepShim = [IO.Path]::GetFullPath((Join-Path $prefixDir 'office-partner.cmd'))
  $binDirs = New-Object System.Collections.ArrayList
  [void]$binDirs.Add((Join-Path $appData 'npm'))
  [void]$binDirs.Add($nodeHome)
  if ($nodeExe) { $npmRoot = Get-NpmRoot $nodeExe; if ($npmRoot) { [void]$binDirs.Add((Split-Path -Parent $npmRoot)) } }
  foreach ($dir in ($binDirs | Where-Object { $_ } | Select-Object -Unique)) {
    foreach ($name in 'office-partner', 'office-partner-doctor', 'agent-gateway') {
      foreach ($extension in '', '.cmd', '.ps1') {
        $file = Join-Path $dir ($name + $extension)
        if (-not (Test-Path -LiteralPath $file)) { continue }
        if ([IO.Path]::GetFullPath($file).ToLower() -eq $keepShim.ToLower()) { continue }
        Write-Log "    removing an old command: $file"
        try { Remove-Item -LiteralPath $file -Force -ErrorAction Stop } catch { Write-Log "    could not remove ${file}: $($_.Exception.Message)" }
      }
    }
  }
}

function Remove-OldStartup {
  foreach ($name in 'Excel Partner Bridge', 'Agent Gateway') {
    try { Remove-ItemProperty -Path $runKey -Name $name -ErrorAction SilentlyContinue } catch { }
  }
}

# ---------------------------------------------------------------------------------------------------------------------
# Node.js: the pinned release is the one the shipped SQLite binary is built for (the ABI must match). An existing Node.js
# is adopted when its ABI matches, whatever its exact version; nothing the person installed is ever deleted.
# ---------------------------------------------------------------------------------------------------------------------
function Read-Manifest {
  $file = Join-Path $Package 'bin\runtime-manifest.json'
  if (-not (Test-Path -LiteralPath $file)) { throw "The package has no bin\runtime-manifest.json ($Package)." }
  return Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
}
function Get-NodeAbi([string] $exe) {
  try { return (& $exe -p "process.versions.modules" 2>$null | Select-Object -First 1).Trim() } catch { return $null }
}
function Get-NodeVersion([string] $exe) {
  try { return (& $exe --version 2>$null | Select-Object -First 1).Trim() } catch { return $null }
}

function Install-PinnedNode($manifest) {
  $version = $manifest.node.version
  $expected = $manifest.node.win32.x64.sha256
  if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { $expected = $manifest.node.win32.arm64.sha256 }
  if (-not $version -or -not $expected) { throw 'The manifest pins no Node.js for this architecture.' }
  $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { 'arm64' } else { 'x64' }
  $existing = Join-Path $nodeHome 'node.exe'
  $target = if (-not (Test-Path -LiteralPath $existing)) { $nodeHome } else { $altNodeHome }
  Write-Log "    downloading Node.js v$version into $target (pinned, SHA-256 checked)"
  $name = "node-v$version-win-$arch"
  $work = Join-Path ([IO.Path]::GetTempPath()) "office-partner-node-$([guid]::NewGuid().ToString('N'))"
  New-Item -ItemType Directory -Path $work -Force | Out-Null
  try {
    $zip = Join-Path $work "$name.zip"
    $source = if ($env:OP_NODE_MIRROR) { $env:OP_NODE_MIRROR.TrimEnd('/') } else { 'https://nodejs.org/dist' }
    $downloaded = $false
    for ($attempt = 1; $attempt -le 3 -and -not $downloaded; $attempt++) {
      try {
        # Progress must be off: it makes Invoke-WebRequest far slower and can time out on a slow network.
        $old = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
        try { Invoke-WebRequest -Uri "$source/v$version/$name.zip" -OutFile $zip -UseBasicParsing } finally { $ProgressPreference = $old }
        $downloaded = $true
      } catch {
        Write-Log "    download attempt $attempt failed: $($_.Exception.Message)"
        Start-Sleep -Seconds (3 * $attempt)
      }
    }
    if (-not $downloaded) { throw "Could not download Node.js v$version from ${source}. Check the network or proxy, or set OP_NODE_MIRROR to an internal mirror of nodejs.org/dist." }
    $actual = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
    if ($actual.ToLower() -ne $expected.ToLower()) { throw 'The downloaded Node.js does not match the SHA-256 pinned in this release. Nothing was installed.' }
    Expand-Archive -LiteralPath $zip -DestinationPath $work -Force
    if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
    Move-Item -LiteralPath (Join-Path $work $name) -Destination $target
    Set-Content -LiteralPath (Join-Path $target '.installed-by-office-partner') -Value $version
    return (Join-Path $target 'node.exe')
  } finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}

function Ensure-Node($manifest, [switch] $NoDownload) {
  $wanted = "$($manifest.node.abi)"
  $candidates = @(
    (Join-Path $nodeHome 'node.exe'),
    (Join-Path $altNodeHome 'node.exe')
  )
  foreach ($candidate in $candidates) {
    if (-not (Test-Path -LiteralPath $candidate)) { continue }
    $abi = Get-NodeAbi $candidate
    if ($abi -eq $wanted) { Write-Log "    using $candidate ($(Get-NodeVersion $candidate))"; return $candidate }
    Write-Log "    $candidate has ABI $abi, this release needs $wanted"
  }
  $onPath = $null
  try { $onPath = (Get-Command node.exe -ErrorAction Stop).Source } catch { }
  if ($onPath -and (Get-NodeAbi $onPath) -eq $wanted) { Write-Log "    using the Node.js on PATH: $onPath ($(Get-NodeVersion $onPath))"; return $onPath }
  if ($NoDownload) { Write-Log "    no Node.js with ABI $wanted on this machine yet"; return $null }
  return (Install-PinnedNode $manifest)
}

# ---------------------------------------------------------------------------------------------------------------------
# npm's global folder: `office-partner update` replaces the package through npm, so npm must see the package where it
# expects it. A machine whose npm prefix points somewhere else gets a junction instead of an error message.
# ---------------------------------------------------------------------------------------------------------------------
function Ensure-NpmLayout([string] $nodeExe) {
  if (-not $nodeExe) { return }
  $npmRoot = Get-NpmRoot $nodeExe
  if (-not $npmRoot) { Write-Log '    npm root -g is unavailable; the package stays where the installer put it'; return }
  $packageParent = Split-Path -Parent $Package
  if ($npmRoot.TrimEnd('\').ToLower() -eq $packageParent.TrimEnd('\').ToLower()) { Write-Log "    npm root -g already is $npmRoot"; return }
  Write-Log "    npm root -g is $npmRoot, the package is in ${packageParent}: linking it where npm looks"
  New-Item -ItemType Directory -Path $npmRoot -Force | Out-Null
  $link = Join-Path $npmRoot 'office-partner'
  if (Test-Path -LiteralPath $link) {
    $item = Get-Item -LiteralPath $link -Force
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { [IO.Directory]::Delete($link) }
    else { Remove-Item -LiteralPath $link -Recurse -Force }
  }
  New-Item -ItemType Junction -Path $link -Value $Package | Out-Null
  # The command shim belongs next to npm's binaries (the prefix), which is the parent of its node_modules folder.
  $prefix = Split-Path -Parent $npmRoot
  $source = Join-Path $packageParent 'office-partner.cmd'
  $destination = Join-Path $prefix 'office-partner.cmd'
  if ((Test-Path -LiteralPath $source) -and $source.ToLower() -ne $destination.ToLower()) { Copy-Item -LiteralPath $source -Destination $destination -Force }
}

# ---------------------------------------------------------------------------------------------------------------------
# Certificate, add-in registration and the bridge: the product's own setup does it; this retries and never throws.
# ---------------------------------------------------------------------------------------------------------------------
function Invoke-Cli([string[]] $arguments) {
  # Run the CLI without PowerShell pipelines: `office-partner setup` leaves the bridge and the gateway running, and a
  # child that inherits a pipe would keep the caller waiting for the pipe to close. cmd.exe appends the CLI output to
  # the log file and we wait on the process handle only, with a hard cap. CreateNoWindow gives cmd.exe no console at
  # all; -WindowStyle Hidden alone still lets Windows Terminal show one on Windows 11.
  $quoted = ($arguments | ForEach-Object { '"' + $_ + '"' }) -join ' '
  $shellExe = $env:ComSpec
  if (-not $shellExe -or -not (Test-Path -LiteralPath $shellExe)) { $shellExe = Join-Path $env:SystemRoot 'System32\cmd.exe' }
  $command = "`"$quoted <nul >> `"$logFile`" 2>&1`""
  try {
    $start = New-Object System.Diagnostics.ProcessStartInfo
    $start.FileName = $shellExe
    $start.Arguments = "/d /c " + $command
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.WorkingDirectory = $Package
    $process = [System.Diagnostics.Process]::Start($start)
  } catch { throw "Could not run ${shellExe}: $($_.Exception.Message)" }
  if (-not $process.WaitForExit(1200000)) {
    Write-Log '    the command did not finish in 20 minutes; stopping it'
    try { $process.Kill() } catch { }
    return -1
  }
  $process.WaitForExit()
  return $process.ExitCode
}

function Invoke-Setup([string] $nodeExe, [string] $providerList, [string] $target) {
  if (-not (Test-Path -LiteralPath $cli)) { throw "The package is not where the installer put it: $cli" }
  $arguments = @($nodeExe, $cli, 'setup', $target)
  if ($providerList -and $providerList -ne 'none') { $arguments += @('--providers', $providerList) } else { $arguments += @('--skip-providers') }
  if ($NoSideload) { $arguments += '--no-sideload' }
  for ($attempt = 1; $attempt -le 2; $attempt++) {
    Write-Log "    office-partner setup $target (attempt $attempt)"
    $code = Invoke-Cli $arguments
    Write-Log "    exit $code"
    if ($code -eq 0) { return $true }
    Start-Sleep -Seconds 3
  }
  # A CLI download that fails must not keep the add-in from being registered or the gateway from starting.
  Write-Log '    retrying without the provider CLI'
  $fallback = @($nodeExe, $cli, 'setup', $target, '--skip-providers')
  if ($NoSideload) { $fallback += '--no-sideload' }
  $code = Invoke-Cli $fallback
  Write-Log "    exit $code"
  if ($code -eq 0) { return $true }
  Write-Log "    starting $target directly"
  $code = Invoke-Cli @($nodeExe, $cli, 'start', $target)
  Write-Log "    start exit $code"
  return ($code -eq 0)
}

# Runs one CLI command and reports it in the log; any failure becomes a problem line, never an exception.
function Invoke-Product([string] $nodeExe, [string[]] $arguments) {
  if (-not (Test-Path -LiteralPath $cli)) { throw "The package is not where the installer put it: $cli" }
  $code = Invoke-Cli (@($nodeExe, $cli) + $arguments)
  Write-Log "    $($arguments -join ' ') -> exit $code"
  return $code
}

function Write-Status([string] $nodeExe, [bool] $setupOk) {
  $status = [ordered]@{
    at = (Get-Date).ToString('s'); action = $Action; targets = $targets
    addin = [bool]$wantAddin; gateway = [bool]$wantGateway; providers = $Providers
    package = $Package; node = $nodeExe; setupOk = $setupOk
    problems = @($problems); log = $logFile
  }
  try { $status | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $dataDir 'install-status.json') -Encoding UTF8 } catch { }
  Write-Log "status: setupOk=$setupOk problems=$($problems.Count)"
}

function Test-LocalPort([int] $port) {
  $client = New-Object Net.Sockets.TcpClient
  try { $client.Connect('127.0.0.1', $port); return $client.Connected } catch { return $false } finally { $client.Dispose() }
}

# Show which parts were set up, whether they are answering, and where the log lives.
function Show-InstallStatus {
  try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    $status = $null
    try { $status = Get-Content -LiteralPath (Join-Path $dataDir 'install-status.json') -Raw | ConvertFrom-Json } catch { }
    $setupOk = if ($status) { [bool]$status.setupOk } else { $false }
    $addin = if ($status) { [bool]$status.addin } else { $true }
    $gateway = if ($status) { [bool]$status.gateway } else { $false }
    $problemCount = if ($status) { @($status.problems).Count } else { 0 }
    $addinLine = if (-not $addin) { 'not selected by this install' }
      elseif (Test-LocalPort 8765) { 'running at https://localhost:8765' }
      else { 'installed, but the bridge is not answering' }
    $gatewayLine = if (-not $gateway) { 'not selected by this install' }
      elseif (Test-LocalPort 8766) { 'running at https://localhost:8766' }
      else { 'installed, but not answering on port 8766' }
    $nodeExe = if ($status -and $status.node) { $status.node } else { Join-Path $nodeHome 'node.exe' }
    $nodeVersion = try { (& $nodeExe --version 2>$null | Select-Object -First 1) } catch { $null }
    $providerList = if ($status -and $status.providers) { $status.providers } else { $Providers }
    $lines = @(
      $(if (-not $addin -and -not $gateway) { 'The files are installed, but no part was selected, so nothing is running.' }
        elseif ($setupOk) { 'Office Partner is installed and ready.' }
        else { 'The files are installed, but setup did not finish.' }),
      '',
      "Add-in      : $addinLine",
      "Gateway     : $gatewayLine",
      "Node.js     : $nodeExe $nodeVersion",
      "Providers   : $providerList",
      "Installed in: $Package",
      "Your data   : $dataDir (kept through updates and reinstalls)",
      "Log         : $logFile"
    )
    if ($problemCount -gt 0) { $lines += ''; $lines += "$problemCount step(s) reported a problem; open the log above." }
    $icon = if ($setupOk) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Warning }
    [void][System.Windows.Forms.MessageBox]::Show(($lines -join "`r`n"), $(if ($setupOk) { 'Office Partner installed' } else { 'Office Partner: setup needs attention' }), [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
  } catch { Write-Log "status window: $($_.Exception.Message)" }
}

# The product deletes localhost.crt before the trust entry is removed, so read the thumbprint while the file is there.
function Get-CertificateThumbprint {
  $crt = Join-Path $dataDir 'localhost.crt'
  if (-not (Test-Path -LiteralPath $crt)) { return $null }
  try { return (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $crt).Thumbprint } catch { return $null }
}

function Remove-CertificateTrust([string] $thumbprint) {
  if (-not $thumbprint) { Write-Log '    no localhost certificate found; nothing to untrust'; return }
  & certutil.exe -user -delstore Root $thumbprint 2>&1 | Out-Null
  if ($LASTEXITCODE -eq 0) { Write-Log "    removed $thumbprint" } else { Write-Log "    certutil exit $LASTEXITCODE for $thumbprint" }
}

# Whether an earlier version registered the add-in for this user (used when this install did not ask for it).
function Test-AddInRegistered {
  try { if ((Get-ItemProperty -Path $runKey -Name 'Excel Partner Bridge' -ErrorAction Stop).'Excel Partner Bridge') { return $true } } catch { }
  try {
    foreach ($property in (Get-ItemProperty -Path $developerKey -ErrorAction SilentlyContinue).PSObject.Properties) {
      if ($property.Name -notlike 'PS*' -and "$($property.Value)" -like '*Excel Partner*') { return $true }
    }
  } catch { }
  return (Test-Path -LiteralPath (Join-Path $dataDir 'localhost.crt'))
}

# ---------------------------------------------------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------------------------------------------------
if ([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem -or -not $localAppData) {
  Write-Log 'ERROR: Office Partner installs into one person''s profile. Use Install behavior = User (Intune), not SYSTEM.'
  exit 60
}

if ($Action -eq 'check') {
  Step 'Look for a usable Node.js' { $script:node = Ensure-Node (Read-Manifest) -NoDownload }
  Write-Log "node: $node"
  Write-Log "npm root -g: $(if ($node) { Get-NpmRoot $node })"
  Write-Log "package: $Package ($(if (Test-Path $cli) { 'present' } else { 'MISSING' }))"
  Write-Log "data: $dataDir ($(@(Get-ChildItem -LiteralPath $dataDir -ErrorAction SilentlyContinue).Count) entries)"
  foreach ($candidate in @((Join-Path $appData 'npm\node_modules\office-partner'), (Join-Path $nodeHome 'node_modules\office-partner'), (Join-Path $localAppData 'Programs\office-partner-node'))) {
    Write-Log "found: $candidate = $(Test-Path -LiteralPath $candidate)"
  }
  exit 0
}

if ($Action -eq 'status') {
  # Started by -Action install once the wizard is done; waits for the wizard to close so the result lands on top of the
  # desktop rather than behind it, then shows what the person needs to know and exits.
  $deadline = (Get-Date).AddMinutes(10)
  while ((Get-Date) -lt $deadline) {
    $wizard = @(Get-CimInstance Win32_Process -Filter "Name='msiexec.exe'" -ErrorAction SilentlyContinue |
      Where-Object { "$($_.CommandLine)" -match '(?i)office-partner' })
    if ($wizard.Count -eq 0) { break }
    Start-Sleep -Seconds 1
  }
  Start-Sleep -Seconds 1
  Show-InstallStatus
  exit 0
}

if ($Action -eq 'prepare') {
  # Runs before the files are copied. A copy left by `npm install -g` (or by a failed install) at the install path is
  # deleted so the new files land in a clean folder. A folder an earlier Office Partner MSI owns is left alone: Windows
  # Installer replaces its own files, and deleting them behind its back would leave the install with an empty folder
  # because it believes those components are already up to date.
  Step 'Stop the bridge, the gateway and their helpers' { Stop-ProductProcesses }
  Step 'Clear the folder an older copy left behind' {
    $owner = $null
    try { $owner = (Get-ItemProperty -Path $stateKey -Name 'InstallDir' -ErrorAction Stop).InstallDir } catch { }
    if ($owner -and $owner.TrimEnd('\') -ieq $Package) {
      Write-Log "    $Package is owned by an installed Office Partner; its own files replace it"
    } elseif (Test-Path -LiteralPath $Package) {
      Write-Log "    removing $Package"
      $item = Get-Item -LiteralPath $Package -Force
      if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { [IO.Directory]::Delete($Package) }
      else { Remove-Item -LiteralPath $Package -Recurse -Force }
    }
  }
  exit 0
}

if ($Action -eq 'stop') {
  Step 'Stop the bridge, the gateway and their helpers' { Stop-ProductProcesses }
  exit 0
}

if ($Action -eq 'uninstall') {
  Step 'Stop the bridge, the gateway and their helpers' { Stop-ProductProcesses }
  # The product removes its own certificate files, so read the thumbprint while localhost.crt still exists.
  $script:certThumbprint = $null
  Step 'Read the localhost certificate thumbprint' { $script:certThumbprint = Get-CertificateThumbprint }
  Step 'Let the product remove its registration, startup entry, shortcut and certificate trust' {
    $node = $null
    foreach ($candidate in @((Join-Path $nodeHome 'node.exe'), (Join-Path $altNodeHome 'node.exe'))) { if (Test-Path $candidate) { $node = $candidate; break } }
    if ($node -and (Test-Path -LiteralPath $cli)) { & $node $cli uninstall all 2>&1 | Out-Host }
  }
  Step 'Remove the startup entries' { Remove-OldStartup }
  Step 'Remove the add-in registration' {
    if (Test-Path $developerKey) {
      $properties = Get-ItemProperty -Path $developerKey
      foreach ($property in $properties.PSObject.Properties) {
        if ($property.Name -like 'PS*') { continue }
        if ("$($property.Value)" -like "*Excel Partner*") { Remove-ItemProperty -Path $developerKey -Name $property.Name -ErrorAction SilentlyContinue }
      }
    }
  }
  Step 'Remove the Fix connection shortcut' { Remove-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('Programs')) 'Office Partner - Fix connection.lnk') -Force -ErrorAction SilentlyContinue }
  Step 'Remove the localhost certificate trust' { Remove-CertificateTrust $certThumbprint }
  Step 'Remove links and commands from other npm folders' {
    $node = $null
    foreach ($candidate in @((Join-Path $nodeHome 'node.exe'), (Join-Path $altNodeHome 'node.exe'))) { if (Test-Path $candidate) { $node = $candidate; break } }
    if ($node) {
      $npmRoot = Get-NpmRoot $node
      if ($npmRoot) {
        $link = Join-Path $npmRoot 'office-partner'
        if (Test-Path -LiteralPath $link) {
          $item = Get-Item -LiteralPath $link -Force
          if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { [IO.Directory]::Delete($link) }
        }
        foreach ($name in 'office-partner', 'office-partner-doctor', 'agent-gateway') {
          foreach ($extension in '', '.cmd', '.ps1') { Remove-Item -LiteralPath (Join-Path (Split-Path -Parent $npmRoot) ($name + $extension)) -Force -ErrorAction SilentlyContinue }
        }
      }
    }
  }
  Step 'Remove the Node.js this installer added' {
    foreach ($folder in @(@{ dir = $altNodeHome; marker = $altMarker }, @{ dir = $nodeHome; marker = $marker })) {
      if (-not (Test-Path -LiteralPath $folder.marker)) { continue }
      $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
      $parts = @($userPath -split ';' | Where-Object { $_ -and $_.TrimEnd('\').ToLower() -ne $folder.dir.TrimEnd('\').ToLower() })
      [Environment]::SetEnvironmentVariable('Path', ($parts -join ';'), 'User')
      Remove-Item -LiteralPath $folder.dir -Recurse -Force -ErrorAction SilentlyContinue
      Write-Log "    removed $($folder.dir)"
    }
  }
  Write-Log ''
  Write-Log "Uninstalled. Diagnostics and logs stay in $dataDir"
  exit 0
}

# Install
try {
  $manifest = $null
  Step 'Read the versions this release pins' { $script:manifest = Read-Manifest }
  if (-not $manifest) { Write-Status $null $false; Write-Log 'The package is incomplete; see the log above.'; exit 0 }

  Step 'Stop the bridge, the gateway and their helpers' { Stop-ProductProcesses }
  Step 'Remove startup entries of older versions' { Remove-OldStartup }

  $node = $null
  Step 'Adopt an existing Node.js or install the pinned one' { $script:node = Ensure-Node $manifest }
  if (-not $node) { Write-Status $null $false; Write-Log 'No usable Node.js; the add-in files are installed but nothing was started.'; exit 0 }
  $nodeDir = Split-Path -Parent $node
  $env:Path = "$nodeDir;$env:Path"

  Step 'Remove older installations and their commands' { Remove-PreviousInstall $node }
  Step 'Make npm see the package where it looks for it' { Ensure-NpmLayout $node }
  Remove-OldStartup

  $setupTarget = if ($wantAddin -and $wantGateway) { 'all' } elseif ($wantGateway) { 'gateway' } elseif ($wantAddin) { 'addin' } else { '' }
  $setupOk = $true
  if ($setupTarget) {
    Step "Set up and start the selected parts ($setupTarget)" { $script:setupOk = Invoke-Setup $node $Providers $setupTarget }
  } else {
    Write-Log 'Nothing was selected in the installer: the files are installed and no part was set up.'
  }

  # A part the person did not tick is taken away when an earlier install set it up: the choice is the source of truth.
  if (-not $wantAddin) {
    Step 'Remove the add-in registration this install did not ask for' {
      if (-not (Test-AddInRegistered)) { Write-Log '    no add-in registration on this machine' }
      else {
        # Read the thumbprint before the product deletes localhost.crt, then untrust it (same bug as a full uninstall).
        $thumbprint = Get-CertificateThumbprint
        Invoke-Product $node @('uninstall', 'addin') | Out-Null
        Remove-CertificateTrust $thumbprint
      }
    }
  }
  if (-not $wantGateway) {
    Step 'Stop the Agent Gateway this install did not ask for' {
      if (-not (Test-Path -LiteralPath (Join-Path $gatewayDir 'token.txt'))) { Write-Log '    the gateway was never set up on this machine' }
      else { Invoke-Product $node @('gateway', 'uninstall') | Out-Null }
    }
  }

  Write-Status $node $setupOk
  if ($setupOk) {
    Write-Log ''
    Write-Log "Office Partner is installed ($targets). Open Excel, Word or PowerPoint, then Home > Office Partner, and sign in from Settings."
  } else {
    Write-Log ''
    Write-Log "The files are installed, but the last step did not finish. Open 'Office Partner - Fix connection' from the Start Menu, or run: `"$node`" `"$cli`" fix"
    Write-Log "Log: $logFile"
  }
  # The MSI owns completion UI. A status child can keep QuietExec alive after setup finishes.
} catch {
  Add-Problem "unexpected: $($_.Exception.Message)"
  Write-Log $_.ScriptStackTrace
}
exit 0
