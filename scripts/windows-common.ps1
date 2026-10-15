# Shared helpers for the Windows install/update scripts (excel-partner).
# Dot-source this file; it defines functions and the provider table only.

function Write-Step($m) { Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-Info($m) { Write-Host "    $m" }
function Fail($m) { Write-Host "`nERROR: $m" -ForegroundColor Red; exit 1 }

$script:ProviderSpecs = @{
  copilot = @{ pkg = '@github/copilot';          bin = 'copilot'; login = 'copilot login' }
  claude  = @{ pkg = '@anthropic-ai/claude-code'; bin = 'claude';  login = 'claude  (then /login)' }
  codex   = @{ pkg = '@openai/codex';             bin = 'codex';   login = 'codex login' }
}

# Pick native Windows when node.exe is present, otherwise WSL2. Asks when both
# are available unless a mode was forced on the command line.
function Resolve-Mode([string]$Mode, [string]$Root) {
  if ($Mode -ne 'auto') { return $Mode }
  $hasNode = [bool](Get-Command node -ErrorAction SilentlyContinue)
  $hasWsl = [bool](Get-Command wsl.exe -ErrorAction SilentlyContinue)
  if ($hasNode -and $hasWsl) {
    Write-Host ""
    Write-Host "Where should the bridge run?" -ForegroundColor Yellow
    Write-Info "  1) Native Windows  (simplest: no port-forwarding, no WSL needed)"
    Write-Info "  2) WSL2"
    $choice = Read-Host "Choice [1]"
    return $(if ($choice -eq '2') { 'wsl' } else { 'native' })
  }
  if ($hasNode) { return 'native' }
  if ($hasWsl) { Write-Info "node.exe not found on Windows; using WSL2"; return 'wsl' }
  Fail "Neither Node.js on Windows nor WSL2 was found. Install Node.js 20+ (https://nodejs.org) or WSL2, then re-run."
}

function Get-WslRoot([string]$Root) {
  if (-not (Get-Command wsl.exe -ErrorAction SilentlyContinue)) { Fail "wsl.exe not found. Install WSL2 or use -Mode native." }
  $wslRoot = wsl.exe wslpath -a "$Root" 2>$null
  if ($wslRoot) { $wslRoot = $wslRoot.Trim() }
  if (-not $wslRoot) { Fail "Could not translate the repo path to a WSL path." }
  return $wslRoot
}

function Assert-Node {
  if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Fail "Node.js is not installed. Install Node.js 20+ (https://nodejs.org) and re-run."
  }
  $v = (node -v) -replace '^v', ''
  $major = [int]($v.Split('.')[0])
  if ($major -lt 20) { Fail "Node.js $v found; version 20+ is required." }
  Write-Info "node v$v"
}

function Install-Providers([string]$spec) {
  foreach ($name in ($spec -split ',')) {
    $name = $name.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($name) -or $name -eq 'none') { continue }
    if (-not $ProviderSpecs.ContainsKey($name)) { Write-Info "skipping unknown provider '$name'"; continue }
    $p = $ProviderSpecs[$name]
    if (Get-Command $p.bin -ErrorAction SilentlyContinue) {
      Write-Info "$($p.bin) already installed"
    }
    else {
      Write-Info "installing $($p.pkg) (npm -g) ..."
      npm install -g $p.pkg
      if ($LASTEXITCODE -ne 0) { Write-Info "could not install $($p.pkg); install it manually and re-run" }
    }
    Write-Info "  login when ready:  $($p.login)"
  }
}

function Ensure-Cert([string]$Root) {
  $crt = Join-Path $Root 'certs\localhost.crt'
  $key = Join-Path $Root 'certs\localhost.key'
  New-Item -ItemType Directory -Force -Path (Join-Path $Root 'certs') | Out-Null
  if ((Test-Path $crt) -and (Test-Path $key)) { Write-Info "certificate already present"; return }
  if (Get-Command mkcert -ErrorAction SilentlyContinue) {
    Write-Info "generating trusted certificate with mkcert"
    mkcert -install | Out-Null
    mkcert -cert-file $crt -key-file $key localhost 127.0.0.1 ::1
  }
  elseif (Get-Command openssl -ErrorAction SilentlyContinue) {
    Write-Info "mkcert not found; generating a self-signed certificate with openssl"
    $cnf = Join-Path $Root 'certs\openssl.cnf'
    @"
[req]
distinguished_name = dn
x509_extensions = v3
prompt = no
[dn]
CN = localhost
[v3]
subjectAltName = @alt
keyUsage = digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth
[alt]
DNS.1 = localhost
IP.1 = 127.0.0.1
IP.2 = ::1
"@ | Set-Content -Encoding ASCII $cnf
    openssl req -x509 -nodes -newkey rsa:2048 -days 825 -keyout $key -out $crt -config $cnf
  }
  else {
    Fail "Neither mkcert nor openssl was found. Install one (e.g. 'winget install FiloSottile.mkcert') and re-run."
  }
}

function Trust-Cert([string]$Root) {
  $crt = Join-Path $Root 'certs\localhost.crt'
  $thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $crt).Thumbprint
  if (Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $thumb }) {
    Write-Info "certificate already trusted"
    return
  }
  certutil -addstore -user Root $crt | Out-Null
  if ($LASTEXITCODE -ne 0) { Write-Info "certutil failed; import certs\localhost.crt into Trusted Root manually" }
  else { Write-Info "certificate trusted (CurrentUser\Root)" }
}

function Write-Config([string]$Root, [string]$ref) {
  if ([string]::IsNullOrWhiteSpace($ref)) { return }
  $path = Join-Path $Root 'excel-partner.config.json'
  $obj = [ordered]@{}
  if (Test-Path $path) {
    $existing = Get-Content -Raw $path | ConvertFrom-Json
    foreach ($p in $existing.PSObject.Properties) { $obj[$p.Name] = $p.Value }
  }
  $obj['referenceDirs'] = @($ref)
  $json = $obj | ConvertTo-Json -Depth 5
  # UTF-8 without BOM: the bridge uses JSON.parse, which rejects a leading BOM.
  [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
  Write-Info "wrote excel-partner.config.json (referenceDirs = $ref)"
}

function Stop-Bridge([string]$Root) {
  $pidFile = Join-Path $Root 'run\bridge.pid'
  if (-not (Test-Path $pidFile)) { return }
  $old = Get-Content $pidFile -ErrorAction SilentlyContinue
  if ($old -and (Get-Process -Id $old -ErrorAction SilentlyContinue)) {
    Write-Info "stopping running bridge (pid $old)"
    Stop-Process -Id $old -Force -ErrorAction SilentlyContinue
  }
  Remove-Item $pidFile -ErrorAction SilentlyContinue
}

function Start-Bridge([string]$Root) {
  $runDir = Join-Path $Root 'run'
  New-Item -ItemType Directory -Force -Path $runDir | Out-Null
  $log = Join-Path $runDir 'bridge.log'
  $err = Join-Path $runDir 'bridge.err.log'
  $proc = Start-Process -FilePath (Get-Command node).Source `
    -ArgumentList @((Join-Path $Root 'dist\bridge\src\index.js')) `
    -WorkingDirectory $Root -WindowStyle Hidden -PassThru `
    -RedirectStandardOutput $log -RedirectStandardError $err
  Set-Content -Path (Join-Path $runDir 'bridge.pid') -Value $proc.Id
  Start-Sleep -Seconds 1
  if (Get-Process -Id $proc.Id -ErrorAction SilentlyContinue) {
    Write-Info "bridge started (pid $($proc.Id)); log: $log"
  }
  else {
    Write-Info "bridge did not stay up; see $log and $err"
    Get-Content $err -Tail 20 -ErrorAction SilentlyContinue
  }
}

function Invoke-Sideload([string]$Root) {
  & (Join-Path $Root 'scripts\sideload.ps1') -ManifestPath (Join-Path $Root 'addin\manifest.xml')
}
