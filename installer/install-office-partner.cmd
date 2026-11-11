@echo off
rem Office Partner - all-in-one installer. Double-click to install for the current user. No administrator rights needed.
rem One file: the batch lines below start PowerShell, which runs everything after the marker line.
rem Optional overrides (set before running): OP_PROVIDERS (default copilot), OP_NODE_MAJOR (default 22), OP_REGISTRY (private npm registry).
setlocal
set "OP_SELF=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:OP_SELF); $i=$t.LastIndexOf('#'+'#PS-BEGIN##'); Invoke-Expression $t.Substring($i)"
set "code=%errorlevel%"
echo.
pause
exit /b %code%
##PS-BEGIN##
# Office Partner installer for Windows (PowerShell part). Runs as the signed-in user: no administrator rights needed.
# 1. Uses the Node.js already on this PC when it is new enough, otherwise unpacks a portable Node.js 22 for this user only.
# 2. Installs the office-partner package and registers the add-in with GitHub Copilot as the provider.
$Providers = if ($env:OP_PROVIDERS) { $env:OP_PROVIDERS } else { 'copilot' }
$NodeMajor = if ($env:OP_NODE_MAJOR) { $env:OP_NODE_MAJOR } else { '22' }
$Registry = $env:OP_REGISTRY
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Corporate proxies that need the Windows sign-in.
[Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials

$dataDir = Join-Path $env:LOCALAPPDATA 'Excel Partner'
$nodeHome = Join-Path $env:LOCALAPPDATA 'Programs\nodejs'
New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
$log = Join-Path $dataDir 'install.log'
Start-Transcript -Path $log -Force | Out-Null

function Step([string] $message) { Write-Host ""; Write-Host "==> $message" -ForegroundColor Cyan }
function Fail([string] $message) { Write-Host ""; Write-Host "LOI: $message" -ForegroundColor Red; Write-Host "Nhat ky cai dat: $log"; Stop-Transcript | Out-Null; exit 1 }

# Node.js supported by the package: ^20.19.0 || >=22.12.0 <26.0.0
function Test-NodeVersion([string] $version) {
  if ($version -notmatch '^v?(\d+)\.(\d+)\.(\d+)') { return $false }
  $major = [int]$Matches[1]; $minor = [int]$Matches[2]
  return ($major -eq 20 -and $minor -ge 19) -or ($major -eq 22 -and $minor -ge 12) -or ($major -gt 22 -and $major -lt 26)
}

function Get-InstalledNode {
  $candidates = @((Get-Command node.exe -ErrorAction SilentlyContinue).Source, (Join-Path $nodeHome 'node.exe')) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
  foreach ($candidate in $candidates) {
    $version = (& $candidate --version 2>$null)
    if ($LASTEXITCODE -eq 0 -and (Test-NodeVersion $version) -and (Test-Path -LiteralPath (Join-Path (Split-Path $candidate) 'npm.cmd'))) { return $candidate }
  }
  return $null
}

function Install-PortableNode {
  Step "Tai Node.js $NodeMajor (chi cho tai khoan cua ban, khong can quyen admin)"
  $index = Invoke-RestMethod -Uri 'https://nodejs.org/dist/index.json'
  $release = $index | Where-Object { $_.version -like "v$NodeMajor.*" -and $_.lts -and ($_.files -contains 'win-x64-zip') } | Select-Object -First 1
  if (-not $release) { Fail "Khong tim thay ban Node.js $NodeMajor cho Windows tren nodejs.org." }
  $name = "node-$($release.version)-win-x64"
  $work = Join-Path $env:TEMP "office-partner-node-$([guid]::NewGuid().ToString('N'))"
  New-Item -ItemType Directory -Path $work -Force | Out-Null
  try {
    $zip = Join-Path $work "$name.zip"
    Invoke-WebRequest -Uri "https://nodejs.org/dist/$($release.version)/$name.zip" -OutFile $zip -UseBasicParsing
    $sums = (Invoke-WebRequest -Uri "https://nodejs.org/dist/$($release.version)/SHASUMS256.txt" -UseBasicParsing).Content
    $expected = ($sums -split "`n" | Where-Object { $_ -match "\s$([regex]::Escape("$name.zip"))\s*$" } | Select-Object -First 1) -replace '\s.*$', ''
    $actual = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
    if (-not $expected -or $expected.ToLower() -ne $actual.ToLower()) { Fail 'File Node.js tai ve bi sai ma kiem tra (SHA256). Khong cai dat.' }
    Expand-Archive -LiteralPath $zip -DestinationPath $work -Force
    if (Test-Path -LiteralPath $nodeHome) { Remove-Item -LiteralPath $nodeHome -Recurse -Force }
    New-Item -ItemType Directory -Path (Split-Path $nodeHome) -Force | Out-Null
    Move-Item -LiteralPath (Join-Path $work $name) -Destination $nodeHome
  } finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}

function Add-UserPath([string] $directory) {
  $current = [Environment]::GetEnvironmentVariable('Path', 'User')
  $parts = @($current -split ';' | Where-Object { $_ })
  if ($parts -notcontains $directory) { [Environment]::SetEnvironmentVariable('Path', (@($directory) + $parts) -join ';', 'User') }
  if (($env:Path -split ';') -notcontains $directory) { $env:Path = "$directory;$env:Path" }
}

Write-Host 'Office Partner - cai dat' -ForegroundColor Green
Step 'Kiem tra Word, Excel, PowerPoint'
$office = Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue
if ($office) {
  Write-Host 'Dang mo: ' ($office.ProcessName | Sort-Object -Unique) -ForegroundColor Yellow
  Write-Host 'Hay luu cong viec va dong Word, Excel, PowerPoint truoc khi cai.'
  Read-Host 'Dong xong thi nhan Enter de tiep tuc'
  if (Get-Process -Name WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue) { Fail 'Word, Excel hoac PowerPoint van dang mo. Dong chung roi chay lai bo cai.' }
}

Step 'Kiem tra Node.js'
$node = Get-InstalledNode
if ($node) { Write-Host "Dung Node.js co san: $node ($(& $node --version))" } else { Install-PortableNode; $node = Join-Path $nodeHome 'node.exe' }
$nodeDir = Split-Path $node
Add-UserPath $nodeDir
$npm = Join-Path $nodeDir 'npm.cmd'
Write-Host "Node.js $(& $node --version), npm $(& $npm --version)"

Step 'Cai dat Office Partner'
$npmArgs = @('install', '-g', 'office-partner')
if ($Registry) { $npmArgs += @('--registry', $Registry) }
& $npm @npmArgs
if ($LASTEXITCODE -ne 0) { Fail "npm install khong thanh cong (ma $LASTEXITCODE). Kiem tra mang hoac proxy cong ty." }
$prefix = (& $npm prefix -g).Trim()
Add-UserPath $prefix
$command = Join-Path $prefix 'office-partner.cmd'
if (-not (Test-Path -LiteralPath $command)) { Fail "Khong tim thay lenh office-partner tai $prefix." }

Step 'Dang ky add-in va khoi dong bridge'
& $command setup --providers $Providers
if ($LASTEXITCODE -ne 0) { Fail "office-partner setup khong thanh cong (ma $LASTEXITCODE)." }
& $command status

Write-Host ""
Write-Host 'Cai dat xong.' -ForegroundColor Green
Write-Host '1. Mo Excel, Word hoac PowerPoint, vao Home > Office Partner.'
Write-Host '2. Lan dau: vao Settings trong khung Office Partner va dang nhap GitHub Copilot.'
Write-Host "Neu gap loi, chay: office-partner doctor   (nhat ky cai dat: $log)"
Stop-Transcript | Out-Null
