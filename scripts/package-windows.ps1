[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [ValidateSet('register', 'unregister')] [string] $Action,
  [Parameter(Mandatory = $true)] [string] $DataDir,
  [Parameter(Mandatory = $true)] [string] $ManifestPath,
  [Parameter(Mandatory = $true)] [string] $NodePath,
  [Parameter(Mandatory = $true)] [string] $CliPath
)
$ErrorActionPreference = 'Stop'
$certPath = Join-Path $DataDir 'localhost.crt'
$installedManifest = Join-Path $DataDir 'manifest.xml'
$developerKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$manifestId = ([xml](Get-Content -LiteralPath $ManifestPath -Raw)).OfficeApp.Id
$runName = 'Excel Partner Bridge'

if ($Action -eq 'register') {
  if (-not (Test-Path -LiteralPath $certPath)) { throw "Certificate missing: $certPath" }
  New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
  $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $certPath
  $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
  try {
    $rootStore.Open('ReadOnly')
    $trusted = $rootStore.Certificates.Find('FindByThumbprint', $cert.Thumbprint, $false).Count -gt 0
  } finally { $rootStore.Close() }
  if (-not $trusted) {
    & certutil.exe -addstore -user Root $certPath | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Could not trust the localhost certificate.' }
  }
  Copy-Item -LiteralPath $ManifestPath -Destination $installedManifest -Force
  New-Item -Path $developerKey -Force | Out-Null
  New-ItemProperty -Path $developerKey -Name $manifestId -Value $installedManifest -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $runKey -Name $runName -Value ('"' + $NodePath + '" "' + $CliPath + '" start') -PropertyType String -Force | Out-Null
  Write-Host 'Registered the add-in and bridge startup for this Windows user.'
} else {
  Remove-ItemProperty -Path $developerKey -Name $manifestId -ErrorAction SilentlyContinue
  Remove-ItemProperty -Path $runKey -Name $runName -ErrorAction SilentlyContinue
  if (Test-Path -LiteralPath $certPath) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $certPath
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
    try {
      $store.Open('ReadWrite')
      foreach ($match in $store.Certificates.Find('FindByThumbprint', $cert.Thumbprint, $false)) { $store.Remove($match) }
    } finally { $store.Close() }
  }
  Remove-Item -LiteralPath $installedManifest -ErrorAction SilentlyContinue
  Write-Host 'Removed the add-in registration, startup entry and local certificate trust.'
}
