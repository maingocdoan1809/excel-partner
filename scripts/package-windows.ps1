[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] [ValidateSet('register', 'unregister')] [string] $Action,
  [Parameter(Mandatory = $true)] [string] $DataDir,
  [Parameter(Mandatory = $true)] [string] $ManifestPath,
  [Parameter(Mandatory = $true)] [string] $NodePath,
  [Parameter(Mandatory = $true)] [string] $CliPath,
  [switch] $NoSideload
)
$ErrorActionPreference = 'Stop'
$certPath = Join-Path $DataDir 'localhost.crt'
$installedManifest = Join-Path $DataDir 'manifest.xml'
$developerKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$runKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$manifestId = ([xml](Get-Content -LiteralPath $ManifestPath -Raw)).OfficeApp.Id
$runName = 'Excel Partner Bridge'
$fixShortcut = Join-Path ([Environment]::GetFolderPath('Programs')) 'Office Partner - Fix connection.lnk'

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
  if ($NoSideload) {
    Remove-ItemProperty -Path $developerKey -Name $manifestId -ErrorAction SilentlyContinue
  } else {
    Copy-Item -LiteralPath $ManifestPath -Destination $installedManifest -Force
    try {
      New-Item -Path $developerKey -Force | Out-Null
      New-ItemProperty -Path $developerKey -Name $manifestId -Value $installedManifest -PropertyType String -Force | Out-Null
    } catch { throw "Could not register the add-in for this user (company policy may block developer add-ins; deploy the manifest from the Microsoft 365 admin center and run setup --no-sideload): $($_.Exception.Message)" }
  }
  New-ItemProperty -Path $runKey -Name $runName -Value ('"' + $NodePath + '" "' + $CliPath + '" start --at-logon') -PropertyType String -Force | Out-Null
  # A button outside the task pane: the pane is served by the bridge, so it cannot offer one while the bridge is down.
  $shell = New-Object -ComObject WScript.Shell
  $link = $shell.CreateShortcut($fixShortcut)
  $link.TargetPath = $NodePath
  $link.Arguments = '"' + $CliPath + '" fix'
  $link.WorkingDirectory = $DataDir
  $link.Description = 'Start Office Partner again if its pane does not open or shows an error'
  $link.Save()
  Write-Host 'Registered the add-in and bridge startup for this Windows user.'
} else {
  Remove-ItemProperty -Path $developerKey -Name $manifestId -ErrorAction SilentlyContinue
  Remove-ItemProperty -Path $runKey -Name $runName -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $fixShortcut -ErrorAction SilentlyContinue
  # The Agent Gateway may use the same certificate: its trust stays while the gateway still has it.
  $gatewayCert = Join-Path $env:LOCALAPPDATA 'Agent Gateway\localhost.crt'
  $sharedWithGateway = $false
  if ((Test-Path -LiteralPath $certPath) -and (Test-Path -LiteralPath $gatewayCert)) {
    $sharedWithGateway = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $gatewayCert).Thumbprint -eq (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $certPath).Thumbprint
  }
  if ((Test-Path -LiteralPath $certPath) -and -not $sharedWithGateway) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $certPath
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')
    try {
      $store.Open('ReadWrite')
      foreach ($match in $store.Certificates.Find('FindByThumbprint', $cert.Thumbprint, $false)) { $store.Remove($match) }
    } finally { $store.Close() }
  }
  Remove-Item -LiteralPath $installedManifest -ErrorAction SilentlyContinue
  Write-Host 'Removed the add-in registration, startup entry, Fix connection shortcut and local certificate trust.'
}
