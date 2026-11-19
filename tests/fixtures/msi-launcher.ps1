param([string] $Root, [string] $Work)
$ErrorActionPreference = 'Stop'
$scripts = Join-Path $Root 'installer\msi'
foreach ($file in @('bootstrap.ps1', 'stop-processes.ps1', 'purge-data.ps1')) {
  $tokens = $null; $errors = $null
  [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $Root "installer\msi\$file"), [ref]$tokens, [ref]$errors)
  if ($errors.Count) { throw ($errors -join "`n") }
}

# Every process operation is mocked. No installed product, user data, or real process is accessed.
$dataDir = Join-Path $Work 'data'
$logFile = Join-Path $Work 'stop.log'
$mockStops = New-Object System.Collections.Generic.List[int]
function Get-CimInstance {
  param($ClassName, [string] $Filter)
  if ($Filter -like 'Name=*') {
    return @(
      [pscustomobject]@{ ProcessId = 101; Name = 'node.exe'; CommandLine = 'node C:\npm\node_modules\office-partner\dist\bridge\src\index.js' },
      [pscustomobject]@{ ProcessId = 202; Name = 'node.exe'; CommandLine = 'node C:\other\server.js' }
    )
  }
  if ($Filter -eq 'ProcessId=303') { return [pscustomobject]@{ Name = 'node.exe' } }
  if ($Filter -eq 'ProcessId=404') { return [pscustomobject]@{ Name = 'other.exe' } }
  throw "Unexpected process query: $Filter"
}
function Get-NetTCPConnection {
  param($State, [int] $LocalPort)
  return [pscustomobject]@{ OwningProcess = $(if ($LocalPort -eq 8765) { 303 } else { 404 }) }
}
function Stop-Process { param([int] $Id, [switch] $Force) $mockStops.Add($Id) }
function Start-Sleep { param([int] $Milliseconds) }
. (Join-Path $scripts 'stop-processes.ps1')
Stop-ProductProcesses
if (($mockStops -join ',') -ne '101,303') { throw "Unexpected processes stopped: $($mockStops -join ',')" }
Write-Output 'PASS: embedded stop works without an installed bootstrap and leaves unrelated processes alone.'

$env:TEMP = $env:TMP = $Work
$local = Join-Path $Work 'LocalAppData'
$profileDir = Join-Path $Work 'Profile'
$gateway = Join-Path $local 'Agent Gateway'
$app = Join-Path $local 'Excel Partner'
$skills = Join-Path $profileDir '.excel-partner'
$external = Join-Path $Work 'external'
foreach ($folder in @($gateway, $app, $skills, $external)) {
  New-Item -ItemType Directory -Path $folder -Force | Out-Null
  Set-Content -LiteralPath (Join-Path $folder 'keep.txt') -Value 'example'
}
New-Item -ItemType Junction -Path (Join-Path $app 'linked') -Target $external | Out-Null
. (Join-Path $scripts 'purge-data.ps1')
Remove-ProductData -LocalAppData $local -UserProfile $profileDir
if ((Test-Path $app) -or (Test-Path $skills) -or -not (Test-Path $gateway) -or -not (Test-Path (Join-Path $external 'keep.txt'))) {
  throw 'Purge failed to remove the selected data or followed a junction.'
}
Remove-ProductData -LocalAppData $local -UserProfile $profileDir -Gateway
if (Test-Path $gateway) { throw 'Purge did not remove the selected Gateway data.' }

New-Item -ItemType Junction -Path $app -Target $external | Out-Null
try {
  Remove-ProductData -LocalAppData $local -UserProfile $profileDir
  throw 'Purge unexpectedly accepted a redirected data root.'
} catch {
  if ($_.Exception.Message -notlike '*Refusing to follow a redirected folder*') { throw }
}
if (-not (Test-Path (Join-Path $external 'keep.txt'))) { throw 'Purge deleted data from the redirected root.' }
[IO.Directory]::Delete($app)
Write-Output 'PASS: embedded purge removes selected data only.'
